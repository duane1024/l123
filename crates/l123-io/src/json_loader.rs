//! `/File Import Json` (v0.4) — read an array-of-objects or a
//! JSON-Lines file into a [`LoadedRecords`].
//!
//! Format detection looks at the first non-whitespace byte:
//!   `[`  → strict-JSON array of objects
//!   `{`  → JSON-Lines (one object per line)
//!
//! Type widening (per PLAN §M11):
//!   bool   → 1 / 0          (Value::Number)
//!   null   → empty cell     (Value::Empty)
//!   number → Value::Number
//!   string → Value::Text
//!   arrays / nested objects → stringified (`serde_json::to_string`)
//!     and surfaced as Value::Text — a v0.4 punt; users who need
//!     nested data should flatten upstream.
//!
//! Headers come from object keys in the order seen across all rows
//! (the `preserve_order` feature on serde_json keeps each object's
//! key order intact; row scanning unions across rows).

use std::path::Path;

use l123_core::Value;
use serde_json::Value as JsonValue;

use crate::records::{LoadError, LoadedRecords};

/// Read `path` and decode it as either array-of-objects JSON or
/// JSON-Lines, returning a header + rows. Format is auto-detected by
/// the first non-whitespace byte.
pub fn load(path: &Path) -> Result<LoadedRecords, LoadError> {
    let body = std::fs::read_to_string(path).map_err(|e| LoadError::Io(e.to_string()))?;
    parse(&body)
}

/// Same as [`load`] but takes the JSON text directly. Exposed so unit
/// tests don't need a tempfile.
pub fn parse(body: &str) -> Result<LoadedRecords, LoadError> {
    let first = body
        .chars()
        .find(|c| !c.is_whitespace())
        .ok_or_else(|| LoadError::Parse {
            location: "file start".into(),
            message: "input is empty or whitespace-only".into(),
        })?;
    match first {
        '[' => parse_array(body),
        '{' => parse_lines(body),
        other => Err(LoadError::UnsupportedShape(format!(
            "expected `[` (array) or `{{` (JSON-Lines), got `{other}`"
        ))),
    }
}

fn parse_array(body: &str) -> Result<LoadedRecords, LoadError> {
    let parsed: JsonValue = serde_json::from_str(body).map_err(json_err)?;
    let JsonValue::Array(items) = parsed else {
        return Err(LoadError::UnsupportedShape(
            "top-level JSON must be an array of objects".into(),
        ));
    };
    let mut headers: Vec<String> = Vec::new();
    let mut row_maps: Vec<&serde_json::Map<String, JsonValue>> = Vec::new();
    for (i, item) in items.iter().enumerate() {
        let JsonValue::Object(map) = item else {
            return Err(LoadError::UnsupportedShape(format!(
                "array element {i} is not an object"
            )));
        };
        for k in map.keys() {
            if !headers.iter().any(|h| h == k) {
                headers.push(k.clone());
            }
        }
        row_maps.push(map);
    }
    let rows = row_maps
        .iter()
        .map(|map| project_row(map, &headers))
        .collect();
    Ok(LoadedRecords::new(headers, rows))
}

fn parse_lines(body: &str) -> Result<LoadedRecords, LoadError> {
    let mut headers: Vec<String> = Vec::new();
    let mut row_maps: Vec<serde_json::Map<String, JsonValue>> = Vec::new();
    for (i, line) in body.lines().enumerate() {
        if line.trim().is_empty() {
            continue;
        }
        let parsed: JsonValue = serde_json::from_str(line).map_err(|e| LoadError::Parse {
            location: format!("line {}", i + 1),
            message: e.to_string(),
        })?;
        let JsonValue::Object(map) = parsed else {
            return Err(LoadError::UnsupportedShape(format!(
                "line {} is not a JSON object",
                i + 1
            )));
        };
        for k in map.keys() {
            if !headers.iter().any(|h| h == k) {
                headers.push(k.clone());
            }
        }
        row_maps.push(map);
    }
    let rows = row_maps
        .iter()
        .map(|map| project_row(map, &headers))
        .collect();
    Ok(LoadedRecords::new(headers, rows))
}

fn project_row(map: &serde_json::Map<String, JsonValue>, headers: &[String]) -> Vec<Value> {
    headers
        .iter()
        .map(|h| match map.get(h) {
            Some(v) => json_to_value(v),
            None => Value::Empty,
        })
        .collect()
}

fn json_to_value(v: &JsonValue) -> Value {
    match v {
        JsonValue::Null => Value::Empty,
        JsonValue::Bool(true) => Value::Number(1.0),
        JsonValue::Bool(false) => Value::Number(0.0),
        JsonValue::Number(n) => n
            .as_f64()
            .map(Value::Number)
            .unwrap_or_else(|| Value::Text(n.to_string())),
        JsonValue::String(s) => Value::Text(s.clone()),
        JsonValue::Array(_) | JsonValue::Object(_) => {
            Value::Text(serde_json::to_string(v).unwrap_or_default())
        }
    }
}

fn json_err(e: serde_json::Error) -> LoadError {
    LoadError::Parse {
        location: format!("line {}, column {}", e.line(), e.column()),
        message: e.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn array_of_objects_preserves_key_order() {
        let body = r#"[{"id": 1, "name": "widget", "qty": 100}]"#;
        let got = parse(body).unwrap();
        assert_eq!(got.header, vec!["id", "name", "qty"]);
        assert_eq!(
            got.rows,
            vec![vec![
                Value::Number(1.0),
                Value::Text("widget".into()),
                Value::Number(100.0),
            ]]
        );
    }

    #[test]
    fn bool_widens_to_one_or_zero() {
        let body = r#"[{"a": true, "b": false}]"#;
        let got = parse(body).unwrap();
        assert_eq!(got.rows[0], vec![Value::Number(1.0), Value::Number(0.0)]);
    }

    #[test]
    fn null_becomes_empty() {
        let body = r#"[{"id": 1, "qty": null}]"#;
        let got = parse(body).unwrap();
        assert_eq!(got.rows[0], vec![Value::Number(1.0), Value::Empty]);
    }

    #[test]
    fn jsonl_one_object_per_line() {
        let body = "{\"id\": 1, \"name\": \"widget\"}\n\
                    {\"id\": 2, \"name\": \"gadget\"}\n";
        let got = parse(body).unwrap();
        assert_eq!(got.header, vec!["id", "name"]);
        assert_eq!(got.rows.len(), 2);
        assert_eq!(got.rows[0][0], Value::Number(1.0));
        assert_eq!(got.rows[1][1], Value::Text("gadget".into()));
    }

    #[test]
    fn jsonl_blank_lines_ignored() {
        let body = "{\"id\": 1}\n\n\n{\"id\": 2}\n";
        let got = parse(body).unwrap();
        assert_eq!(got.rows.len(), 2);
    }

    #[test]
    fn missing_key_in_a_row_yields_empty() {
        let body = r#"[{"id": 1, "name": "a"}, {"id": 2}]"#;
        let got = parse(body).unwrap();
        assert_eq!(got.header, vec!["id", "name"]);
        assert_eq!(got.rows[1], vec![Value::Number(2.0), Value::Empty]);
    }

    #[test]
    fn array_with_extra_key_in_second_row_extends_header() {
        // Header is the union of keys in encounter order across rows.
        let body = r#"[{"id": 1}, {"id": 2, "extra": "x"}]"#;
        let got = parse(body).unwrap();
        assert_eq!(got.header, vec!["id", "extra"]);
        assert_eq!(got.rows[0], vec![Value::Number(1.0), Value::Empty]);
        assert_eq!(
            got.rows[1],
            vec![Value::Number(2.0), Value::Text("x".into())]
        );
    }

    #[test]
    fn nested_object_stringified_as_text() {
        let body = r#"[{"id": 1, "tags": {"a": 1, "b": 2}}]"#;
        let got = parse(body).unwrap();
        let Value::Text(t) = &got.rows[0][1] else {
            panic!("expected Text, got {:?}", got.rows[0][1]);
        };
        assert!(t.contains("\"a\""));
    }

    #[test]
    fn malformed_array_returns_parse_error() {
        let body = "[{\"id\": 1,";
        let err = parse(body).unwrap_err();
        assert!(matches!(err, LoadError::Parse { .. }), "got {err:?}");
    }

    #[test]
    fn malformed_jsonl_returns_parse_error_with_line() {
        let body = "{\"id\": 1}\n{\"id\":";
        let err = parse(body).unwrap_err();
        let LoadError::Parse { location, .. } = err else {
            panic!("expected Parse");
        };
        assert!(location.contains("line 2"), "got {location:?}");
    }

    #[test]
    fn empty_input_is_an_error() {
        assert!(parse("   ").is_err());
    }

    #[test]
    fn non_array_non_object_top_level_is_unsupported() {
        assert!(matches!(
            parse("42"),
            Err(LoadError::UnsupportedShape(_))
        ));
    }
}
