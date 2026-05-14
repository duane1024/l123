//! `/File Import Parquet` (v0.4) — read a parquet file into a
//! [`LoadedRecords`] via the parquet+arrow row API.
//!
//! Type mapping (PLAN §M11):
//!   bool   → 1 / 0           (Value::Number)
//!   int*   → Value::Number   (any signed/unsigned width, cast to f64)
//!   float* → Value::Number
//!   utf8   → Value::Text
//!   null   → Value::Empty
//!   date/timestamp → Value::Number raw integer (D1 format-tag plumbing
//!     is a follow-up; today the user sees the raw days-since-epoch /
//!     nanos value and can apply a date format manually)
//!   nested types → stringified via Debug (v0.4 punt — same idea as the
//!     JSON loader's nested-object handling)

use std::fs::File;
use std::path::Path;
use std::sync::Arc;

use arrow_array::{
    Array, BooleanArray, Float32Array, Float64Array, Int16Array, Int32Array, Int64Array,
    Int8Array, RecordBatch, StringArray, UInt16Array, UInt32Array, UInt64Array, UInt8Array,
};
use arrow_schema::{DataType, Schema};
use l123_core::Value;
use parquet::arrow::arrow_reader::ParquetRecordBatchReaderBuilder;

use crate::records::{LoadError, LoadedRecords};

pub fn load(path: &Path) -> Result<LoadedRecords, LoadError> {
    let file = File::open(path).map_err(|e| LoadError::Io(e.to_string()))?;
    let builder = ParquetRecordBatchReaderBuilder::try_new(file)
        .map_err(|e| LoadError::Parse {
            location: "parquet header".into(),
            message: e.to_string(),
        })?;
    let schema = builder.schema().clone();
    let header: Vec<String> = schema.fields().iter().map(|f| f.name().clone()).collect();
    let reader = builder.build().map_err(|e| LoadError::Parse {
        location: "parquet reader".into(),
        message: e.to_string(),
    })?;

    let mut rows: Vec<Vec<Value>> = Vec::new();
    for batch in reader {
        let batch = batch.map_err(|e| LoadError::Parse {
            location: "row batch".into(),
            message: e.to_string(),
        })?;
        append_batch_rows(&batch, &schema, &mut rows)?;
    }
    Ok(LoadedRecords { header, rows })
}

fn append_batch_rows(
    batch: &RecordBatch,
    schema: &Arc<Schema>,
    out: &mut Vec<Vec<Value>>,
) -> Result<(), LoadError> {
    let num_rows = batch.num_rows();
    let num_cols = batch.num_columns();
    for r in 0..num_rows {
        let mut row = Vec::with_capacity(num_cols);
        for c in 0..num_cols {
            let array = batch.column(c);
            let field = schema.field(c);
            row.push(extract_value(array, field.data_type(), r)?);
        }
        out.push(row);
    }
    Ok(())
}

fn extract_value(
    array: &dyn Array,
    data_type: &DataType,
    row: usize,
) -> Result<Value, LoadError> {
    if array.is_null(row) {
        return Ok(Value::Empty);
    }
    match data_type {
        DataType::Boolean => {
            let a = array.as_any().downcast_ref::<BooleanArray>().unwrap();
            Ok(Value::Number(if a.value(row) { 1.0 } else { 0.0 }))
        }
        DataType::Int8 => Ok(Value::Number(downcast::<Int8Array>(array).value(row) as f64)),
        DataType::Int16 => Ok(Value::Number(downcast::<Int16Array>(array).value(row) as f64)),
        DataType::Int32 => Ok(Value::Number(downcast::<Int32Array>(array).value(row) as f64)),
        DataType::Int64 => Ok(Value::Number(downcast::<Int64Array>(array).value(row) as f64)),
        DataType::UInt8 => Ok(Value::Number(downcast::<UInt8Array>(array).value(row) as f64)),
        DataType::UInt16 => Ok(Value::Number(
            downcast::<UInt16Array>(array).value(row) as f64
        )),
        DataType::UInt32 => Ok(Value::Number(
            downcast::<UInt32Array>(array).value(row) as f64
        )),
        DataType::UInt64 => Ok(Value::Number(
            downcast::<UInt64Array>(array).value(row) as f64
        )),
        DataType::Float32 => Ok(Value::Number(
            downcast::<Float32Array>(array).value(row) as f64
        )),
        DataType::Float64 => Ok(Value::Number(downcast::<Float64Array>(array).value(row))),
        DataType::Utf8 => Ok(Value::Text(
            downcast::<StringArray>(array).value(row).to_string(),
        )),
        // Dates / timestamps surface as their raw integer value for
        // v0.4. The (D1) format-tag plumbing is a follow-up; for now
        // the user can apply a date format manually via /RFD1.
        DataType::Date32 => Ok(Value::Number(downcast::<Int32Array>(array).value(row) as f64)),
        DataType::Date64 => Ok(Value::Number(downcast::<Int64Array>(array).value(row) as f64)),
        // Anything we don't have explicit support for falls through
        // to a debug-stringified Text cell so the user at least
        // sees the data.
        _ => Ok(Value::Text(format!("{array:?}"))),
    }
}

fn downcast<T: 'static>(array: &dyn Array) -> &T {
    array
        .as_any()
        .downcast_ref::<T>()
        .expect("data_type and array kind agree (checked by arrow)")
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn fixture() -> PathBuf {
        // The acceptance fixture lives at repo-root; cargo test runs
        // from the crate dir, so step up two levels.
        let mut p = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        p.push("../../tests/acceptance/fixtures/m11_import.parquet");
        p
    }

    #[test]
    fn header_in_schema_order() {
        let r = load(&fixture()).expect("fixture loads");
        assert_eq!(r.header, vec!["id", "name", "qty", "rate", "active"]);
    }

    #[test]
    fn three_rows_with_mixed_types() {
        let r = load(&fixture()).expect("fixture loads");
        assert_eq!(r.rows.len(), 3);
        // Row 1: 1, widget, 100, 0.5, true(=1)
        assert_eq!(r.rows[0][0], Value::Number(1.0));
        assert_eq!(r.rows[0][1], Value::Text("widget".into()));
        assert_eq!(r.rows[0][2], Value::Number(100.0));
        assert_eq!(r.rows[0][3], Value::Number(0.5));
        assert_eq!(r.rows[0][4], Value::Number(1.0));
    }

    #[test]
    fn bool_false_widens_to_zero() {
        let r = load(&fixture()).expect("fixture loads");
        assert_eq!(r.rows[1][4], Value::Number(0.0));
    }

    #[test]
    fn null_int_becomes_empty() {
        let r = load(&fixture()).expect("fixture loads");
        // Row 3 has NULL in qty.
        assert_eq!(r.rows[2][2], Value::Empty);
    }

    #[test]
    fn missing_file_is_io_error() {
        let p = PathBuf::from("/no/such/path.parquet");
        match load(&p) {
            Err(LoadError::Io(_)) => {}
            other => panic!("expected Io error, got {other:?}"),
        }
    }
}
