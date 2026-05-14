//! `/File Import Parquet` (v0.4) — read a parquet file into a
//! [`LoadedRecords`] via the parquet+arrow row API.
//!
//! Type mapping (PLAN §M11):
//!   bool   → 1 / 0           (Value::Number)
//!   int*   → Value::Number   (any signed/unsigned width, cast to f64)
//!   float* → Value::Number
//!   utf8   → Value::Text
//!   null   → Value::Empty
//!   Date32 / Date64 / Timestamp(*) → Value::Number Excel serial
//!     (days since 1899-12-30) AND `column_formats[col] = Some((D1))`
//!     so cells render as `DD-MMM-YY` instead of raw integers.
//!   nested types → stringified via Debug (v0.4 punt — same idea as the
//!     JSON loader's nested-object handling)

use std::fs::File;
use std::path::Path;
use std::sync::Arc;

use arrow_array::{
    Array, BooleanArray, Date32Array, Date64Array, Float32Array, Float64Array, Int16Array,
    Int32Array, Int64Array, Int8Array, RecordBatch, StringArray, TimestampMicrosecondArray,
    TimestampMillisecondArray, TimestampNanosecondArray, TimestampSecondArray, UInt16Array,
    UInt32Array, UInt64Array, UInt8Array,
};
use arrow_schema::{DataType, Schema, TimeUnit};
use l123_core::{Format, FormatKind, Value};
use parquet::arrow::arrow_reader::ParquetRecordBatchReaderBuilder;

use crate::records::{LoadError, LoadedRecords};

/// Days between 1899-12-30 (Excel epoch) and 1970-01-01 (Unix epoch).
/// Add this to a Date32 (days since 1970-01-01) to get a 1-2-3 /
/// Excel date serial.
const EPOCH_OFFSET_DAYS: i64 = 25_569;
const SECS_PER_DAY: i64 = 86_400;
const MILLIS_PER_DAY: i64 = SECS_PER_DAY * 1_000;
const MICROS_PER_DAY: i64 = MILLIS_PER_DAY * 1_000;
const NANOS_PER_DAY: i64 = MICROS_PER_DAY * 1_000;

pub fn load(path: &Path) -> Result<LoadedRecords, LoadError> {
    let file = File::open(path).map_err(|e| LoadError::Io(e.to_string()))?;
    let builder = ParquetRecordBatchReaderBuilder::try_new(file)
        .map_err(|e| LoadError::Parse {
            location: "parquet header".into(),
            message: e.to_string(),
        })?;
    let schema = builder.schema().clone();
    let header: Vec<String> = schema.fields().iter().map(|f| f.name().clone()).collect();
    let column_formats: Vec<Option<Format>> = schema
        .fields()
        .iter()
        .map(|f| format_for_data_type(f.data_type()))
        .collect();
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
    Ok(LoadedRecords {
        header,
        rows,
        column_formats,
    })
}

/// Per-column format hint. Date / timestamp columns get the 1-2-3
/// `(D1)` format so the Excel serial we emit renders as `DD-MMM-YY`.
fn format_for_data_type(t: &DataType) -> Option<Format> {
    match t {
        DataType::Date32
        | DataType::Date64
        | DataType::Timestamp(_, _) => Some(Format::from_kind(FormatKind::DateDmy)),
        _ => None,
    }
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
        // Date / timestamp columns are converted to 1-2-3 / Excel
        // serial (days since 1899-12-30) here; the loader also tags
        // the column with `(D1)` so the renderer formats it as
        // `DD-MMM-YY`.
        DataType::Date32 => {
            let days = downcast::<Date32Array>(array).value(row) as i64;
            Ok(Value::Number((days + EPOCH_OFFSET_DAYS) as f64))
        }
        DataType::Date64 => {
            let ms = downcast::<Date64Array>(array).value(row);
            let days = div_floor(ms, MILLIS_PER_DAY);
            Ok(Value::Number((days + EPOCH_OFFSET_DAYS) as f64))
        }
        DataType::Timestamp(unit, _tz) => {
            let raw = match unit {
                TimeUnit::Second => downcast::<TimestampSecondArray>(array).value(row),
                TimeUnit::Millisecond => downcast::<TimestampMillisecondArray>(array).value(row),
                TimeUnit::Microsecond => downcast::<TimestampMicrosecondArray>(array).value(row),
                TimeUnit::Nanosecond => downcast::<TimestampNanosecondArray>(array).value(row),
            };
            let divisor = match unit {
                TimeUnit::Second => SECS_PER_DAY,
                TimeUnit::Millisecond => MILLIS_PER_DAY,
                TimeUnit::Microsecond => MICROS_PER_DAY,
                TimeUnit::Nanosecond => NANOS_PER_DAY,
            };
            let days = div_floor(raw, divisor);
            Ok(Value::Number((days + EPOCH_OFFSET_DAYS) as f64))
        }
        // Anything we don't have explicit support for falls through
        // to a debug-stringified Text cell so the user at least
        // sees the data.
        _ => Ok(Value::Text(format!("{array:?}"))),
    }
}

/// Euclidean division — `(-1).div_floor(86_400)` is `-1`, not `0`, so
/// dates before the Unix epoch round to the correct earlier day. Pre-
/// 1970 timestamps shouldn't show up in modern parquet data but the
/// arithmetic stays honest at the boundary.
fn div_floor(n: i64, d: i64) -> i64 {
    let q = n / d;
    let r = n % d;
    if (r != 0) && ((r < 0) != (d < 0)) {
        q - 1
    } else {
        q
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
        assert_eq!(
            r.header,
            vec!["id", "name", "qty", "rate", "active", "purchased_on"]
        );
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
    fn date32_converts_to_excel_serial() {
        let r = load(&fixture()).expect("fixture loads");
        // 2024-01-01 fixture: Date32 19723 -> serial 45292.
        assert_eq!(r.rows[0][5], Value::Number(45_292.0));
        // 2024-07-15: Date32 19919 -> serial 45488.
        assert_eq!(r.rows[1][5], Value::Number(45_488.0));
        // NULL stays empty.
        assert_eq!(r.rows[2][5], Value::Empty);
    }

    #[test]
    fn date32_column_tagged_with_d1_format() {
        let r = load(&fixture()).expect("fixture loads");
        let purchased_on = r.column_formats[5].as_ref().expect("date column tagged");
        assert_eq!(purchased_on.kind, FormatKind::DateDmy);
        // Non-date columns stay untagged.
        assert!(r.column_formats[0].is_none());
        assert!(r.column_formats[3].is_none());
    }

    #[test]
    fn missing_file_is_io_error() {
        let p = PathBuf::from("/no/such/path.parquet");
        match load(&p) {
            Err(LoadError::Io(_)) => {}
            other => panic!("expected Io error, got {other:?}"),
        }
    }

    #[test]
    fn div_floor_handles_negative_inputs() {
        assert_eq!(div_floor(0, 86_400), 0);
        assert_eq!(div_floor(86_400, 86_400), 1);
        assert_eq!(div_floor(86_399, 86_400), 0);
        assert_eq!(div_floor(-1, 86_400), -1);
        assert_eq!(div_floor(-86_400, 86_400), -1);
        assert_eq!(div_floor(-86_401, 86_400), -2);
    }
}
