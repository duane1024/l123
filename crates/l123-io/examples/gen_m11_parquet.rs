//! One-shot fixture generator for `/File Import Parquet` acceptance.
//!
//! Run via `cargo run -p l123-io --example gen_m11_parquet` from the
//! workspace root. Writes a tiny typed parquet file to
//! `tests/acceptance/fixtures/m11_import.parquet`. The file is
//! checked in as a binary fixture; this program documents how it was
//! produced and lets us regenerate after a parquet/arrow bump.
//!
//! Schema (3 rows):
//!   id (Int32):    1, 2, 3
//!   name (Utf8):   widget, gadget, gizmo
//!   qty (Int64):   100, 42, NULL
//!   rate (F64):    0.5, 1.25, 3.0
//!   active (Bool): true, false, true

use std::fs::File;
use std::path::Path;
use std::sync::Arc;

use arrow_array::{
    ArrayRef, BooleanArray, Float64Array, Int32Array, Int64Array, RecordBatch, StringArray,
};
use arrow_schema::{DataType, Field, Schema};
use parquet::arrow::ArrowWriter;
use parquet::file::properties::WriterProperties;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let schema = Arc::new(Schema::new(vec![
        Field::new("id", DataType::Int32, false),
        Field::new("name", DataType::Utf8, false),
        Field::new("qty", DataType::Int64, true),
        Field::new("rate", DataType::Float64, false),
        Field::new("active", DataType::Boolean, false),
    ]));

    let id: ArrayRef = Arc::new(Int32Array::from(vec![1, 2, 3]));
    let name: ArrayRef = Arc::new(StringArray::from(vec!["widget", "gadget", "gizmo"]));
    let qty: ArrayRef = Arc::new(Int64Array::from(vec![Some(100), Some(42), None]));
    let rate: ArrayRef = Arc::new(Float64Array::from(vec![0.5, 1.25, 3.0]));
    let active: ArrayRef = Arc::new(BooleanArray::from(vec![true, false, true]));

    let batch = RecordBatch::try_new(schema.clone(), vec![id, name, qty, rate, active])?;

    let out = Path::new("tests/acceptance/fixtures/m11_import.parquet");
    let file = File::create(out)?;
    let props = WriterProperties::builder().build();
    let mut writer = ArrowWriter::try_new(file, schema, Some(props))?;
    writer.write(&batch)?;
    writer.close()?;

    println!("Wrote {} ({} bytes)", out.display(), std::fs::metadata(out)?.len());
    Ok(())
}
