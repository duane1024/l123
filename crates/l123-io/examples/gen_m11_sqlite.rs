//! Fixture generator for `/File Import Sqlite` acceptance.
//!
//! Run via `cargo run -p l123-io --example gen_m11_sqlite`. Writes a
//! tiny SQLite database with two tables to
//! `tests/acceptance/fixtures/m11_import.sqlite`:
//!
//! `items`     (id INT, name TEXT, qty INT NULL, rate REAL)
//! `suppliers` (id INT, name TEXT)
//!
//! The two-table shape exercises the NAMES-mode picker that the
//! `/File Import Sqlite` flow opens after the user picks the file.

use std::path::Path;

use rusqlite::{params, Connection};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let out = Path::new("tests/acceptance/fixtures/m11_import.sqlite");
    if out.exists() {
        std::fs::remove_file(out)?;
    }
    let conn = Connection::open(out)?;

    conn.execute_batch(
        "CREATE TABLE items (id INTEGER, name TEXT, qty INTEGER, rate REAL);
         CREATE TABLE suppliers (id INTEGER, name TEXT);",
    )?;

    let rows: &[(i64, &str, Option<i64>, f64)] = &[
        (1, "widget", Some(100), 0.5),
        (2, "gadget", Some(42), 1.25),
        (3, "gizmo", None, 3.0),
    ];
    for r in rows {
        conn.execute(
            "INSERT INTO items (id, name, qty, rate) VALUES (?1, ?2, ?3, ?4)",
            params![r.0, r.1, r.2, r.3],
        )?;
    }

    for r in &[(101, "acme"), (102, "globex")] {
        conn.execute(
            "INSERT INTO suppliers (id, name) VALUES (?1, ?2)",
            params![r.0, r.1],
        )?;
    }

    println!(
        "Wrote {} ({} bytes)",
        out.display(),
        std::fs::metadata(out)?.len()
    );
    Ok(())
}
