//! File adapters: .xlsx, .csv, .wk3 (stretch — Lotus 1-2-3 R3 legacy format).
//!
//! xlsx I/O is driven through the `Engine` trait in `l123-engine`; this
//! crate owns the formats that sit outside the IronCalc backend — CSV
//! today, and the legacy `.wk3` reader at stretch-tier.

pub mod csv;
