//! Engine abstraction. IronCalc is the primary implementation.

pub mod engine;
pub mod ironcalc_adapter;
pub mod num_fmt;

pub use engine::{CellView, Engine, EngineError, RecalcMode};
pub use ironcalc_adapter::IronCalcEngine;
