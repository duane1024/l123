//! Cell value representation.

#[derive(Clone, Debug, PartialEq)]
pub enum Value {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrKind),
}

/// 1-2-3 error kinds. 1-2-3 only surfaces ERR and NA to users but we keep the
/// underlying Excel-style kinds (from IronCalc) for diagnostics / F1 help.
#[derive(Copy, Clone, Debug, PartialEq, Eq)]
pub enum ErrKind {
    /// Generic 1-2-3 "ERR" result.
    Err,
    /// `@NA` — not available.
    Na,
    /// Division by zero. Displays as `ERR` in 1-2-3 mode.
    DivZero,
    /// Reference to a deleted cell. Displays as `ERR`.
    Ref,
    /// Undefined name. Displays as `ERR`.
    Name,
    /// Wrong-type argument. Displays as `ERR`.
    Value,
    /// Out-of-domain numeric result. Displays as `ERR`.
    Num,
    /// Circular reference. Displays as `ERR` with CIRC indicator.
    Circular,
}

impl ErrKind {
    /// Short Lotus-style tag ("ERR" or "NA") for on-screen display.
    pub fn lotus_tag(self) -> &'static str {
        match self {
            ErrKind::Na => "NA",
            _ => "ERR",
        }
    }

    /// Excel-style code, for F1 disclosure and .xlsx round-trip.
    pub fn excel_code(self) -> &'static str {
        match self {
            ErrKind::Err => "#ERR!",
            ErrKind::Na => "#N/A",
            ErrKind::DivZero => "#DIV/0!",
            ErrKind::Ref => "#REF!",
            ErrKind::Name => "#NAME?",
            ErrKind::Value => "#VALUE!",
            ErrKind::Num => "#NUM!",
            ErrKind::Circular => "#CIRC!",
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn err_tags() {
        assert_eq!(ErrKind::Err.lotus_tag(), "ERR");
        assert_eq!(ErrKind::Na.lotus_tag(), "NA");
        assert_eq!(ErrKind::DivZero.lotus_tag(), "ERR");
        assert_eq!(ErrKind::DivZero.excel_code(), "#DIV/0!");
    }
}
