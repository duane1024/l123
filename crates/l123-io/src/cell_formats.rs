//! Cell-format-extras sidecar inside .xlsx files.
//!
//! Excel's `num_fmt` system carries the standard 1-2-3 numeric kinds
//! (Fixed/Sci/Currency/Comma/Percent/Date/Time) and survives a save
//! cleanly, but it has no encoding for:
//!
//! * **L123-specific kinds** — `(+)` PlusMinus, `(T)` Text-shows-formula,
//!   `(H)` Hidden, `(A)` Automatic, `(L)` Label-only.
//! * **Per-format flags** — `parens` (the `/Range Format Other
//!   Parentheses Yes` toggle), `negative_color` (per-format negative-
//!   value font color).
//!
//! These extras live in this sidecar (`l123/cell_formats.tsv` inside
//! the xlsx zip). On save we emit one row per cell that carries any
//! extra plus a `[global]` row for the workbook-wide default's extras.
//! On load we layer the sidecar on top of whatever the engine pulled
//! from `num_fmt`, so a save → reload round-trip is exact.
//!
//! Vanilla Excel ignores unknown parts; round-tripping through Excel
//! drops the sidecar — that's the explicit limitation, mirroring
//! `formula_sources` (see SPEC §14 / PLAN §5.1).
//!
//! ## Format
//!
//! ```text
//! # l123 cell format extras v1
//! g\t<kind|->\t<0|1>\t<RRGGBB|->
//! <sheet>\t<col>\t<row>\t<kind|->\t<0|1>\t<RRGGBB|->
//! ```
//!
//! All addresses 0-based (matching `formula_sources`). `kind` tokens:
//! `plus_minus`, `text`, `hidden`, `automatic`, `label_only`. `-` in
//! the kind column means "use whatever the engine loaded from
//! num_fmt". The `g` line is optional; per-cell rows are optional;
//! a sidecar with no extras is omitted entirely.

use std::collections::HashMap;
use std::fs::File;
use std::io::{self, Read, Write};
use std::path::Path;

use l123_core::{Address, Format, FormatKind, RgbColor, SheetId};

const SIDECAR_PATH: &str = "l123/cell_formats.tsv";
const SIDECAR_HEADER: &str = "# l123 cell format extras v1\n";

/// Bundle of everything the sidecar carries: the workbook-wide
/// default's extras (when any), plus per-cell overlays.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellFormatExtras {
    pub global: Option<FormatExtras>,
    pub cells: HashMap<Address, FormatExtras>,
}

/// The fields of a `Format` that don't round-trip through Excel's
/// `num_fmt` cleanly. Construct via [`FormatExtras::from_format`];
/// apply via [`FormatExtras::apply_to`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct FormatExtras {
    /// `Some` when the kind is one Excel can't represent — overrides
    /// the engine-loaded kind on reload. `None` means "leave kind
    /// alone, just apply the flag overrides below."
    pub kind_override: Option<FormatKind>,
    pub parens: bool,
    pub negative_color: Option<RgbColor>,
}

impl FormatExtras {
    /// Extract the sidecar-worthy fields from a [`Format`]. Returns
    /// `None` when the format has nothing extra (Excel can round-trip
    /// it whole).
    pub fn from_format(f: Format) -> Option<Self> {
        let kind_override = match f.kind {
            FormatKind::PlusMinus
            | FormatKind::Text
            | FormatKind::Hidden
            | FormatKind::Automatic
            | FormatKind::LabelOnly => Some(f.kind),
            _ => None,
        };
        if kind_override.is_none() && !f.parens && f.negative_color.is_none() {
            return None;
        }
        Some(Self {
            kind_override,
            parens: f.parens,
            negative_color: f.negative_color,
        })
    }

    /// Layer this overlay on top of `f`. Replaces `kind` when the
    /// overlay carries one; always replaces `parens` and
    /// `negative_color`.
    pub fn apply_to(self, mut f: Format) -> Format {
        if let Some(kind) = self.kind_override {
            f.kind = kind;
        }
        f.parens = self.parens;
        f.negative_color = self.negative_color;
        f
    }
}

/// Read the cell-format-extras sidecar from an xlsx file. Returns
/// empty extras when the file has no sidecar (vanilla Excel xlsx, or
/// an L123 file saved before this feature).
pub fn read_from_xlsx(path: &Path) -> io::Result<CellFormatExtras> {
    let f = File::open(path)?;
    let mut zip = match zip::ZipArchive::new(f) {
        Ok(z) => z,
        Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
    };
    let mut entry = match zip.by_name(SIDECAR_PATH) {
        Ok(e) => e,
        Err(zip::result::ZipError::FileNotFound) => return Ok(CellFormatExtras::default()),
        Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
    };
    let mut text = String::new();
    entry.read_to_string(&mut text)?;
    Ok(parse_tsv(&text))
}

/// Write the cell-format-extras sidecar into an xlsx file in place.
/// If the bundle is empty (no global, no per-cell entries), the
/// existing sidecar entry is removed and no new one written —
/// keeps unmodified xlsx files free of empty noise.
pub fn write_to_xlsx(path: &Path, extras: &CellFormatExtras) -> io::Result<()> {
    let tmp_path = path.with_extension("xlsx.l123tmp");
    let payload = serialize_tsv(extras);
    let omit_sidecar = extras.global.is_none() && extras.cells.is_empty();
    {
        let in_file = File::open(path)?;
        let mut in_zip = match zip::ZipArchive::new(in_file) {
            Ok(z) => z,
            Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
        };
        let out_file = File::create(&tmp_path)?;
        let mut out_zip = zip::ZipWriter::new(out_file);
        for i in 0..in_zip.len() {
            let entry = match in_zip.by_index_raw(i) {
                Ok(e) => e,
                Err(e) => return Err(io::Error::new(io::ErrorKind::InvalidData, e)),
            };
            if entry.name() == SIDECAR_PATH {
                continue;
            }
            if let Err(e) = out_zip.raw_copy_file(entry) {
                return Err(io::Error::other(e));
            }
        }
        if !omit_sidecar {
            let opts = zip::write::FileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            if let Err(e) = out_zip.start_file(SIDECAR_PATH, opts) {
                return Err(io::Error::other(e));
            }
            out_zip.write_all(payload.as_bytes())?;
        }
        if let Err(e) = out_zip.finish() {
            return Err(io::Error::other(e));
        }
    }
    std::fs::rename(&tmp_path, path)?;
    Ok(())
}

fn kind_token(k: FormatKind) -> Option<&'static str> {
    match k {
        FormatKind::PlusMinus => Some("plus_minus"),
        FormatKind::Text => Some("text"),
        FormatKind::Hidden => Some("hidden"),
        FormatKind::Automatic => Some("automatic"),
        FormatKind::LabelOnly => Some("label_only"),
        _ => None,
    }
}

fn kind_from_token(s: &str) -> Option<FormatKind> {
    match s {
        "plus_minus" => Some(FormatKind::PlusMinus),
        "text" => Some(FormatKind::Text),
        "hidden" => Some(FormatKind::Hidden),
        "automatic" => Some(FormatKind::Automatic),
        "label_only" => Some(FormatKind::LabelOnly),
        _ => None,
    }
}

fn extras_to_fields(e: FormatExtras) -> (String, String, String) {
    let kind = e
        .kind_override
        .and_then(kind_token)
        .map(|s| s.to_string())
        .unwrap_or_else(|| "-".to_string());
    let parens = if e.parens {
        "1".to_string()
    } else {
        "0".to_string()
    };
    let neg = match e.negative_color {
        Some(rgb) => rgb.to_rgb_hex(),
        None => "-".to_string(),
    };
    (kind, parens, neg)
}

fn fields_to_extras(kind: &str, parens: &str, neg: &str) -> Option<FormatExtras> {
    let kind_override = if kind == "-" {
        None
    } else {
        Some(kind_from_token(kind)?)
    };
    let parens = match parens {
        "0" => false,
        "1" => true,
        _ => return None,
    };
    let negative_color = if neg == "-" {
        None
    } else {
        Some(RgbColor::from_hex(neg)?)
    };
    Some(FormatExtras {
        kind_override,
        parens,
        negative_color,
    })
}

fn serialize_tsv(extras: &CellFormatExtras) -> String {
    let mut out = String::with_capacity(SIDECAR_HEADER.len() + 64);
    out.push_str(SIDECAR_HEADER);
    if let Some(g) = extras.global {
        let (k, p, n) = extras_to_fields(g);
        out.push_str(&format!("g\t{k}\t{p}\t{n}\n"));
    }
    let mut entries: Vec<(&Address, &FormatExtras)> = extras.cells.iter().collect();
    // Stable order for clean diffs across saves.
    entries.sort_by_key(|(a, _)| (a.sheet.0, a.row, a.col));
    for (addr, fe) in entries {
        let (k, p, n) = extras_to_fields(*fe);
        out.push_str(&format!(
            "{}\t{}\t{}\t{}\t{}\t{}\n",
            addr.sheet.0, addr.col, addr.row, k, p, n
        ));
    }
    out
}

fn parse_tsv(text: &str) -> CellFormatExtras {
    let mut out = CellFormatExtras::default();
    for line in text.lines() {
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let parts: Vec<&str> = line.split('\t').collect();
        // `g` line: 4 fields (`g`, kind, parens, neg).
        if parts.len() == 4 && parts[0] == "g" {
            if let Some(fe) = fields_to_extras(parts[1], parts[2], parts[3]) {
                out.global = Some(fe);
            }
            continue;
        }
        // Cell line: 6 fields.
        if parts.len() == 6 {
            let sheet = parts[0].parse::<u16>().ok();
            let col = parts[1].parse::<u16>().ok();
            let row = parts[2].parse::<u32>().ok();
            let extras = fields_to_extras(parts[3], parts[4], parts[5]);
            if let (Some(sheet), Some(col), Some(row), Some(fe)) = (sheet, col, row, extras) {
                out.cells.insert(
                    Address {
                        sheet: SheetId(sheet),
                        col,
                        row,
                    },
                    fe,
                );
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fe(k: Option<FormatKind>, parens: bool, neg: Option<RgbColor>) -> FormatExtras {
        FormatExtras {
            kind_override: k,
            parens,
            negative_color: neg,
        }
    }

    #[test]
    fn from_format_skips_when_nothing_extra() {
        assert_eq!(FormatExtras::from_format(Format::GENERAL), None);
        assert_eq!(FormatExtras::from_format(Format::fixed(2)), None);
        assert_eq!(FormatExtras::from_format(Format::currency(0)), None);
    }

    #[test]
    fn from_format_captures_non_excel_kinds() {
        let f = Format::from_kind(FormatKind::PlusMinus);
        assert_eq!(
            FormatExtras::from_format(f),
            Some(fe(Some(FormatKind::PlusMinus), false, None))
        );
        let f = Format::from_kind(FormatKind::LabelOnly);
        assert_eq!(
            FormatExtras::from_format(f),
            Some(fe(Some(FormatKind::LabelOnly), false, None))
        );
    }

    #[test]
    fn from_format_captures_parens_on_excel_kind() {
        let f = Format::currency(2).with_parens();
        assert_eq!(
            FormatExtras::from_format(f),
            Some(fe(None, true, None)),
            "Currency 2 + parens: only the parens flag is sidecar-worthy"
        );
    }

    #[test]
    fn from_format_captures_negative_color() {
        let mut f = Format::currency(0);
        f.negative_color = Some(RgbColor { r: 255, g: 0, b: 0 });
        assert_eq!(
            FormatExtras::from_format(f),
            Some(fe(None, false, Some(RgbColor { r: 255, g: 0, b: 0 })))
        );
    }

    #[test]
    fn apply_to_replaces_kind_and_flags() {
        let base = Format::currency(2);
        let overlay = fe(Some(FormatKind::PlusMinus), true, Some(RgbColor::BLACK));
        let after = overlay.apply_to(base);
        assert_eq!(after.kind, FormatKind::PlusMinus);
        assert!(after.parens);
        assert_eq!(after.negative_color, Some(RgbColor::BLACK));
        assert_eq!(after.decimals, 2, "decimals must survive the overlay");
    }

    #[test]
    fn apply_to_keeps_kind_when_overlay_omits_it() {
        let base = Format::currency(0);
        let overlay = fe(None, true, None);
        let after = overlay.apply_to(base);
        assert_eq!(after.kind, FormatKind::Currency);
        assert!(after.parens);
    }

    #[test]
    fn serialize_then_parse_round_trips() {
        let mut cells = HashMap::new();
        cells.insert(
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 0,
            },
            fe(Some(FormatKind::PlusMinus), false, None),
        );
        cells.insert(
            Address {
                sheet: SheetId(0),
                col: 1,
                row: 5,
            },
            fe(
                None,
                true,
                Some(RgbColor {
                    r: 255,
                    g: 128,
                    b: 0,
                }),
            ),
        );
        let extras = CellFormatExtras {
            global: Some(fe(None, true, None)),
            cells,
        };
        let tsv = serialize_tsv(&extras);
        assert!(tsv.starts_with("# l123 cell format extras v1"));
        let parsed = parse_tsv(&tsv);
        assert_eq!(parsed, extras);
    }

    #[test]
    fn parse_skips_garbled_lines() {
        let tsv = "# l123 cell format extras v1\n\
                   not enough fields\n\
                   g\tplus_minus\t0\t-\n\
                   0\t0\t5\t-\t1\tFF0000\n\
                   abc\tdef\tghi\tjkl\tmno\tpqr\n";
        let parsed = parse_tsv(tsv);
        assert_eq!(
            parsed.global,
            Some(fe(Some(FormatKind::PlusMinus), false, None))
        );
        assert_eq!(parsed.cells.len(), 1);
        assert_eq!(
            parsed.cells.get(&Address {
                sheet: SheetId(0),
                col: 0,
                row: 5,
            }),
            Some(&fe(None, true, Some(RgbColor { r: 255, g: 0, b: 0 })))
        );
    }

    #[test]
    fn missing_sidecar_returns_empty_extras() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");
        {
            let f = File::create(&path).unwrap();
            let mut zw = zip::ZipWriter::new(f);
            zw.start_file("hello.txt", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"world").unwrap();
            zw.finish().unwrap();
        }
        let got = read_from_xlsx(&path).unwrap();
        assert_eq!(got, CellFormatExtras::default());
    }

    #[test]
    fn write_then_read_round_trips_through_zip() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");
        {
            let f = File::create(&path).unwrap();
            let mut zw = zip::ZipWriter::new(f);
            zw.start_file("xl/workbook.xml", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"<workbook/>").unwrap();
            zw.finish().unwrap();
        }
        let mut cells = HashMap::new();
        cells.insert(
            Address {
                sheet: SheetId(0),
                col: 2,
                row: 3,
            },
            fe(Some(FormatKind::Text), false, None),
        );
        let extras = CellFormatExtras {
            global: None,
            cells,
        };
        write_to_xlsx(&path, &extras).unwrap();
        let got = read_from_xlsx(&path).unwrap();
        assert_eq!(got, extras);
    }

    #[test]
    fn empty_extras_omits_sidecar_entry() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("test.xlsx");
        {
            let f = File::create(&path).unwrap();
            let mut zw = zip::ZipWriter::new(f);
            zw.start_file("xl/workbook.xml", zip::write::FileOptions::default())
                .unwrap();
            zw.write_all(b"<workbook/>").unwrap();
            zw.finish().unwrap();
        }
        // First write extras, then write empty: should remove the
        // sidecar from the zip on the empty write.
        let mut cells = HashMap::new();
        cells.insert(
            Address {
                sheet: SheetId(0),
                col: 0,
                row: 0,
            },
            fe(Some(FormatKind::Text), false, None),
        );
        let extras = CellFormatExtras {
            global: None,
            cells,
        };
        write_to_xlsx(&path, &extras).unwrap();
        write_to_xlsx(&path, &CellFormatExtras::default()).unwrap();
        let f = File::open(&path).unwrap();
        let mut zip = zip::ZipArchive::new(f).unwrap();
        assert!(zip.by_name(SIDECAR_PATH).is_err());
    }
}
