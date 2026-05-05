//! §4.7 long-running ops: queue, dispatch, and drain on tokio.
//!
//! The App-side methods (`tick`, `spawn_async_op`, `apply_async_result`,
//! …) live here alongside the `worker_*` free functions that run on the
//! `spawn_blocking` pool. Anything else under `app::` reaches the
//! workers only through `App::queue_async_op`.

use std::collections::HashMap;
use std::path::Path;
use std::sync::atomic::Ordering;

use tokio::sync::oneshot;

use l123_core::{Address, CellContents, LabelPrefix, Mode, SheetId, Value};
use l123_engine::{Engine, IronCalcEngine};

use super::types::{AsyncProgress, AsyncResult, OpState, PendingAsyncOp, QueuedOp};
use super::{is_wk3_path, App};

impl App {
    /// §4.7 acceptance hook — lowers (or raises) the F9-recalc cell-
    /// count threshold so a small transcript can exercise the WAIT
    /// path. No-op outside tests; production code never calls this.
    pub fn test_set_recalc_wait_threshold(&mut self, n: usize) {
        self.recalc_wait_cell_threshold = n;
    }

    /// §4.7 acceptance hook — write synthetic progress numbers to
    /// the currently-pending async op's shared state so a transcript
    /// can render-and-assert the `[████░░] N%` bar without having
    /// to time the worker. No-op when nothing is queued.
    pub fn test_seed_async_progress(&mut self, done: u64, total: u64) {
        if let Some(op) = self.pending_async_op.as_ref() {
            op.progress.done.store(done, Ordering::Relaxed);
            op.progress.total.store(total, Ordering::Relaxed);
        }
    }

    /// §4.7 acceptance hook: park the next async op in `Queued`
    /// state so a transcript can observe mid-flight WAIT mode and
    /// pre-spawn cancellation. Sticky until `test_resume_async_op`
    /// or Ctrl-Break clears it.
    pub fn test_block_next_async_op(&mut self) {
        self.block_next_async_op = true;
    }

    /// §4.7 acceptance hook: clear the block and run the parked op
    /// to completion synchronously, applying the result before
    /// returning. Equivalent to ticking until the queue drains, but
    /// uses the runtime's `block_on` so xlsx loads / saves don't
    /// busy-loop.
    pub fn test_resume_async_op(&mut self) {
        self.block_next_async_op = false;
        self.tick_inner(true);
    }

    /// Acceptance-harness companion to `tick`: blocks until the
    /// worker returns instead of polling non-blocking. Respects
    /// `block_next_async_op` so a transcript that called
    /// `BLOCK_NEXT_OP` keeps the op parked. Mirrors the
    /// pre-§4.7 sync drain semantics so existing transcripts that
    /// don't care about WAIT mid-flight stay green without
    /// sprinkling `RESUME_OP` everywhere.
    pub fn drain_async_op_blocking(&mut self) {
        self.tick_inner(true);
    }

    /// §4.7 — drain a queued long-running op. The production event
    /// loop calls this each iteration after rendering. Non-blocking:
    /// if the worker is still running the op stays in WAIT and the
    /// next render frame sees updated progress. Gated by
    /// `block_next_async_op` so tests can hold the queue at the
    /// pre-spawn boundary indefinitely.
    pub fn tick(&mut self) {
        self.tick_inner(false);
    }

    /// Shared body of `tick` and `test_resume_async_op`. With
    /// `wait == false` we poll the worker non-blocking; with
    /// `wait == true` we block on the oneshot until the worker
    /// returns.
    pub(super) fn tick_inner(&mut self, wait: bool) {
        if self.block_next_async_op {
            return;
        }
        let Some(op) = self.pending_async_op.take() else {
            return;
        };
        let PendingAsyncOp {
            verb,
            display_name,
            progress,
            state,
        } = op;
        let mut rx = match state {
            OpState::Queued(queued) => self.spawn_async_op(*queued, progress.clone()),
            OpState::Running(rx) => rx,
        };
        let result = if wait {
            self.runtime.block_on(&mut rx).ok()
        } else {
            match rx.try_recv() {
                Ok(r) => Some(r),
                Err(oneshot::error::TryRecvError::Empty) => None,
                Err(oneshot::error::TryRecvError::Closed) => Some(AsyncResult::Errored {
                    engine: None,
                    message: "worker task dropped without result".to_string(),
                }),
            }
        };
        match result {
            None => {
                self.pending_async_op = Some(PendingAsyncOp {
                    verb,
                    display_name,
                    progress,
                    state: OpState::Running(rx),
                });
            }
            Some(res) => {
                self.apply_async_result(res);
                if matches!(self.mode, Mode::Wait) {
                    self.mode = Mode::Ready;
                }
            }
        }
    }

    /// Spawn the worker for a queued op on the tokio blocking pool;
    /// returns the oneshot receiver the main thread polls in `tick`.
    fn spawn_async_op(
        &self,
        queued: QueuedOp,
        progress: AsyncProgress,
    ) -> oneshot::Receiver<AsyncResult> {
        let (tx, rx) = oneshot::channel();
        match queued {
            QueuedOp::FileRetrieve { path } => {
                self.runtime.spawn_blocking(move || {
                    let _ = tx.send(worker_file_retrieve(path, progress));
                });
            }
            QueuedOp::FileSave {
                engine,
                path,
                formula_sources,
                cell_format_extras,
            } => {
                self.runtime.spawn_blocking(move || {
                    let _ = tx.send(worker_file_save(
                        engine,
                        path,
                        formula_sources,
                        cell_format_extras,
                        progress,
                    ));
                });
            }
            QueuedOp::FileImportNumbers {
                engine,
                path,
                origin,
            } => {
                self.runtime.spawn_blocking(move || {
                    let _ = tx.send(worker_file_import(
                        engine, path, origin, progress, /* numeric_split = */ true,
                    ));
                });
            }
            QueuedOp::FileImportText {
                engine,
                path,
                origin,
            } => {
                self.runtime.spawn_blocking(move || {
                    let _ = tx.send(worker_file_import(
                        engine, path, origin, progress, /* numeric_split = */ false,
                    ));
                });
            }
            QueuedOp::Recalc { engine } => {
                self.runtime.spawn_blocking(move || {
                    let _ = tx.send(worker_recalc(engine, progress));
                });
            }
        }
        rx
    }

    /// Apply a worker's result to App state. Mode-flip back to READY
    /// happens in `tick_inner` once this returns.
    fn apply_async_result(&mut self, res: AsyncResult) {
        match res {
            AsyncResult::FileRetrieveXlsx {
                engine,
                path,
                is_wk3,
            } => {
                self.wb_mut().engine = engine;
                self.repopulate_after_xlsx_load(path, is_wk3);
                self.try_autoexec();
            }
            AsyncResult::FileRetrieveCsv {
                engine,
                path,
                cells,
            } => {
                // Mirror `execute_file_new`'s wipe (sans the engine
                // swap, which we do explicitly below) so a /FR onto
                // a dirty workbook lands at A:A1 with no leftover
                // styles or pending modal state.
                self.wb_mut().cells.clear();
                self.wb_mut().clear_all_cell_formats();
                self.wb_mut().cell_text_styles.clear();
                self.wb_mut().cell_alignments.clear();
                self.wb_mut().cell_fills.clear();
                self.wb_mut().cell_font_styles.clear();
                self.wb_mut().cell_borders.clear();
                self.wb_mut().comments.clear();
                self.wb_mut().merges.clear();
                self.wb_mut().frozen.clear();
                self.wb_mut().sheet_states.clear();
                self.wb_mut().tables.clear();
                self.wb_mut().sheet_colors.clear();
                self.wb_mut().col_widths.clear();
                self.wb_mut().default_col_width = 9;
                self.wb_mut().hidden_cols.clear();
                self.entry = None;
                self.wb_mut().pointer = Address::A1;
                self.wb_mut().viewport_col_offset = 0;
                self.wb_mut().viewport_row_offset = 0;
                self.recalc_pending = false;
                self.wb_mut().engine = engine;
                for (a, c) in cells {
                    self.wb_mut().cells.insert(a, c);
                }
                self.refresh_formula_caches();
                self.wb_mut().active_path = Some(path);
                self.wb_mut().dirty = false;
                self.try_autoexec();
            }
            AsyncResult::FileSave {
                engine,
                path,
                result,
            } => {
                self.wb_mut().engine = engine;
                if result.is_ok() {
                    self.wb_mut().active_path = Some(path);
                    self.wb_mut().dirty = false;
                } else if let Err(msg) = result {
                    self.set_error(format!("Cannot save: {msg}"));
                }
            }
            AsyncResult::FileImport { engine, cells } => {
                self.wb_mut().engine = engine;
                for (a, c) in cells {
                    self.wb_mut().cells.insert(a, c);
                }
                self.refresh_formula_caches();
            }
            AsyncResult::Recalc { engine } => {
                self.wb_mut().engine = engine;
                self.refresh_formula_caches();
                self.recalc_pending = false;
            }
            AsyncResult::Cancelled { engine } => {
                if let Some(e) = engine {
                    self.wb_mut().engine = e;
                }
            }
            AsyncResult::Errored { engine, message } => {
                if let Some(e) = engine {
                    self.wb_mut().engine = e;
                }
                self.set_error(message);
            }
        }
    }

    /// Queue a long-running op and flip into WAIT mode. The next
    /// `tick()` spawns the worker (unless `block_next_async_op` is
    /// set, which holds it at the pre-spawn boundary for tests).
    pub(super) fn queue_async_op(
        &mut self,
        verb: &'static str,
        display_name: String,
        queued: QueuedOp,
    ) {
        self.pending_async_op = Some(PendingAsyncOp {
            verb,
            display_name,
            progress: AsyncProgress::default(),
            state: OpState::Queued(Box::new(queued)),
        });
        self.mode = Mode::Wait;
    }

    /// SPEC §7 / §4.7: Ctrl-Break aborts an in-flight long op. For
    /// a Queued op this is instant — the op never spawns. For a
    /// Running op we set the cancel flag and synchronously wait for
    /// the worker to return so the main thread can put back any
    /// engine that was moved out (preventing a stranded placeholder
    /// engine in the workbook). Returns true when an op was cancelled.
    pub(super) fn cancel_pending_async_op(&mut self) -> bool {
        let Some(op) = self.pending_async_op.take() else {
            return false;
        };
        self.block_next_async_op = false;
        match op.state {
            OpState::Queued(_) => {
                // Worker never started — nothing to wait on. Engine
                // was never moved out either, so workbook is intact.
            }
            OpState::Running(mut rx) => {
                op.progress.cancel.store(true, Ordering::SeqCst);
                if let Ok(res) = self.runtime.block_on(&mut rx) {
                    // Apply just the engine-restore portion — drop
                    // any cells/cache deltas the worker had already
                    // computed. Cancel is "leaves no partial state."
                    self.restore_engine_only(res);
                }
            }
        }
        self.mode = Mode::Ready;
        true
    }

    /// Cancel-path companion to `apply_async_result`: put back the
    /// engine if the worker returned one, but discard any cells or
    /// success metadata so the workbook looks like nothing happened.
    fn restore_engine_only(&mut self, res: AsyncResult) {
        let engine = match res {
            AsyncResult::FileRetrieveXlsx { .. } | AsyncResult::FileRetrieveCsv { .. } => None,
            AsyncResult::FileSave { engine, .. }
            | AsyncResult::FileImport { engine, .. }
            | AsyncResult::Recalc { engine } => Some(engine),
            AsyncResult::Cancelled { engine } | AsyncResult::Errored { engine, .. } => engine,
        };
        if let Some(e) = engine {
            self.wb_mut().engine = e;
        }
    }
}

/// Convert a CSV field to (UI cache contents, engine input string).
/// Numeric tokens become `Constant(Number)` with a general-format
/// engine string; everything else becomes an apostrophe-prefixed
/// label. Used by both the worker and the sync CSV paths.
pub(super) fn csv_field_to_cell(field: &str) -> (CellContents, String) {
    match field.parse::<f64>() {
        Ok(n) => (
            CellContents::Constant(Value::Number(n)),
            l123_core::format_number_general(n),
        ),
        Err(_) => (
            CellContents::Label {
                prefix: LabelPrefix::Apostrophe,
                text: field.to_string(),
            },
            format!("'{field}"),
        ),
    }
}

/// Read a file into a String, ticking `progress.done` as bytes
/// arrive and bailing on `progress.cancel`. `progress.total` is
/// pre-set to the file size so the renderer can draw a real
/// `[████░░] N%` bar.
fn read_file_with_progress(
    path: &Path,
    progress: &AsyncProgress,
) -> std::result::Result<String, String> {
    use std::io::Read;
    let f =
        std::fs::File::open(path).map_err(|e| format!("Cannot read {}: {e}", path.display()))?;
    let total = f.metadata().map(|m| m.len()).unwrap_or(0);
    progress.done.store(0, Ordering::Relaxed);
    progress.total.store(total, Ordering::Relaxed);
    let mut reader = std::io::BufReader::new(f);
    let mut out = Vec::with_capacity(total as usize);
    let mut chunk = [0u8; 8 * 1024];
    loop {
        if progress.cancel.load(Ordering::Relaxed) {
            return Err("cancelled".to_string());
        }
        let n = reader
            .read(&mut chunk)
            .map_err(|e| format!("read error on {}: {e}", path.display()))?;
        if n == 0 {
            break;
        }
        out.extend_from_slice(&chunk[..n]);
        progress.done.fetch_add(n as u64, Ordering::Relaxed);
    }
    String::from_utf8(out).map_err(|e| format!("invalid UTF-8 in {}: {e}", path.display()))
}

/// §4.7 worker — `/File Retrieve`. Dispatches by extension. Builds
/// a fresh engine on the worker so the caller's existing engine is
/// untouched until the result is applied (cancellation = no-op
/// against the workbook).
fn worker_file_retrieve(path: std::path::PathBuf, progress: AsyncProgress) -> AsyncResult {
    if progress.cancel.load(Ordering::Relaxed) {
        return AsyncResult::Cancelled { engine: None };
    }
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|s| s.to_ascii_lowercase());
    if ext.as_deref() == Some("csv") {
        worker_csv_retrieve(path, progress)
    } else {
        worker_xlsx_retrieve(path, progress)
    }
}

fn worker_xlsx_retrieve(path: std::path::PathBuf, progress: AsyncProgress) -> AsyncResult {
    let is_wk3 = is_wk3_path(&path);
    let mut engine = match IronCalcEngine::new() {
        Ok(e) => e,
        Err(err) => {
            return AsyncResult::Errored {
                engine: None,
                message: err.to_string(),
            }
        }
    };
    let load_result = if is_wk3 {
        #[cfg(feature = "wk3")]
        {
            engine.load_wk3(&path)
        }
        #[cfg(not(feature = "wk3"))]
        {
            engine.load_xlsx(&path)
        }
    } else {
        engine.load_xlsx(&path)
    };
    if let Err(e) = load_result {
        return AsyncResult::Errored {
            engine: None,
            message: format!("Cannot open {}: {e}", path.display()),
        };
    }
    if progress.cancel.load(Ordering::Relaxed) {
        return AsyncResult::Cancelled { engine: None };
    }
    AsyncResult::FileRetrieveXlsx {
        engine,
        path,
        is_wk3,
    }
}

fn worker_csv_retrieve(path: std::path::PathBuf, progress: AsyncProgress) -> AsyncResult {
    let body = match read_file_with_progress(&path, &progress) {
        Ok(b) => b,
        Err(e) => {
            return if e == "cancelled" {
                AsyncResult::Cancelled { engine: None }
            } else {
                AsyncResult::Errored {
                    engine: None,
                    message: e,
                }
            }
        }
    };
    let mut engine = match IronCalcEngine::new() {
        Ok(e) => e,
        Err(err) => {
            return AsyncResult::Errored {
                engine: None,
                message: err.to_string(),
            }
        }
    };
    let rows = l123_io::csv::parse(&body);
    let total_rows = rows.len() as u64;
    progress.done.store(0, Ordering::Relaxed);
    progress.total.store(total_rows.max(1), Ordering::Relaxed);
    let mut cells = Vec::new();
    for (dr, row) in rows.iter().enumerate() {
        if progress.cancel.load(Ordering::Relaxed) {
            return AsyncResult::Cancelled { engine: None };
        }
        for (dc, field) in row.iter().enumerate() {
            if field.is_empty() {
                continue;
            }
            let addr = Address::new(SheetId(0), dc as u16, dr as u32);
            let (contents, engine_input) = csv_field_to_cell(field);
            let _ = engine.set_user_input(addr, &engine_input);
            cells.push((addr, contents));
        }
        progress.done.store((dr as u64) + 1, Ordering::Relaxed);
    }
    engine.recalc();
    AsyncResult::FileRetrieveCsv {
        engine,
        path,
        cells,
    }
}

/// §4.7 worker — `/File Save`. The engine has been moved out of
/// `Workbook` and travels with the op; the worker writes the xlsx
/// (and the formula-source sidecar on success), then ships the
/// engine back. Cancel before the save is honored; once
/// `save_xlsx` is in flight the call is opaque.
fn worker_file_save(
    engine: IronCalcEngine,
    path: std::path::PathBuf,
    formula_sources: HashMap<Address, String>,
    cell_format_extras: l123_io::cell_formats::CellFormatExtras,
    progress: AsyncProgress,
) -> AsyncResult {
    if progress.cancel.load(Ordering::Relaxed) {
        return AsyncResult::Cancelled {
            engine: Some(engine),
        };
    }
    if let Some(parent) = path.parent() {
        if !parent.as_os_str().is_empty() {
            let _ = std::fs::create_dir_all(parent);
        }
    }
    if let Err(e) = engine.save_xlsx(&path) {
        return AsyncResult::FileSave {
            engine,
            path,
            result: Err(e.to_string()),
        };
    }
    let _ = l123_io::formula_sources::write_to_xlsx(&path, &formula_sources);
    let _ = l123_io::cell_formats::write_to_xlsx(&path, &cell_format_extras);
    AsyncResult::FileSave {
        engine,
        path,
        result: Ok(()),
    }
}

/// §4.7 worker — `/File Import {Numbers,Text}`. The engine travels
/// with the op so set_user_input on N rows runs off-thread. With
/// `numeric_split = true`, each line is CSV-split and numeric
/// fields become numbers; with `false`, each line becomes one
/// apostrophe-prefixed label down a single column (matching the
/// existing sync `/FIT` behavior, which doesn't split on commas).
fn worker_file_import(
    mut engine: IronCalcEngine,
    path: std::path::PathBuf,
    origin: Address,
    progress: AsyncProgress,
    numeric_split: bool,
) -> AsyncResult {
    let body = match read_file_with_progress(&path, &progress) {
        Ok(b) => b,
        Err(e) => {
            return if e == "cancelled" {
                AsyncResult::Cancelled {
                    engine: Some(engine),
                }
            } else {
                AsyncResult::Errored {
                    engine: Some(engine),
                    message: e,
                }
            }
        }
    };
    let mut cells = Vec::new();
    if numeric_split {
        let rows = l123_io::csv::parse(&body);
        let total_rows = rows.len() as u64;
        progress.done.store(0, Ordering::Relaxed);
        progress.total.store(total_rows.max(1), Ordering::Relaxed);
        for (dr, row) in rows.iter().enumerate() {
            if progress.cancel.load(Ordering::Relaxed) {
                return AsyncResult::Cancelled {
                    engine: Some(engine),
                };
            }
            for (dc, field) in row.iter().enumerate() {
                if field.is_empty() {
                    continue;
                }
                let addr =
                    Address::new(origin.sheet, origin.col + dc as u16, origin.row + dr as u32);
                let (contents, engine_input) = csv_field_to_cell(field);
                let _ = engine.set_user_input(addr, &engine_input);
                cells.push((addr, contents));
            }
            progress.done.store((dr as u64) + 1, Ordering::Relaxed);
        }
    } else {
        let lines: Vec<&str> = body.lines().collect();
        let total_rows = lines.len() as u64;
        progress.done.store(0, Ordering::Relaxed);
        progress.total.store(total_rows.max(1), Ordering::Relaxed);
        for (dr, line) in lines.iter().enumerate() {
            if progress.cancel.load(Ordering::Relaxed) {
                return AsyncResult::Cancelled {
                    engine: Some(engine),
                };
            }
            if !line.is_empty() {
                let addr = Address::new(origin.sheet, origin.col, origin.row + dr as u32);
                let engine_input = format!("'{line}");
                let _ = engine.set_user_input(addr, &engine_input);
                cells.push((
                    addr,
                    CellContents::Label {
                        prefix: LabelPrefix::Apostrophe,
                        text: (*line).to_string(),
                    },
                ));
            }
            progress.done.store((dr as u64) + 1, Ordering::Relaxed);
        }
    }
    engine.recalc();
    AsyncResult::FileImport { engine, cells }
}

/// §4.7 worker — F9 recalc on a workbook above `super::RECALC_WAIT_CELL_THRESHOLD`.
/// IronCalc's `recalc` is opaque so progress stays indeterminate
/// (renderer falls back to the verb-only line). Cancel is honored
/// before the call but not during.
fn worker_recalc(mut engine: IronCalcEngine, progress: AsyncProgress) -> AsyncResult {
    if progress.cancel.load(Ordering::Relaxed) {
        return AsyncResult::Cancelled {
            engine: Some(engine),
        };
    }
    engine.recalc();
    AsyncResult::Recalc { engine }
}
