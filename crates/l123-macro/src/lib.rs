//! L123 macro language: lexer + interpreter for `/X` and `{...}`
//! commands plus the keystroke playback substrate.
//!
//! The macro language is at heart "type these keys for me." A small
//! and growing set of `{TOKEN}` and `{DIRECTIVE ...}` forms layer
//! control flow, variables, and prompts on top of the keystroke
//! stream. SPEC §18 puts macros in the Complete tier; PLAN.md M9 is
//! where they live.
//!
//! This crate is layered below `l123-ui` so it cannot speak in
//! crossterm `KeyEvent`s; instead it emits its own [`MacroKey`]
//! enum which the UI translates at the dispatch boundary.

#![cfg_attr(not(test), forbid(unsafe_code))]

use std::fmt;

/// A single keystroke produced by lexing a macro source string.
///
/// The variants correspond 1:1 to the keystrokes the UI dispatcher
/// already understands. The macro lexer's job is to map between
/// 1-2-3's source forms (`~`, `{DOWN}`, `{TILDE}`, `{GOTO}`, ...)
/// and these tokens.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MacroKey {
    /// A literal character to type, exactly as the user would.
    Char(char),
    /// Enter / commit. Spelled `~` in macro source.
    Enter,
    Up,
    Down,
    Left,
    Right,
    Home,
    End,
    PageUp,
    PageDown,
    /// Ctrl+Left — `{BIGLEFT}` in macros.
    BigLeft,
    /// Ctrl+Right — `{BIGRIGHT}` in macros.
    BigRight,
    Escape,
    Backspace,
    Delete,
    Insert,
    Tab,
    /// Function key. F1..F10 are valid; the lexer only emits values
    /// in that range.
    Function(u8),
}

/// One element in the action stream produced by [`lex_actions`].
///
/// Macros are line-oriented (one cell = one line). Each cell's
/// source lexes into a sequence of [`MacroAction`]s that the
/// interpreter executes in order. Some directives end execution of
/// the current line (e.g. an [`If(false)`](MacroAction::If) skip);
/// others alter the program counter (e.g.
/// [`Branch`](MacroAction::Branch)).
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MacroAction {
    /// A keystroke. The interpreter feeds this through the UI
    /// dispatcher exactly as if the user had pressed the key.
    Key(MacroKey),
    /// `{BRANCH loc}` — set the program counter to `loc`.
    Branch(String),
    /// `{QUIT}` — abort the entire macro.
    Quit,
    /// `{RETURN}` — pop one frame off the call stack. At the
    /// outermost frame this is equivalent to [`Quit`](MacroAction::Quit).
    Return,
    /// `{IF expr}` — evaluate `expr`; if zero / false, skip the
    /// remaining actions on the current line.
    If(String),
    /// Subroutine call `{loc arg1,arg2,...}` — push a return frame
    /// and set the PC to `loc`.
    Subroutine {
        loc: String,
        args: Vec<String>,
    },
    /// `{DEFINE arg-loc[: type]...}` — declare a subroutine's
    /// positional argument cells. Stub: the interpreter currently
    /// ignores the arg-binding side; recognized so it doesn't
    /// trigger an "unknown directive" failure.
    Define(Vec<String>),
    /// `{LET loc, expr}` — write `expr`'s value to `loc`. Numbers,
    /// strings, and formulas all parse through the same source
    /// pipeline as user-typed cell entries. The optional `:string` /
    /// `:value` suffix on `expr` forces a coercion (Lotus syntax);
    /// for now we drop it and let the source parser figure out the
    /// shape.
    Let {
        loc: String,
        expr: String,
    },
    /// `{BLANK range}` — erase the cells at `range` (no journal
    /// note: we re-use the existing /Range Erase plumbing on the UI
    /// side so this is undoable).
    Blank(String),
    /// `{RECALC [range]}` — force a recalc. Optional range argument
    /// is currently treated as "the whole workbook" — IronCalc's
    /// recalc is workbook-wide, partial-range recalc isn't exposed
    /// through the engine surface yet.
    Recalc(String),
    /// `{?}` — pause the macro and let the user interact freely.
    /// Any keystrokes until the user presses Enter go to the active
    /// mode normally; Enter resumes the macro.
    QuestionPause,
    /// `{GETLABEL prompt-text, loc}` — pause, show `prompt-text` in
    /// the control panel, accept user input, store the result at
    /// `loc` as a label.
    GetLabel {
        prompt_text: String,
        loc: String,
    },
    /// `{GETNUMBER prompt-text, loc}` — same as `{GETLABEL}` but the
    /// commit is parsed as a number.
    GetNumber {
        prompt_text: String,
        loc: String,
    },
    /// `{MENUBRANCH loc}` — open a custom menu whose definition
    /// lives at `loc` (item names in row 0, descriptions in row 1,
    /// macro bodies starting at row 2). On selection the
    /// interpreter branches to the chosen item's body.
    MenuBranch(String),
    /// `{MENUCALL loc}` — same as [`MenuBranch`] but the chosen
    /// body runs as a subroutine (`{RETURN}` returns to the caller).
    MenuCall(String),
    /// `{BEEP [n]}` — emit a tone (pitch ignored). Bumps the
    /// app-level beep counter so transcripts can assert on it.
    Beep,
    /// `{WAIT serial}` — pause until a time-serial. Stub: the
    /// interpreter recognizes it but doesn't actually sleep, since
    /// driving wall-clock time through transcripts isn't useful.
    Wait(String),
    /// `{BREAKOFF}` / `{BREAKON}` — disable / enable Ctrl-Break
    /// macro abort. Stubs until we wire Ctrl-Break interception.
    BreakOff,
    BreakOn,
    /// `{ONERROR branch[, msg_loc]}` — install an error trap. Stub:
    /// recognized so a macro source using it doesn't fail to lex;
    /// proper trap behavior is a follow-up once `set_error` learns
    /// to consult the macro state.
    OnError {
        branch: String,
        msg_loc: String,
    },
}

/// Errors the lexer can raise. Macro execution converts these into a
/// user-visible "Macro Error" message.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum MacroError {
    /// `{` opened but no matching `}` before end-of-input.
    UnclosedBrace,
    /// `{TOKEN}` whose name is not in the recognized set. Carries
    /// the offending body (without braces) for diagnostics.
    UnknownToken(String),
    /// `{TOKEN n}` whose count failed to parse as a non-negative
    /// integer. Carries the count fragment.
    BadRepeatCount(String),
}

impl fmt::Display for MacroError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            MacroError::UnclosedBrace => write!(f, "macro: unclosed `{{`"),
            MacroError::UnknownToken(b) => write!(f, "macro: unknown token `{{{b}}}`"),
            MacroError::BadRepeatCount(c) => write!(f, "macro: bad repeat count `{c}`"),
        }
    }
}

impl std::error::Error for MacroError {}

/// Lex a flat-keystroke macro source string. Used by the
/// `MACRO` test directive and any other caller that wants only
/// keystrokes (not control-flow directives).
///
/// Grammar:
/// - `~` outside braces → [`MacroKey::Enter`]
/// - `{TOKEN}` or `{TOKEN n}` → the corresponding key, repeated
/// - `{TILDE}` / `{LBRACE}` / `{RBRACE}` → literal `~` / `{` / `}`
/// - any other char → [`MacroKey::Char`]
///
/// Directives like `{BRANCH ...}` are rejected as unknown tokens
/// here — callers that need them go through [`lex_actions`] instead.
pub fn lex(text: &str) -> Result<Vec<MacroKey>, MacroError> {
    let mut out = Vec::new();
    for action in lex_actions(text)? {
        match action {
            MacroAction::Key(k) => out.push(k),
            other => {
                return Err(MacroError::UnknownToken(format!("{other:?}")));
            }
        }
    }
    Ok(out)
}

/// Lex a macro source string into a stream of [`MacroAction`]s,
/// recognizing both keystroke tokens and control-flow directives.
pub fn lex_actions(text: &str) -> Result<Vec<MacroAction>, MacroError> {
    let mut out = Vec::new();
    let chars: Vec<char> = text.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        match c {
            '~' => {
                i += 1;
                out.push(MacroAction::Key(MacroKey::Enter));
            }
            '{' => {
                i += 1;
                let mut body = String::new();
                let mut closed = false;
                while i < chars.len() {
                    let ch = chars[i];
                    i += 1;
                    if ch == '}' {
                        closed = true;
                        break;
                    }
                    body.push(ch);
                }
                if !closed {
                    return Err(MacroError::UnclosedBrace);
                }
                emit_brace_body(body.trim(), &mut out)?;
            }
            // `/X<verb>` — Lotus's pre-`{}` legacy macro commands.
            // Desugar inline to the equivalent brace directive so
            // the interpreter doesn't have to know about both forms.
            '/' if i + 1 < chars.len()
                && (chars[i + 1] == 'X' || chars[i + 1] == 'x')
                && i + 2 < chars.len() =>
            {
                let verb = chars[i + 2].to_ascii_uppercase();
                i += 3;
                let mut consumed = false;
                match verb {
                    'Q' => {
                        out.push(MacroAction::Quit);
                        consumed = true;
                    }
                    'R' => {
                        out.push(MacroAction::Return);
                        consumed = true;
                    }
                    'G' => {
                        let arg = read_until_tilde(&chars, &mut i);
                        out.push(MacroAction::Branch(arg));
                        consumed = true;
                    }
                    'I' => {
                        let arg = read_until_tilde(&chars, &mut i);
                        out.push(MacroAction::If(arg));
                        consumed = true;
                    }
                    'L' => {
                        let prompt_text = read_until_tilde(&chars, &mut i);
                        let loc = read_until_tilde(&chars, &mut i);
                        out.push(MacroAction::GetLabel { prompt_text, loc });
                        consumed = true;
                    }
                    'N' => {
                        let prompt_text = read_until_tilde(&chars, &mut i);
                        let loc = read_until_tilde(&chars, &mut i);
                        out.push(MacroAction::GetNumber { prompt_text, loc });
                        consumed = true;
                    }
                    'M' => {
                        let loc = read_until_tilde(&chars, &mut i);
                        out.push(MacroAction::MenuBranch(loc));
                        consumed = true;
                    }
                    'C' => {
                        let loc = read_until_tilde(&chars, &mut i);
                        out.push(MacroAction::Subroutine {
                            loc,
                            args: Vec::new(),
                        });
                        consumed = true;
                    }
                    _ => {}
                }
                if !consumed {
                    // Not a known /X verb — back up and let the
                    // bare-`/` fallthrough emit a literal `/` and
                    // re-consume the X.
                    i -= 3;
                    out.push(MacroAction::Key(MacroKey::Char('/')));
                    i += 1;
                }
            }
            _ => {
                i += 1;
                out.push(MacroAction::Key(MacroKey::Char(c)));
            }
        }
    }
    Ok(out)
}

/// Consume chars from `i` up to the next `~` (or end-of-input)
/// and advance `i` past the `~`. Used by the /X command lexer to
/// pull out arguments which Lotus terminates with `~`.
fn read_until_tilde(chars: &[char], i: &mut usize) -> String {
    let mut out = String::new();
    while *i < chars.len() {
        let ch = chars[*i];
        *i += 1;
        if ch == '~' {
            break;
        }
        out.push(ch);
    }
    out.trim().to_string()
}

/// Dispatch `{...}` body to either a keystroke token, a recognized
/// directive, or a subroutine call.
fn emit_brace_body(body: &str, out: &mut Vec<MacroAction>) -> Result<(), MacroError> {
    if body.is_empty() {
        return Err(MacroError::UnknownToken(String::new()));
    }
    // Split into head (first whitespace-delimited word) and tail.
    let (head, tail) = match body.split_once(char::is_whitespace) {
        Some((h, t)) => (h, t.trim()),
        None => (body, ""),
    };
    let head_upper = head.to_ascii_uppercase();

    // Directive forms first — they take a non-numeric argument so
    // they can't be confused with a `{TOKEN n}` repeat.
    match head_upper.as_str() {
        "BRANCH" => {
            out.push(MacroAction::Branch(tail.to_string()));
            return Ok(());
        }
        "QUIT" => {
            out.push(MacroAction::Quit);
            return Ok(());
        }
        "RETURN" => {
            out.push(MacroAction::Return);
            return Ok(());
        }
        "IF" => {
            out.push(MacroAction::If(tail.to_string()));
            return Ok(());
        }
        "DEFINE" => {
            let args = split_args(tail);
            out.push(MacroAction::Define(args));
            return Ok(());
        }
        "LET" => {
            // `loc, expr` — split on the first comma only. The
            // expr may itself contain commas (e.g. @SUM(B1,B2)),
            // which the naive `split_args` would mishandle.
            let (loc, expr) = match tail.split_once(',') {
                Some((l, e)) => (l.trim().to_string(), e.trim().to_string()),
                None => (tail.to_string(), String::new()),
            };
            out.push(MacroAction::Let { loc, expr });
            return Ok(());
        }
        "BLANK" => {
            out.push(MacroAction::Blank(tail.to_string()));
            return Ok(());
        }
        "RECALC" | "RECALCCOL" => {
            out.push(MacroAction::Recalc(tail.to_string()));
            return Ok(());
        }
        "?" => {
            // `{?}` — head is "?" (no whitespace, so tail is empty).
            out.push(MacroAction::QuestionPause);
            return Ok(());
        }
        "GETLABEL" => {
            let (prompt_text, loc) = split_last_comma(tail);
            out.push(MacroAction::GetLabel { prompt_text, loc });
            return Ok(());
        }
        "GETNUMBER" => {
            let (prompt_text, loc) = split_last_comma(tail);
            out.push(MacroAction::GetNumber { prompt_text, loc });
            return Ok(());
        }
        "MENUBRANCH" => {
            out.push(MacroAction::MenuBranch(tail.to_string()));
            return Ok(());
        }
        "MENUCALL" => {
            out.push(MacroAction::MenuCall(tail.to_string()));
            return Ok(());
        }
        "BEEP" => {
            out.push(MacroAction::Beep);
            return Ok(());
        }
        "WAIT" => {
            out.push(MacroAction::Wait(tail.to_string()));
            return Ok(());
        }
        "BREAKOFF" => {
            out.push(MacroAction::BreakOff);
            return Ok(());
        }
        "BREAKON" => {
            out.push(MacroAction::BreakOn);
            return Ok(());
        }
        "ONERROR" => {
            let (branch, msg_loc) = split_last_comma(tail);
            // No comma → entire tail is the branch target; msg_loc
            // is unset.
            let (branch, msg_loc) = if msg_loc.is_empty() && !branch.contains(',') {
                (branch, String::new())
            } else {
                (branch, msg_loc)
            };
            out.push(MacroAction::OnError { branch, msg_loc });
            return Ok(());
        }
        _ => {}
    }

    // Keystroke token (with optional integer repeat count).
    if let Some(key) = key_for_name(&head_upper) {
        let count = if tail.is_empty() {
            1
        } else {
            // Try to parse as an integer; if it fails, this is
            // either a subroutine call or a malformed token. We
            // lean on the heuristic: head was a known *key* name
            // (e.g. "DOWN"), so a non-integer tail is a real
            // error rather than "actually a subroutine called
            // DOWN".
            tail.parse::<u32>()
                .map_err(|_| MacroError::BadRepeatCount(tail.to_string()))?
        };
        for _ in 0..count {
            out.push(MacroAction::Key(key));
        }
        return Ok(());
    }

    // Otherwise: subroutine call. The head is the location (a
    // range name or cell address); the tail is a comma-separated
    // arg list.
    let args = split_args(tail);
    out.push(MacroAction::Subroutine {
        loc: head.to_string(),
        args,
    });
    Ok(())
}

fn split_args(tail: &str) -> Vec<String> {
    if tail.is_empty() {
        return Vec::new();
    }
    tail.split(',').map(|s| s.trim().to_string()).collect()
}

/// Split `tail` on the last comma so the prefix can contain commas
/// (handy for prompt strings) while the trailing arg is the
/// destination cell. If `tail` has no comma, the entire thing is
/// treated as the prefix and `loc` is empty.
fn split_last_comma(tail: &str) -> (String, String) {
    match tail.rfind(',') {
        Some(i) => (
            tail[..i].trim().to_string(),
            tail[i + 1..].trim().to_string(),
        ),
        None => (tail.trim().to_string(), String::new()),
    }
}

fn key_for_name(name: &str) -> Option<MacroKey> {
    Some(match name {
        // Literal-char escapes.
        "TILDE" => MacroKey::Char('~'),
        "LBRACE" | "OPENBRACE" => MacroKey::Char('{'),
        "RBRACE" | "CLOSEBRACE" => MacroKey::Char('}'),

        // Direction keys.
        "UP" | "U" => MacroKey::Up,
        "DOWN" | "D" => MacroKey::Down,
        "LEFT" | "L" => MacroKey::Left,
        "RIGHT" | "R" => MacroKey::Right,
        "HOME" => MacroKey::Home,
        "END" => MacroKey::End,
        "PGUP" | "PAGEUP" => MacroKey::PageUp,
        "PGDN" | "PAGEDOWN" => MacroKey::PageDown,
        "BIGLEFT" | "BIGSCREENLEFT" => MacroKey::BigLeft,
        "BIGRIGHT" | "BIGSCREENRIGHT" => MacroKey::BigRight,

        // Editing.
        "ESC" | "ESCAPE" => MacroKey::Escape,
        "BS" | "BACKSPACE" => MacroKey::Backspace,
        "DEL" | "DELETE" => MacroKey::Delete,
        "INS" | "INSERT" => MacroKey::Insert,
        "TAB" => MacroKey::Tab,

        // Function-key bindings (1-2-3 R3.4a names per SPEC §7).
        "EDIT" => MacroKey::Function(2),
        "NAME" => MacroKey::Function(3),
        "ABS" => MacroKey::Function(4),
        "GOTO" => MacroKey::Function(5),
        "WINDOW" => MacroKey::Function(6),
        "QUERY" => MacroKey::Function(7),
        "TABLE" => MacroKey::Function(8),
        "CALC" => MacroKey::Function(9),
        "GRAPH" => MacroKey::Function(10),
        "HELP" => MacroKey::Function(1),

        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn lex_ok(s: &str) -> Vec<MacroKey> {
        lex(s).expect("lex should succeed")
    }

    fn lex_actions_ok(s: &str) -> Vec<MacroAction> {
        lex_actions(s).expect("lex_actions should succeed")
    }

    #[test]
    fn empty_source_lexes_to_empty_stream() {
        assert!(lex_ok("").is_empty());
    }

    #[test]
    fn literal_chars_passthrough() {
        assert_eq!(lex_ok("ab"), vec![MacroKey::Char('a'), MacroKey::Char('b')]);
    }

    #[test]
    fn tilde_is_enter() {
        assert_eq!(lex_ok("~"), vec![MacroKey::Enter]);
    }

    #[test]
    fn unicode_chars_pass_through() {
        let want = vec![
            MacroKey::Char('c'),
            MacroKey::Char('a'),
            MacroKey::Char('f'),
            MacroKey::Char('é'),
            MacroKey::Enter,
        ];
        assert_eq!(lex_ok("café~"), want);
    }

    #[test]
    fn direction_tokens() {
        assert_eq!(
            lex_ok("{UP}{DOWN}{LEFT}{RIGHT}"),
            vec![
                MacroKey::Up,
                MacroKey::Down,
                MacroKey::Left,
                MacroKey::Right,
            ]
        );
    }

    #[test]
    fn token_case_insensitive() {
        assert_eq!(lex_ok("{down}"), vec![MacroKey::Down]);
        assert_eq!(lex_ok("{DoWn}"), vec![MacroKey::Down]);
    }

    #[test]
    fn repeat_count_emits_n_copies() {
        assert_eq!(
            lex_ok("{DOWN 3}"),
            vec![MacroKey::Down, MacroKey::Down, MacroKey::Down]
        );
    }

    #[test]
    fn repeat_count_zero_emits_nothing() {
        assert!(lex_ok("{UP 0}").is_empty());
    }

    #[test]
    fn metachar_escapes_round_trip() {
        assert_eq!(
            lex_ok("{TILDE}{LBRACE}{RBRACE}"),
            vec![
                MacroKey::Char('~'),
                MacroKey::Char('{'),
                MacroKey::Char('}')
            ]
        );
    }

    #[test]
    fn function_key_aliases() {
        assert_eq!(lex_ok("{CALC}"), vec![MacroKey::Function(9)]);
        assert_eq!(lex_ok("{GOTO}"), vec![MacroKey::Function(5)]);
        assert_eq!(lex_ok("{EDIT}"), vec![MacroKey::Function(2)]);
    }

    #[test]
    fn unclosed_brace_is_error() {
        assert_eq!(lex("{DOWN"), Err(MacroError::UnclosedBrace));
    }

    #[test]
    fn unknown_token_is_error() {
        // `{NOPE}` isn't a key, isn't a directive, has no args → it
        // falls through to "subroutine call". `lex` (flat-keystroke)
        // sees a non-Key action and reports it as unknown. That's
        // the right shape for the existing API.
        match lex("{NOPE}") {
            Err(MacroError::UnknownToken(_)) => {}
            other => panic!("expected UnknownToken, got {other:?}"),
        }
    }

    #[test]
    fn bad_repeat_count_is_error() {
        match lex("{DOWN xyz}") {
            Err(MacroError::BadRepeatCount(c)) => assert_eq!(c, "xyz"),
            other => panic!("expected BadRepeatCount, got {other:?}"),
        }
    }

    // ---- lex_actions: directives ----

    #[test]
    fn directive_branch() {
        assert_eq!(
            lex_actions_ok("{BRANCH \\J}"),
            vec![MacroAction::Branch("\\J".to_string())]
        );
    }

    #[test]
    fn directive_quit() {
        assert_eq!(lex_actions_ok("{QUIT}"), vec![MacroAction::Quit]);
    }

    #[test]
    fn directive_return() {
        assert_eq!(lex_actions_ok("{RETURN}"), vec![MacroAction::Return]);
    }

    #[test]
    fn directive_if_keeps_expression() {
        assert_eq!(
            lex_actions_ok("{IF a1>0}"),
            vec![MacroAction::If("a1>0".to_string())]
        );
    }

    #[test]
    fn subroutine_call_no_args() {
        assert_eq!(
            lex_actions_ok("{stamp}"),
            vec![MacroAction::Subroutine {
                loc: "stamp".to_string(),
                args: vec![]
            }]
        );
    }

    #[test]
    fn subroutine_call_with_args() {
        assert_eq!(
            lex_actions_ok("{stamp 1, 2}"),
            vec![MacroAction::Subroutine {
                loc: "stamp".to_string(),
                args: vec!["1".to_string(), "2".to_string()],
            }]
        );
    }

    #[test]
    fn directives_mix_with_keys() {
        assert_eq!(
            lex_actions_ok("a{BRANCH \\X}b"),
            vec![
                MacroAction::Key(MacroKey::Char('a')),
                MacroAction::Branch("\\X".to_string()),
                MacroAction::Key(MacroKey::Char('b')),
            ]
        );
    }

    // ---- /X commands desugar into directives ----

    #[test]
    fn xq_desugars_to_quit() {
        assert_eq!(lex_actions_ok("/XQ"), vec![MacroAction::Quit]);
    }

    #[test]
    fn xr_desugars_to_return() {
        assert_eq!(lex_actions_ok("/XR"), vec![MacroAction::Return]);
    }

    #[test]
    fn xg_desugars_to_branch() {
        assert_eq!(
            lex_actions_ok("/XG\\J~"),
            vec![MacroAction::Branch("\\J".to_string())]
        );
    }

    #[test]
    fn xi_desugars_to_if() {
        assert_eq!(
            lex_actions_ok("/XIa1>0~"),
            vec![MacroAction::If("a1>0".to_string())]
        );
    }

    #[test]
    fn xl_desugars_to_getlabel() {
        assert_eq!(
            lex_actions_ok("/XLName?~B5~"),
            vec![MacroAction::GetLabel {
                prompt_text: "Name?".to_string(),
                loc: "B5".to_string(),
            }]
        );
    }

    #[test]
    fn xn_desugars_to_getnumber() {
        assert_eq!(
            lex_actions_ok("/XNAge?~C5~"),
            vec![MacroAction::GetNumber {
                prompt_text: "Age?".to_string(),
                loc: "C5".to_string(),
            }]
        );
    }

    #[test]
    fn xm_desugars_to_menubranch() {
        assert_eq!(
            lex_actions_ok("/XMmenu~"),
            vec![MacroAction::MenuBranch("menu".to_string())]
        );
    }

    #[test]
    fn xc_desugars_to_subroutine_call() {
        assert_eq!(
            lex_actions_ok("/XCsub~"),
            vec![MacroAction::Subroutine {
                loc: "sub".to_string(),
                args: Vec::new(),
            }]
        );
    }

    #[test]
    fn slash_without_x_passes_through_as_char() {
        // `/F` is a slash-menu accelerator (not a /X command). The
        // lexer should emit the chars as-is so the dispatcher opens
        // the File submenu.
        assert_eq!(
            lex_actions_ok("/F"),
            vec![
                MacroAction::Key(MacroKey::Char('/')),
                MacroAction::Key(MacroKey::Char('F')),
            ]
        );
    }
}
