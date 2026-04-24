//! Date/time formatting for the status line and status overlay.
//!
//! `libc::localtime_r` converts a UNIX timestamp into a broken-down
//! local time. That's the lightest way to get the user's local date
//! without pulling in a full-fat time crate; we only need to render
//! `DD-MMM-YYYY HH:MM`, not to do date arithmetic.

use std::mem::MaybeUninit;
use std::time::{SystemTime, UNIX_EPOCH};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DateTime {
    pub year: i32,
    pub month: u32,
    pub day: u32,
    pub hour: u32,
    pub minute: u32,
}

pub fn local_now() -> DateTime {
    let secs = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0);
    local_from_unix(secs)
}

pub fn local_from_unix(secs: i64) -> DateTime {
    let t: libc::time_t = secs as libc::time_t;
    unsafe {
        let mut tm: MaybeUninit<libc::tm> = MaybeUninit::zeroed();
        if libc::localtime_r(&t, tm.as_mut_ptr()).is_null() {
            return DateTime {
                year: 1970,
                month: 1,
                day: 1,
                hour: 0,
                minute: 0,
            };
        }
        let tm = tm.assume_init();
        DateTime {
            year: tm.tm_year + 1900,
            month: (tm.tm_mon + 1) as u32,
            day: tm.tm_mday as u32,
            hour: tm.tm_hour as u32,
            minute: tm.tm_min as u32,
        }
    }
}

/// 1-2-3 R3.4a status-line format: `DD-MMM-YYYY HH:MM` with a
/// three-letter English month abbreviation.
pub fn format_ddmmmyyyy_hhmm(dt: DateTime) -> String {
    format!(
        "{:02}-{}-{:04} {:02}:{:02}",
        dt.day,
        month_abbr(dt.month),
        dt.year,
        dt.hour,
        dt.minute,
    )
}

fn month_abbr(m: u32) -> &'static str {
    match m {
        1 => "Jan",
        2 => "Feb",
        3 => "Mar",
        4 => "Apr",
        5 => "May",
        6 => "Jun",
        7 => "Jul",
        8 => "Aug",
        9 => "Sep",
        10 => "Oct",
        11 => "Nov",
        12 => "Dec",
        _ => "???",
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn formats_reference_date() {
        let dt = DateTime {
            year: 2026,
            month: 4,
            day: 23,
            hour: 7,
            minute: 2,
        };
        assert_eq!(format_ddmmmyyyy_hhmm(dt), "23-Apr-2026 07:02");
    }

    #[test]
    fn zero_pads_single_digit_fields() {
        let dt = DateTime {
            year: 2026,
            month: 1,
            day: 5,
            hour: 9,
            minute: 0,
        };
        assert_eq!(format_ddmmmyyyy_hhmm(dt), "05-Jan-2026 09:00");
    }

    #[test]
    fn all_month_names_look_right() {
        for (m, want) in [
            (1, "Jan"),
            (2, "Feb"),
            (3, "Mar"),
            (4, "Apr"),
            (5, "May"),
            (6, "Jun"),
            (7, "Jul"),
            (8, "Aug"),
            (9, "Sep"),
            (10, "Oct"),
            (11, "Nov"),
            (12, "Dec"),
        ] {
            assert_eq!(month_abbr(m), want);
        }
    }

    // Sanity check that libc::localtime_r gives us something
    // plausible. We can't assert an exact result — the test machine's
    // timezone is whatever it is — but the year should be >= 1970.
    #[test]
    fn local_now_is_plausible() {
        let dt = local_now();
        assert!(dt.year >= 1970);
        assert!((1..=12).contains(&dt.month));
        assert!((1..=31).contains(&dt.day));
        assert!(dt.hour < 24);
        assert!(dt.minute < 60);
    }
}
