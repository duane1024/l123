//! "Modern equivalents" of the 1-2-3 R3.4a `Worksheet Status` panel.
//! Where the DOS original proudly reported `80386` + `80387`, we fill
//! the same rows with the host's actual CPU architecture, a quip
//! about the coprocessor now being integrated, and live memory
//! numbers probed from the OS.

use std::process::Command;

#[derive(Debug, Clone)]
pub struct SysInfo {
    pub processor: String,
    pub coprocessor: String,
    pub memory_total: Option<u64>,
    pub memory_free: Option<u64>,
}

impl SysInfo {
    pub fn probe() -> Self {
        Self {
            processor: processor_string(),
            coprocessor: coprocessor_string().to_string(),
            memory_total: memory_total_bytes(),
            memory_free: memory_free_bytes(),
        }
    }
}

fn processor_string() -> String {
    let arch = std::env::consts::ARCH;
    let os = std::env::consts::OS;
    // aarch64 / darwin → "aarch64 (Apple Silicon, darwin)"-ish.
    // Keep it terse: "<arch> (<os>)". The status panel has room for
    // more detail but the column-align look of the original reads
    // better with short values.
    match (arch, os) {
        ("aarch64", "macos") => "aarch64 (Apple Silicon)".into(),
        ("x86_64", "macos") => "x86_64 (Intel Mac)".into(),
        (a, o) => format!("{a} ({o})"),
    }
}

fn coprocessor_string() -> &'static str {
    // Every CPU the project will realistically run on has an
    // integrated FPU and at least one SIMD unit; the standalone math
    // chip stopped shipping in 1989. Keep the field for nostalgia.
    "Integrated (FPU + SIMD)"
}

fn memory_total_bytes() -> Option<u64> {
    #[cfg(target_os = "macos")]
    {
        sysctl_u64("hw.memsize")
    }
    #[cfg(target_os = "linux")]
    {
        meminfo_kb("MemTotal:").map(|kb| kb * 1024)
    }
    #[cfg(not(any(target_os = "macos", target_os = "linux")))]
    {
        None
    }
}

fn memory_free_bytes() -> Option<u64> {
    #[cfg(target_os = "macos")]
    {
        // vm_stat prints `Pages free: 123456.` The page size is the
        // last token of the first line: "Mach Virtual Memory
        // Statistics: (page size of 16384 bytes)".
        let out = Command::new("vm_stat").output().ok()?;
        let body = String::from_utf8(out.stdout).ok()?;
        let page_size = body
            .lines()
            .next()
            .and_then(|l| l.split_whitespace().rev().nth(1))
            .and_then(|s| s.parse::<u64>().ok())
            .unwrap_or(4096);
        let free_pages = body
            .lines()
            .find_map(|l| l.strip_prefix("Pages free:"))
            .map(|rest| rest.trim().trim_end_matches('.'))
            .and_then(|s| s.parse::<u64>().ok())?;
        Some(free_pages * page_size)
    }
    #[cfg(target_os = "linux")]
    {
        meminfo_kb("MemAvailable:")
            .or_else(|| meminfo_kb("MemFree:"))
            .map(|kb| kb * 1024)
    }
    #[cfg(not(any(target_os = "macos", target_os = "linux")))]
    {
        None
    }
}

#[cfg(target_os = "macos")]
fn sysctl_u64(name: &str) -> Option<u64> {
    let out = Command::new("sysctl").args(["-n", name]).output().ok()?;
    String::from_utf8(out.stdout).ok()?.trim().parse().ok()
}

#[cfg(target_os = "linux")]
fn meminfo_kb(prefix: &str) -> Option<u64> {
    let body = std::fs::read_to_string("/proc/meminfo").ok()?;
    body.lines()
        .find_map(|line| line.strip_prefix(prefix))
        .and_then(|rest| rest.split_whitespace().next())
        .and_then(|s| s.parse().ok())
}

/// Render `bytes` with thousands separators — `14,093,582` — to
/// match the DOS screenshot.
pub fn format_bytes(bytes: u64) -> String {
    let s = bytes.to_string();
    let mut out = String::with_capacity(s.len() + s.len() / 3);
    for (i, c) in s.chars().rev().enumerate() {
        if i != 0 && i % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    out.chars().rev().collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn formats_thousands() {
        assert_eq!(format_bytes(0), "0");
        assert_eq!(format_bytes(42), "42");
        assert_eq!(format_bytes(1_000), "1,000");
        assert_eq!(format_bytes(14_093_582), "14,093,582");
        assert_eq!(format_bytes(15_013_244), "15,013,244");
    }

    #[test]
    fn probe_always_returns_something() {
        let info = SysInfo::probe();
        assert!(!info.processor.is_empty());
        assert!(!info.coprocessor.is_empty());
    }
}
