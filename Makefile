# L123 Makefile — convenience targets for the WK3 (Lotus 1-2-3 R3 import)
# build mode toggle. The default committed state is "WK3 off / public clean":
# pure crates.io IronCalc, no path dep on a local IronCalc clone. Toggling
# WK3 on edits two `Cargo.toml` files in place — these edits are NOT meant
# to be committed.
#
# Usage:
#   make wk3-on        Uncomment the WK3 path deps + [patch.crates-io] block
#   make wk3-off       Restore the public-clean state
#   make wk3-status    Report whether WK3 deps are currently uncommented
#   make build-wk3     wk3-on, then `cargo build -p l123 --features wk3`
#   make build         wk3-off, then `cargo build --workspace`
#   make test-wk3      wk3-on, then `cargo test --workspace --features l123-ui/wk3`
#   make test          wk3-off, then `cargo test --workspace`
#
# Targets are idempotent — running wk3-on twice is a no-op the second time.

ENGINE_TOML := crates/l123-engine/Cargo.toml
WS_TOML     := Cargo.toml

# Marker line we use to detect "WK3 currently on": the uncommented
# `ironcalc_lotus` path dep in the engine manifest.
WK3_MARKER  := ^ironcalc_lotus = { path = "../../../IronCalc/lotus" }$$

.PHONY: wk3-on wk3-off wk3-status build build-wk3 test test-wk3

wk3-on:
	@sed -i.bak 's|^# ironcalc_lotus = { path = "../../../IronCalc/lotus" }$$|ironcalc_lotus = { path = "../../../IronCalc/lotus" }|' $(ENGINE_TOML)
	@sed -i.bak 's|^# \[patch.crates-io\]$$|[patch.crates-io]|'                                                                  $(WS_TOML)
	@sed -i.bak 's|^# ironcalc      = { path = "../IronCalc/xlsx" }$$|ironcalc      = { path = "../IronCalc/xlsx" }|'             $(WS_TOML)
	@sed -i.bak 's|^# ironcalc_base = { path = "../IronCalc/base" }$$|ironcalc_base = { path = "../IronCalc/base" }|'             $(WS_TOML)
	@rm -f $(ENGINE_TOML).bak $(WS_TOML).bak
	@echo "WK3: enabled. Use --features wk3 (e.g. cargo build -p l123 --features wk3)."

wk3-off:
	@sed -i.bak 's|^ironcalc_lotus = { path = "../../../IronCalc/lotus" }$$|# ironcalc_lotus = { path = "../../../IronCalc/lotus" }|' $(ENGINE_TOML)
	@sed -i.bak 's|^\[patch.crates-io\]$$|# [patch.crates-io]|'                                                                      $(WS_TOML)
	@sed -i.bak 's|^ironcalc      = { path = "../IronCalc/xlsx" }$$|# ironcalc      = { path = "../IronCalc/xlsx" }|'                 $(WS_TOML)
	@sed -i.bak 's|^ironcalc_base = { path = "../IronCalc/base" }$$|# ironcalc_base = { path = "../IronCalc/base" }|'                 $(WS_TOML)
	@rm -f $(ENGINE_TOML).bak $(WS_TOML).bak
	@echo "WK3: disabled. Plain cargo build / test will use crates.io IronCalc."

wk3-status:
	@if grep -qE '$(WK3_MARKER)' $(ENGINE_TOML); then \
		echo "WK3: ENABLED  (path deps uncommented; --features wk3 required to compile in WK3 code)"; \
	else \
		echo "WK3: disabled (default public-clean state)"; \
	fi

build: wk3-off
	cargo build --workspace

build-wk3: wk3-on
	cargo build -p l123 --features wk3

test: wk3-off
	cargo test --workspace

test-wk3: wk3-on
	cargo test --workspace --features l123-ui/wk3
