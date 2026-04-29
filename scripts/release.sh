#!/usr/bin/env bash
# Cut a new l123 release end-to-end:
#   1. bump workspace version in Cargo.toml
#   2. refresh Cargo.lock, commit, tag, push
#   3. create a GitHub Release (so brew livecheck sees it)
#   4. update the homebrew tap formula and push
#
# Usage: scripts/release.sh <new-version>   (e.g. scripts/release.sh 1.0.1)
#
# Env overrides:
#   TAP_DIR   path to the homebrew-l123 checkout (default: ../homebrew-l123)
#   REMOTE    git remote to push to (default: origin)
#   SKIP_TESTS=1  skip `cargo test` (not recommended)

set -euo pipefail

if [[ $# -ne 1 ]]; then
  echo "usage: $0 <new-version>   (e.g. 1.0.1)" >&2
  exit 1
fi

new="$1"
tag="v${new}"
remote="${REMOTE:-origin}"

repo_root="$(cd "$(dirname "$0")/.." && pwd)"
tap_dir="${TAP_DIR:-${repo_root}/../homebrew-l123}"

cd "$repo_root"

# semver-ish sanity check
if ! [[ "$new" =~ ^[0-9]+\.[0-9]+\.[0-9]+([.-].+)?$ ]]; then
  echo "error: '$new' doesn't look like a version (expected X.Y.Z)" >&2
  exit 1
fi

# ----- preflight -----

current_branch="$(git rev-parse --abbrev-ref HEAD)"
if [[ "$current_branch" != "master" ]]; then
  echo "error: must be on master (currently on '$current_branch')" >&2
  exit 1
fi

if ! git diff-index --quiet HEAD --; then
  echo "error: working tree is dirty; commit or stash first" >&2
  git status --short >&2
  exit 1
fi

if git rev-parse "$tag" >/dev/null 2>&1; then
  echo "error: tag $tag already exists" >&2
  exit 1
fi

current="$(awk -F'"' '/^version = /{print $2; exit}' Cargo.toml)"
if [[ -z "$current" ]]; then
  echo "error: couldn't read current version from Cargo.toml" >&2
  exit 1
fi
echo "bumping ${current} -> ${new}" >&2

if [[ ! -d "$tap_dir/.git" ]]; then
  echo "error: tap dir '$tap_dir' is not a git checkout" >&2
  exit 1
fi
if [[ ! -x "$tap_dir/bump.sh" ]]; then
  echo "error: '$tap_dir/bump.sh' missing or not executable" >&2
  exit 1
fi

git fetch --quiet "$remote"
local_head="$(git rev-parse HEAD)"
remote_head="$(git rev-parse "$remote/master")"
if [[ "$local_head" != "$remote_head" ]]; then
  echo "error: local master is not in sync with $remote/master" >&2
  exit 1
fi

if ! command -v gh >/dev/null 2>&1; then
  echo "error: 'gh' CLI not found (needed to create the GitHub Release)" >&2
  exit 1
fi
gh auth status >/dev/null 2>&1 || { echo "error: gh not authenticated; run 'gh auth login'" >&2; exit 1; }

# ----- tests -----

if [[ "${SKIP_TESTS:-0}" != "1" ]]; then
  echo "running tests..." >&2
  cargo test --workspace --quiet
  cargo clippy --workspace --all-targets --quiet -- -D warnings
fi

# ----- bump -----

# Replace the workspace [workspace.package] version line. macOS sed needs `-i ''`.
sed -i '' -E "s/^version = \"${current//./\\.}\"$/version = \"${new}\"/" Cargo.toml

if ! grep -q "^version = \"${new}\"$" Cargo.toml; then
  echo "error: Cargo.toml version did not get rewritten as expected" >&2
  exit 1
fi

# Refresh lockfile so the new version flows through.
cargo build --workspace --quiet

git add Cargo.toml Cargo.lock
git commit -m "Bump to ${new}"

git tag -a "$tag" -m "l123 ${new}"
git push "$remote" master
git push "$remote" "$tag"

# ----- GitHub Release -----
echo "creating GitHub release ${tag}..." >&2
gh release create "$tag" --title "l123 ${new}" --generate-notes

# ----- tap formula -----
echo "updating homebrew tap in ${tap_dir}..." >&2
(
  cd "$tap_dir"
  git pull --ff-only
  ./bump.sh "$new"
  git push
)

echo "done. brew users can: brew update && brew upgrade l123" >&2
