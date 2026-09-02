#!/usr/bin/env bash
set -Eeuo pipefail

ROOT="${STOCKWIDGET_ROOT:-$HOME/dev/StockWidget}"
DATA_BRANCH="${DATA_BRANCH:-codes-data}"
TOKEN_FILE="${GITHUB_TOKEN_FILE:-$HOME/.github_token}"
PYTHON_BIN="${PYTHON_BIN:-python3}"
DATA_DIR="${TMPDIR:-/tmp}/stockwidget-codes-data-$$"

FILES=(
  stock_sh.json
  stock_sz.json
  stock_bj.json
  fund_cn.json
  stock_hk.json
  stock_us.json
  cache_us_cn_aliases.json
  index_cn.json
  index_global.json
  futures_sh.json
  codes_update_status.json
)

cleanup() {
  git -C "$ROOT" worktree remove --force "$DATA_DIR" >/dev/null 2>&1 || true
}
trap cleanup EXIT

echo "===== StockWidget updater started at $(date -Is) ====="
cd "$ROOT"

echo "[1/6] Fetching main and $DATA_BRANCH..."
git fetch origin main "$DATA_BRANCH"

echo "[2/6] Updating main..."
git switch main
git pull --ff-only origin main

echo "[3/6] Preparing $DATA_BRANCH worktree..."
git worktree prune
git worktree add --detach "$DATA_DIR" "origin/$DATA_BRANCH"
mkdir -p "$DATA_DIR/resources"
for file in "${FILES[@]}"; do
  if [[ ! -f "$DATA_DIR/resources/$file" ]]; then
    cp "$ROOT/resources/$file" "$DATA_DIR/resources/$file"
  fi
done

echo "[4/6] Updating category JSON files..."
CODES_OUTPUT_DIR="$DATA_DIR/resources" "$PYTHON_BIN" -u scripts/update_codes.py

echo "[5/6] Creating data-only commit..."
git -C "$DATA_DIR" rm -f --ignore-unmatch \
  resources/stock_codes_list.json resources/futures_codes_list.json
paths=()
for file in "${FILES[@]}"; do
  paths+=("resources/$file")
done
git -C "$DATA_DIR" add -- "${paths[@]}"

if git -C "$DATA_DIR" diff --cached --quiet; then
  echo "No data changes to commit."
  exit 0
fi
git -C "$DATA_DIR" commit -m "chore: 更新市场代码列表"

echo "[6/6] Pushing $DATA_BRANCH..."
if [[ ! -r "$TOKEN_FILE" ]]; then
  echo "Token file is not readable: $TOKEN_FILE" >&2
  exit 1
fi
token="$(tr -d '\r\n' < "$TOKEN_FILE")"
auth="$(printf 'x-access-token:%s' "$token" | base64 -w0)"
git -C "$DATA_DIR" \
  -c "http.https://github.com/.extraheader=AUTHORIZATION: basic $auth" \
  push origin "HEAD:$DATA_BRANCH"
unset token auth

echo "===== StockWidget updater finished at $(date -Is) ====="
