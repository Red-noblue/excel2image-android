#!/usr/bin/env bash
set -euo pipefail

# Local repro: export PDF on an emulator/device using a host xlsx file.
#
# Usage:
#   scripts/local_repro_export_pdf.sh "/path/to/file.xlsx"
#   scripts/local_repro_export_pdf.sh "/path/to/file.xlsx" 2   # sheetIndex=2
#
# Output:
#   .ys_files/temp/repro-out.pdf

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

XLSX_PATH="${1:-.ys_files/outputs/0221-初步/excel文件/未结算送审1006个单项明细.xlsx}"
SHEET_INDEX="${2:-${SHEET_INDEX:-0}}"

if [ ! -f "$XLSX_PATH" ]; then
  echo "ERROR: xlsx not found: $XLSX_PATH" >&2
  exit 1
fi

SDK_DIR="$(python3 - <<'PY'
import pathlib
p = pathlib.Path("local.properties")
if not p.exists():
  raise SystemExit(2)
for line in p.read_text(encoding="utf-8", errors="ignore").splitlines():
  line=line.strip()
  if line.startswith("sdk.dir="):
    print(line.split("=",1)[1].strip())
    raise SystemExit(0)
raise SystemExit(3)
PY
)"

ADB="${SDK_DIR}/platform-tools/adb"
EMU="${SDK_DIR}/emulator/emulator"

if [ ! -x "$ADB" ]; then
  echo "ERROR: adb not found: $ADB" >&2
  exit 1
fi
if [ ! -x "$EMU" ]; then
  echo "ERROR: emulator not found: $EMU" >&2
  exit 1
fi

AVD_NAME="${AVD_NAME:-$("$EMU" -list-avds | head -n 1)}"
if [ -z "${AVD_NAME:-}" ]; then
  echo "ERROR: No AVD found. Create one in Android Studio first." >&2
  exit 1
fi

PACKAGE_NAME="${PACKAGE_NAME:-com.zys.excel2image.debug}"
DEVICE_XLSX="/sdcard/Android/data/${PACKAGE_NAME}/files/repro.xlsx"
DEVICE_PDF="/sdcard/Android/data/${PACKAGE_NAME}/files/repro-out.pdf"

OUT_DIR=".ys_files/temp"
mkdir -p "$OUT_DIR"
OUT_PDF="${OUT_DIR}/repro-out.pdf"
EMULATOR_LOG="${OUT_DIR}/emulator.log"

has_device() {
  "$ADB" devices | awk 'NR>1 && $2=="device" {found=1} END{exit found?0:1}'
}

pick_device_serial() {
  # Prefer an emulator (stable for CI-like repro), otherwise pick the first authorized device.
  local serial
  serial="$("$ADB" devices | awk 'NR>1 && $2=="device" && $1 ~ /^emulator-/ {print $1; exit}')"
  if [ -n "${serial:-}" ]; then
    echo "$serial"
    return 0
  fi
  serial="$("$ADB" devices | awk 'NR>1 && $2=="device" {print $1; exit}')"
  echo "$serial"
}

if ! has_device; then
  echo "No device found. Starting emulator: ${AVD_NAME}"
  # Start emulator in background; if you prefer a GUI window, remove -no-window.
  nohup "$EMU" -avd "$AVD_NAME" -no-snapshot-save -no-boot-anim -gpu swiftshader_indirect -no-window >"$EMULATOR_LOG" 2>&1 &
  STARTED_EMULATOR=1

  echo "Waiting for device..."
  "$ADB" wait-for-device

  echo "Waiting for boot complete..."
  for _ in $(seq 1 180); do
    SERIAL="$(pick_device_serial)"
    if [ -z "${SERIAL:-}" ]; then
      sleep 2
      continue
    fi
    BOOT="$("$ADB" -s "$SERIAL" shell getprop sys.boot_completed 2>/dev/null | tr -d '\r')"
    if [ "$BOOT" = "1" ]; then
      break
    fi
    sleep 2
  done

  # First boot often keeps the system busy (dexopt, Play services), which can cause
  # "failed to complete startup" ANRs when starting instrumentation too early.
  POST_BOOT_WAIT_SECONDS="${POST_BOOT_WAIT_SECONDS:-60}"
  echo "Post-boot wait: ${POST_BOOT_WAIT_SECONDS}s (to avoid startup ANR on fresh boot)"
  sleep "${POST_BOOT_WAIT_SECONDS}"
fi

SERIAL="${ADB_SERIAL:-$(pick_device_serial)}"
if [ -z "${SERIAL:-}" ]; then
  echo "ERROR: No authorized adb device found. (If a phone is 'unauthorized', accept the RSA prompt.)" >&2
  exit 1
fi

ADB_S=( "$ADB" -s "$SERIAL" )

echo "Installing debug APKs..."
./gradlew :app:installDebug :app:installDebugAndroidTest

echo "Preparing device files dir..."
"${ADB_S[@]}" shell mkdir -p "$(dirname "$DEVICE_XLSX")"
"${ADB_S[@]}" shell rm -f "$DEVICE_XLSX" "$DEVICE_PDF" || true

echo "Pushing xlsx to device..."
"${ADB_S[@]}" push "$XLSX_PATH" "$DEVICE_XLSX" >/dev/null

echo "Running instrumentation test (export PDF)..."
TEST_PKG="${PACKAGE_NAME}.test"
"${ADB_S[@]}" shell am instrument -w -r \
  -e sheetIndex "${SHEET_INDEX}" \
  -e class com.zys.excel2image.LocalReproExportPdfTest \
  "${TEST_PKG}/androidx.test.runner.AndroidJUnitRunner"

echo "Pulling exported PDF..."
rm -f "$OUT_PDF"
"${ADB_S[@]}" pull "$DEVICE_PDF" "$OUT_PDF" >/dev/null

echo "OK: $OUT_PDF"
