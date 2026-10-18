#!/usr/bin/env bash
# ==============================================================================
# RHEL 9.8 JBoss EAP TeraTermログ (server.log / access_log.log / gc.log 検索結果)
# Excelレポート変換スクリプト (Meiryo UI / モノトーンベース)
#
# 概要:
#   jboss_log_search.sh / jboss_access_log_search.sh / jboss_gc_log_search.sh の
#   実行結果をそのまま記録したTeraTermログファイルを読み込み、各出力項目をヘッダに
#   設定して1行ずつ確認しやすいExcel (.xlsx) レポート形式で出力します。
#   RHEL 9.8 標準の Python 3 標準ライブラリのみで動作し、外部パッケージ(openpyxl等)は不要です。
#
# 使い方:
#   ./jboss_log_to_excel.sh <TeraTermログファイル...> [-o <出力先.xlsx>]
#   ./jboss_log_to_excel.sh /path/to/log_dir -o jboss_log_report.xlsx
#   INPUT_LOGS="teraterm.log" OUTPUT_EXCEL="report.xlsx" ./jboss_log_to_excel.sh
# ==============================================================================

set -euo pipefail

export TZ=Asia/Tokyo

show_usage() {
  cat <<'EOF'
Usage:
  ./jboss_log_to_excel.sh [OPTIONS] <logfile_or_dir...>

Options:
  -o, --output <path.xlsx>   出力するExcelファイルパス (デフォルト: ./jboss_log_report_<YYYYMMDD_HHMMSS>.xlsx)
  -h, --help                 このヘルプを表示

Environment Variables:
  INPUT_LOGS                 入力ログファイルまたはディレクトリ (空白またはコロン区切り)
  OUTPUT_EXCEL               出力Excelファイルパス
EOF
}

OUTPUT_PATH="${OUTPUT_EXCEL:-}"
 declare -a INPUT_ARGS=()

while [[ $# -gt 0 ]]; do
  case "$1" in
    -o|--output)
      if [[ $# -lt 2 ]]; then
        echo "[ERROR] -o / --output にはファイルパスを指定してください。" >&2
        exit 1
      fi
      OUTPUT_PATH="$2"
      shift 2
      ;;
    -h|--help)
      show_usage
      exit 0
      ;;
    --)
      shift
      while [[ $# -gt 0 ]]; do
        INPUT_ARGS+=("$1")
        shift
      done
      ;;
    -*)
      echo "[ERROR] 不明なオプションです: $1" >&2
      show_usage >&2
      exit 1
      ;;
    *)
      INPUT_ARGS+=("$1")
      shift
      ;;
  esac
done

if [[ ${#INPUT_ARGS[@]} -eq 0 && -n "${INPUT_LOGS:-}" ]]; then
  IFS=' :' read -r -a INPUT_ARGS <<< "${INPUT_LOGS}"
fi

TMP_STDIN_FILE=""
cleanup_tmp() {
  if [[ -n "$TMP_STDIN_FILE" && -f "$TMP_STDIN_FILE" ]]; then
    rm -f "$TMP_STDIN_FILE"
  fi
}
trap cleanup_tmp EXIT

if [[ ${#INPUT_ARGS[@]} -eq 0 && ! -t 0 ]]; then
  TMP_STDIN_FILE="$(mktemp "${TMPDIR:-/tmp}/jboss_teraterm_stdin.XXXXXX")"
  cat > "$TMP_STDIN_FILE"
  if [[ -s "$TMP_STDIN_FILE" ]]; then
    INPUT_ARGS+=("$TMP_STDIN_FILE")
  fi
fi

if [[ ${#INPUT_ARGS[@]} -eq 0 ]]; then
  # 引数未指定時はカレントディレクトリの .log / .txt ファイルを自動検出
  while IFS= read -r -d '' f; do
    INPUT_ARGS+=("$f")
  done < <(find . -maxdepth 1 -type f \( -name "*.log" -o -name "*.txt" \) -print0 2>/dev/null | sort -z)
fi

if [[ ${#INPUT_ARGS[@]} -eq 0 ]]; then
  echo "[ERROR] 入力となるTeraTermログファイルが見つかりません。引数にログファイルを指定してください。" >&2
  show_usage >&2
  exit 1
fi

if [[ -z "$OUTPUT_PATH" ]]; then
  OUTPUT_PATH="./jboss_log_report_$(date +%Y%m%d_%H%M%S).xlsx"
fi

PYTHON_BIN=""
for candidate in python3 /usr/bin/python3 /usr/libexec/platform-python python; do
  if command -v "$candidate" >/dev/null 2>&1; then
    if "$candidate" -c "import sys; sys.exit(0 if sys.version_info[0] >= 3 else 1)" 2>/dev/null; then
      PYTHON_BIN="$candidate"
      break
    fi
  fi
done

if [[ -z "$PYTHON_BIN" ]]; then
  echo "[ERROR] Python 3 が見つかりません。RHEL 9.8 標準の python3 パッケージを確認してください。" >&2
  exit 1
fi

"$PYTHON_BIN" - "$OUTPUT_PATH" "${INPUT_ARGS[@]}" <<'PYEOF'
import sys
import os
import re
import zipfile
import datetime
import unicodedata
from xml.sax.saxutils import escape as xml_escape

# ==============================================================================
# 1. 定数・標準ヘッダ定義
# ==============================================================================
LOG_TYPES_ORDER = ["server.log", "access_log.log", "gc.log"]

DEFAULT_HEADERS = {
    "server.log": [
        "JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Log Message"
    ],
    "access_log.log": [
        "JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Sts", "Client IP", "Request / Access Log Details"
    ],
    "gc.log": [
        "JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Pause", "GC Log Message"
    ],
}

REPORT_TITLES = {
    "server.log": "JBoss EAP Server Log Report (server.log)",
    "access_log.log": "JBoss EAP Access Log Report (access_log.log)",
    "gc.log": "JBoss EAP GC Log Report (gc.log)",
}

MIN_COL_WIDTHS = {
    "JST Timestamp": 25,
    "Cntnr": 10,
    "ECS Svc": 13,
    "Task/Dir": 12,
    "Level": 10,
    "Sts": 8,
    "Client IP": 17,
    "Pause": 13,
    "Log Message": 85,
    "Request / Access Log Details": 85,
    "GC Log Message": 85,
}

MAX_COL_WIDTHS = {
    "JST Timestamp": 30,
    "Cntnr": 18,
    "ECS Svc": 22,
    "Task/Dir": 22,
    "Level": 14,
    "Sts": 12,
    "Client IP": 28,
    "Pause": 18,
    "Log Message": 220,
    "Request / Access Log Details": 220,
    "GC Log Message": 220,
}

# ==============================================================================
# 2. 正規表現パターン (ANSI除去・TeraTermタイムスタンプ除去・各行解析)
# ==============================================================================
RE_ANSI_OSC = re.compile(r"\x1b\][^\x07\x1b]*(?:\x07|\x1b\\)")
RE_ANSI_CSI = re.compile(r"\x1b\[[0-9;?]*[ -/]*[@-~]")
RE_ANSI_ESC = re.compile(r"\x1b[@-Z\\-_]|\x1b")
RE_CTRL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

# TeraTermのログタイムスタンプ付与機能 ([YYYY-MM-DD HH:MM:SS.mmm] 等)
# ※JVM Unified GCログの [2026-08-20T00:00:00.000+0900][info] は空白区切りではなく直後に [ が続くため区別可能
RE_TERATERM_TS = re.compile(
    r"^\[(?:\d{4}[-/]\d{2}[-/]\d{2}|[A-Z][a-z]{2}\s+[A-Z][a-z]{2}\s+\d{1,2})\s+"
    r"\d{2}:\d{2}:\d{2}(?:[.,]\d{1,6})?(?:\s+\d{4})?\]\s+(?!\[(?:info|warn|warning|error|debug|trace|gc)\b)"
)

RE_BANNER_TITLE = re.compile(
    r"^\s*JBoss EAP (Log|Access Log|GC Log) Viewer \(JST Sorted\)\s*\|\s*Range:\s*"
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\d{3})\s+to\s+"
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\d{3})\s*\(JST\)\s*$"
)
RE_BANNER_TARGET = re.compile(
    r"^\s*Target:\s*Container=\[(.*?)\]\s+ECS-Service=\[(.*?)\]\s+ScannedFiles=(\d+)\s*$"
)
RE_BANNER_THRESHOLDS = re.compile(
    r"^\s*Thresholds:\s*(WARN\s*>=\s*\d+ms\s*\|\s*ERROR\s*>=\s*\d+ms.*)$"
)
RE_TABLE_HEADER = re.compile(
    r"^\s*(JST Timestamp\s*\|\s*Cntnr\s*\|\s*ECS Svc\s*\|\s*Task/Dir\s*\|\s*Level\s*\|.+)$"
)
RE_TABLE_DIVIDER = re.compile(r"^\s*-{15,}\+-{5,}\+-{5,}\+")
RE_SEP_LINE = re.compile(r"^\s*={40,}\s*$")
RE_NO_MATCH = re.compile(r"^\s*\[!\]\s*No matching .*entries found in the specified JST time range\.")

RE_SUMMARY_LINE = re.compile(
    r"^\s*Summary:\s*Total=(\d+)\s*\[\s*ERROR(?:\(5xx\))?=(\d+)\s+"
    r"WARN(?:\(4xx\))?=(\d+)\s+INFO(?:/OTHER|\(2xx/3xx\))?=(\d+)\s*\]"
    r"(?:\s+MaxPause=([0-9.]+ms))?\s*$"
)
RE_BREAKDOWN_ITEM = re.compile(
    r"^\s*-\s*([A-Za-z0-9_.-]+)\s*/\s*([A-Za-z0-9_.-]+)\s*:\s*Total=(\d+)\s*"
    r"\(\s*ERROR=(\d+),\s*WARN=(\d+)(?:,\s*INFO=(\d+))?(?:,\s*MaxPause=([0-9.]+ms))?\s*\)\s*$"
)

# データ行 (メイン行 & 継続行 "->")
RE_MAIN_ROW = re.compile(
    r"^\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\d{3})\s*\|\s*"
    r"([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([A-Z0-9_-]+)\s*\|(.*)$"
)
RE_CONT_ROW = re.compile(
    r"^(\s*)\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]*?)\s*\|\s*(->)\s*\|(.*)$"
)

# access_log / gc_log の後半カラム分解用
RE_ACCESS_REST = re.compile(r"^\s*([0-9]{3}|---)\s*\|\s*([^|]*?)\s*\|(.*)$")
RE_ACCESS_CONT_REST = re.compile(r"^\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|(.*)$")
RE_GC_REST = re.compile(r"^\s*([+-]?[0-9]+(?:\.[0-9]+)?ms|-)\s*\|(.*)$")
RE_GC_CONT_REST = re.compile(r"^\s*([^|]*?)\s*\|(.*)$")


def clean_line(raw_line: str) -> str:
    s = raw_line.rstrip("\r\n")
    # バックスペース制御文字が含まれる場合の簡易正規化
    while "\x08" in s:
        new_s = re.sub(r"[^\x08]\x08", "", s, count=1)
        if new_s == s:
            s = s.replace("\x08", "")
            break
        s = new_s
    s = RE_ANSI_OSC.sub("", s)
    s = RE_ANSI_CSI.sub("", s)
    s = RE_ANSI_ESC.sub("", s)
    s = RE_CTRL_CHARS.sub("", s)
    s = RE_TERATERM_TS.sub("", s)
    return s


def read_text_auto_encoding(filepath: str):
    with open(filepath, "rb") as f:
        data = f.read()
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig", errors="replace").splitlines()
    for enc in ("utf-8", "cp932", "euc_jp"):
        try:
            return data.decode(enc).splitlines()
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace").splitlines()


def strip_single_leading_space(s: str) -> str:
    return s[1:] if s.startswith(" ") else s


def infer_log_type_from_rest(rest: str, is_cont: bool) -> str:
    if not is_cont:
        if RE_ACCESS_REST.match(rest):
            return "access_log.log"
        if RE_GC_REST.match(rest):
            return "gc.log"
        return "server.log"
    else:
        # 継続行の場合の推定
        if RE_ACCESS_CONT_REST.match(rest):
            m = RE_ACCESS_CONT_REST.match(rest)
            if m.group(1) == "" and m.group(2) == "":
                return "access_log.log"
        if RE_GC_CONT_REST.match(rest):
            m = RE_GC_CONT_REST.match(rest)
            if m.group(1) == "":
                return "gc.log"
        return "server.log"


def split_row_columns(log_type: str, ts: str, cnt: str, svc: str, tsk: str, lvl: str, rest: str, is_cont: bool):
    if log_type == "access_log.log":
        m = (RE_ACCESS_CONT_REST if is_cont else RE_ACCESS_REST).match(rest)
        if not m:
            m = RE_ACCESS_CONT_REST.match(rest)
        if m:
            sts = m.group(1).strip()
            ip = m.group(2).strip()
            msg = strip_single_leading_space(m.group(3))
            return [ts, cnt, svc, tsk, lvl, sts, ip, msg]
        return [ts, cnt, svc, tsk, lvl, "", "", strip_single_leading_space(rest)]
    elif log_type == "gc.log":
        m = (RE_GC_CONT_REST if is_cont else RE_GC_REST).match(rest)
        if not m:
            m = RE_GC_CONT_REST.match(rest)
        if m:
            pause = m.group(1).strip()
            msg = strip_single_leading_space(m.group(2))
            return [ts, cnt, svc, tsk, lvl, pause, msg]
        return [ts, cnt, svc, tsk, lvl, "", strip_single_leading_space(rest)]
    else:
        msg = strip_single_leading_space(rest)
        return [ts, cnt, svc, tsk, lvl, msg]


# ==============================================================================
# 3. ログデータ構造とパーサ
# ==============================================================================
class LogBucket:
    def __init__(self, log_type: str):
        self.log_type = log_type
        self.headers = list(DEFAULT_HEADERS[log_type])
        self.ranges = []
        self.containers = []
        self.services = []
        self.scanned_files = 0
        self.thresholds = ""
        self.rows = []  # list of dict: {"cols": [...], "is_cont": bool, "severity": str, "entry_idx": int}
        self.seen_section = False
        self.entry_count = 0
        self.err_count = 0
        self.warn_count = 0
        self.info_count = 0
        self.max_pause_ms = 0.0
        self.cs_stats = {}  # (cnt, svc) -> {"total": 0, "err": 0, "warn": 0, "info": 0, "max_pause": 0.0}

    def add_metadata_range(self, start_ts: str, end_ts: str):
        r = f"{start_ts} to {end_ts}"
        if r not in self.ranges:
            self.ranges.append(r)

    def add_metadata_target(self, cnt: str, svc: str, scanned: int):
        if cnt and cnt not in self.containers:
            self.containers.append(cnt)
        if svc and svc not in self.services:
            self.services.append(svc)
        self.scanned_files += scanned

    def add_row(self, cols: list, is_cont: bool):
        self.seen_section = True
        if not is_cont:
            self.entry_count += 1
            lvl = cols[4].strip().upper()
            if lvl in ("ERROR", "FATAL", "SEVERE"):
                sev = "ERROR"
                self.err_count += 1
            elif lvl in ("WARN", "WARNING"):
                sev = "WARN"
                self.warn_count += 1
            elif lvl in ("DEBUG", "TRACE"):
                sev = "DEBUG"
                self.info_count += 1
            else:
                sev = "INFO"
                self.info_count += 1

            cnt = cols[1].strip() or "-"
            svc = cols[2].strip() or "-"
            key = (cnt, svc)
            if key not in self.cs_stats:
                self.cs_stats[key] = {"total": 0, "err": 0, "warn": 0, "info": 0, "max_pause": 0.0}
            self.cs_stats[key]["total"] += 1
            if sev == "ERROR":
                self.cs_stats[key]["err"] += 1
            elif sev == "WARN":
                self.cs_stats[key]["warn"] += 1
            else:
                self.cs_stats[key]["info"] += 1

            if self.log_type == "gc.log" and len(cols) >= 6:
                p_str = cols[5].strip()
                m_p = re.match(r"^([0-9]+(?:\.[0-9]+)?)ms$", p_str)
                if m_p:
                    p_val = float(m_p.group(1))
                    if p_val > self.max_pause_ms:
                        self.max_pause_ms = p_val
                    if p_val > self.cs_stats[key]["max_pause"]:
                        self.cs_stats[key]["max_pause"] = p_val

            self.rows.append({
                "cols": cols,
                "is_cont": False,
                "severity": sev,
                "entry_idx": self.entry_count,
            })
        else:
            parent_sev = self.rows[-1]["severity"] if self.rows else "INFO"
            parent_idx = self.rows[-1]["entry_idx"] if self.rows else 1
            self.rows.append({
                "cols": cols,
                "is_cont": True,
                "severity": parent_sev,
                "entry_idx": parent_idx,
            })


# 生ログフォールバック用正規表現 (TeraTerm上で直接 cat / tail したログにも対応)
RE_RAW_SERVER = re.compile(
    r"^(\d{4}[-/]\d{2}[-/]\d{2})[T ](\d{2}:\d{2}:\d{2})(?:[,.](\d{1,6}))?(?:Z|[+-]\d{2}:?\d{2})?\s+\[?([A-Z]+)\]?\s+(.*)$"
)
RE_RAW_ACCESS = re.compile(
    r'^(\S+)\s+\S+\s+(\S+)\s+\[(\d{1,2})/([A-Za-z]{3})/(\d{4}):(\d{2}:\d{2}:\d{2})(?:[.,](\d{1,6}))?\s*([+-]\d{4}|Z)?\]\s+("[^"]*")\s+([1-5]\d{2})(?:\s+(.*))?$'
)
RE_RAW_GC_UNIFIED = re.compile(
    r"^\[(\d{4}[-/]\d{2}[-/]\d{2})[T ](\d{2}:\d{2}:\d{2})(?:[.,](\d{1,6}))?(?:Z|[+-]\d{2}:?\d{2})?\](?:\[[^\]]*\])*\[?(info|warning|warn|error|debug|trace)\]?(?:\[([^\]]*)\])?\s*(.*)$",
    re.IGNORECASE,
)
MONTH_MAP = {
    "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06",
    "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
}


def parse_teraterm_files(filepaths: list):
    buckets = {lt: LogBucket(lt) for lt in LOG_TYPES_ORDER}

    for fp in filepaths:
        current_type = None
        pending_target = None
        pending_thresholds = None
        file_matched_table_rows = 0
        lines = read_text_auto_encoding(fp)

        for raw_line in lines:
            line = clean_line(raw_line)
            if not line.strip():
                continue

            # 1. バナータイトル行の検出
            m_title = RE_BANNER_TITLE.match(line)
            if m_title:
                viewer_kind = m_title.group(1)
                if viewer_kind == "Log":
                    current_type = "server.log"
                elif viewer_kind == "Access Log":
                    current_type = "access_log.log"
                elif viewer_kind == "GC Log":
                    current_type = "gc.log"
                buckets[current_type].seen_section = True
                buckets[current_type].add_metadata_range(m_title.group(2), m_title.group(3))
                continue

            # 2. Targetメタデータ行の検出
            m_target = RE_BANNER_TARGET.match(line)
            if m_target:
                cnt_t, svc_t, sc_cnt = m_target.group(1), m_target.group(2), int(m_target.group(3))
                if current_type:
                    buckets[current_type].add_metadata_target(cnt_t, svc_t, sc_cnt)
                else:
                    pending_target = (cnt_t, svc_t, sc_cnt)
                continue

            # 3. Thresholds行の検出 (GC Log)
            m_th = RE_BANNER_THRESHOLDS.match(line)
            if m_th:
                th_str = m_th.group(1).strip()
                if current_type == "gc.log":
                    buckets["gc.log"].thresholds = th_str
                else:
                    pending_thresholds = th_str
                continue

            # 4. テーブルヘッダ行の検出 (各出力項目をヘッダとして取得)
            m_hdr = RE_TABLE_HEADER.match(line)
            if m_hdr:
                if "printf " in line or "%s" in line or "\\n" in line:
                    continue
                hdr_parts = [p.strip() for p in m_hdr.group(1).split("|")]
                if len(hdr_parts) == 8 or "Client IP" in hdr_parts:
                    current_type = "access_log.log"
                elif len(hdr_parts) == 7 or "Pause" in hdr_parts:
                    current_type = "gc.log"
                else:
                    current_type = "server.log"

                buckets[current_type].seen_section = True
                expected_len = len(DEFAULT_HEADERS[current_type])
                if len(hdr_parts) == expected_len:
                    buckets[current_type].headers = hdr_parts
                if pending_target:
                    buckets[current_type].add_metadata_target(*pending_target)
                    pending_target = None
                if pending_thresholds and current_type == "gc.log":
                    buckets["gc.log"].thresholds = pending_thresholds
                    pending_thresholds = None
                continue

            # 5. 区切り線・該当なしメッセージ・サマリー行のスキップ/処理
            if RE_TABLE_DIVIDER.match(line) or RE_SEP_LINE.match(line) or RE_NO_MATCH.match(line):
                continue
            if RE_SUMMARY_LINE.match(line) or RE_BREAKDOWN_ITEM.match(line) or line.strip().startswith("Breakdown by Container"):
                current_type = None
                continue

            # 6. メインデータ行の検出
            m_main = RE_MAIN_ROW.match(line)
            if m_main:
                ts, cnt, svc, tsk, lvl, rest = (
                    m_main.group(1).strip(),
                    m_main.group(2).strip(),
                    m_main.group(3).strip(),
                    m_main.group(4).strip(),
                    m_main.group(5).strip(),
                    m_main.group(6),
                )
                lt = current_type if current_type else infer_log_type_from_rest(rest, False)
                current_type = lt
                cols = split_row_columns(lt, ts, cnt, svc, tsk, lvl, rest, False)
                buckets[lt].add_row(cols, False)
                file_matched_table_rows += 1
                continue

            # 7. 継続行 ("->") の検出
            m_cont = RE_CONT_ROW.match(line)
            if m_cont:
                cnt, svc, tsk, lvl, rest = (
                    m_cont.group(2).strip(),
                    m_cont.group(3).strip(),
                    m_cont.group(4).strip(),
                    m_cont.group(5).strip(),
                    m_cont.group(6),
                )
                lt = current_type if current_type else infer_log_type_from_rest(rest, True)
                current_type = lt
                cols = split_row_columns(lt, "", cnt, svc, tsk, lvl, rest, True)
                buckets[lt].add_row(cols, True)
                file_matched_table_rows += 1
                continue

        # フォールバック: パイプ整形テーブル行が1行も含まれない生ログファイルの場合
        if file_matched_table_rows == 0:
            last_raw_lt = None
            for raw_line in lines:
                line = clean_line(raw_line)
                if not line.strip():
                    continue
                m_gc = RE_RAW_GC_UNIFIED.match(line)
                if m_gc:
                    d_s = m_gc.group(1).replace("/", "-")
                    ms_s = ((m_gc.group(3) or "") + "000")[:3]
                    ts = f"{d_s} {m_gc.group(2)},{ms_s}"
                    raw_lvl = (m_gc.group(4) or "INFO").upper()
                    if raw_lvl == "WARNING":
                        raw_lvl = "WARN"
                    msg = m_gc.group(6).strip()
                    p_m = re.search(r"([0-9]+(?:\.[0-9]+)?)ms(?:\s|$)", msg)
                    pause_s = f"{float(p_m.group(1)):.1f}ms" if p_m else "-"
                    if "Pause Full" in msg or "OutOfMemory" in msg or (p_m and float(p_m.group(1)) >= 1000.0):
                        raw_lvl = "ERROR"
                    elif p_m and float(p_m.group(1)) >= 200.0:
                        raw_lvl = "WARN"
                    buckets["gc.log"].add_row([ts, "-", "-", "-", raw_lvl, pause_s, msg], False)
                    last_raw_lt = "gc.log"
                    continue
                m_acc = RE_RAW_ACCESS.match(line)
                if m_acc:
                    ip, usr, dd, mon, yyyy, t_s, ms_raw, _, req, sts, tail = m_acc.groups()
                    mm = MONTH_MAP.get(mon.upper(), "01")
                    ms_s = ((ms_raw or "") + "000")[:3]
                    ts = f"{int(yyyy):04d}-{mm}-{int(dd):02d} {t_s},{ms_s}"
                    lvl = "ERROR" if sts.startswith("5") else ("WARN" if sts.startswith("4") else "INFO")
                    u_part = f"[user:{usr}] " if usr and usr != "-" else ""
                    t_part = f" {tail}" if tail else ""
                    buckets["access_log.log"].add_row([ts, "-", "-", "-", lvl, sts, ip, f"{u_part}{req}{t_part}"], False)
                    last_raw_lt = "access_log.log"
                    continue
                m_srv = RE_RAW_SERVER.match(line)
                if m_srv:
                    d_s = m_srv.group(1).replace("/", "-")
                    ms_s = ((m_srv.group(3) or "") + "000")[:3]
                    ts = f"{d_s} {m_srv.group(2)},{ms_s}"
                    lvl = m_srv.group(4).upper()
                    msg = m_srv.group(5)
                    buckets["server.log"].add_row([ts, "-", "-", "-", lvl, msg], False)
                    last_raw_lt = "server.log"
                    continue
                if last_raw_lt == "server.log" and buckets["server.log"].rows:
                    buckets["server.log"].add_row(["", "-", "-", "", "->", line], True)

    return buckets


# ==============================================================================
# 4. OpenXML (.xlsx) ワークブック生成エンジン (Meiryo UI / モノトーンベース)
# ==============================================================================
def display_width(text: str) -> int:
    if not text:
        return 0
    w = 0
    for ch in str(text):
        eaw = unicodedata.east_asian_width(ch)
        w += 2 if eaw in ("W", "F", "A") else 1
    return w


def col_letter(col_idx_1based: int) -> str:
    res = ""
    n = col_idx_1based
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(65 + rem) + res
    return res


def sanitize_xml_text(s: str) -> str:
    if s is None:
        return ""
    cleaned = RE_CTRL_CHARS.sub("", str(s))
    if len(cleaned) > 32760:
        cleaned = cleaned[:32757] + "..."
    return xml_escape(cleaned, {'"': "&quot;", "'": "&apos;"})


class SharedStringTable:
    def __init__(self):
        self.strings = []
        self.index_map = {}
        self.total_refs = 0

    def get_index(self, text: str) -> int:
        self.total_refs += 1
        s = "" if text is None else str(text)
        if len(s) > 32760:
            s = s[:32757] + "..."
        s = RE_CTRL_CHARS.sub("", s)
        idx = self.index_map.get(s)
        if idx is None:
            idx = len(self.strings)
            self.strings.append(s)
            self.index_map[s] = idx
        return idx

    def to_xml(self) -> str:
        parts = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{self.total_refs}" uniqueCount="{len(self.strings)}">',
        ]
        for s in self.strings:
            esc = xml_escape(s, {'"': "&quot;", "'": "&apos;"})
            if s.startswith(" ") or s.endswith(" ") or "\n" in s or "\t" in s or "  " in s:
                parts.append(f'<si><t xml:space="preserve">{esc}</t></si>')
            else:
                parts.append(f"<si><t>{esc}</t></si>")
        parts.append("</sst>")
        return "".join(parts)


# スタイルインデックス定義 (すべて Meiryo UI / モノトーン配色)
STYLE_DEFAULT = 0
STYLE_TITLE_BANNER = 1       # #1F1F1F 背景 + 白太字 13pt
STYLE_SUBTITLE_BAR = 2       # #333333 背景 + 白文字 9.5pt
STYLE_META_BAR = 3           # #D9D9D9 背景 + 黒太字 9.5pt
STYLE_SUMMARY_BAR = 4        # #F2F2F2 背景 + 濃灰文字 9.5pt
STYLE_SECTION_HDR = 5        # 白背景 + 下太線 + 黒太字 11pt
STYLE_TBL_HEADER_C = 6       # #262626 背景 + 白太字 10pt (中央揃え)
STYLE_TBL_HEADER_L = 7       # #262626 背景 + 白太字 10pt (左揃え)
STYLE_TBL_HEADER_R = 8       # #262626 背景 + 白太字 10pt (右揃え)

# データ行 (INFO 偶数行: #FFFFFF / 奇数行: #F8F8F8)
STYLE_INFO_EVEN_C = 9
STYLE_INFO_EVEN_L = 10
STYLE_INFO_EVEN_R = 11
STYLE_INFO_ODD_C = 12
STYLE_INFO_ODD_L = 13
STYLE_INFO_ODD_R = 14

# データ行 (DEBUG/TRACE 偶数行 / 奇数行: 落ち着いたグレー文字 #595959)
STYLE_DBG_EVEN_C = 15
STYLE_DBG_EVEN_L = 16
STYLE_DBG_EVEN_R = 17
STYLE_DBG_ODD_C = 18
STYLE_DBG_ODD_L = 19
STYLE_DBG_ODD_R = 20

# データ行 (WARN: #ECECEC 背景 + #1A1A1A 太字 / バッジは #595959 背景 + 白太字)
STYLE_WARN_C = 21
STYLE_WARN_L = 22
STYLE_WARN_R = 23
STYLE_WARN_BADGE_C = 24
STYLE_WARN_BADGE_R = 25

# データ行 (ERROR: #D9D9D9 背景 + #000000 太字 / バッジは #1F1F1F 背景 + 白太字)
STYLE_ERR_C = 26
STYLE_ERR_L = 27
STYLE_ERR_R = 28
STYLE_ERR_BADGE_C = 29
STYLE_ERR_BADGE_R = 30

# 継続行 "->" (上部点線ボーダーで親行と視覚的に結合)
STYLE_CONT_INFO_C = 31
STYLE_CONT_INFO_L = 32
STYLE_CONT_INFO_R = 33
STYLE_CONT_WARN_C = 34
STYLE_CONT_WARN_L = 35
STYLE_CONT_WARN_R = 36
STYLE_CONT_ERR_C = 37
STYLE_CONT_ERR_L = 38
STYLE_CONT_ERR_R = 39

# サマリーシート数値セル用 (#,##0)
STYLE_NUM_EVEN = 40
STYLE_NUM_ODD = 41
STYLE_NUM_ERR = 42
STYLE_NUM_WARN = 43


def build_styles_xml() -> str:
    # すべてのフォントを Meiryo UI に統一
    fonts = [
        # 0: Default 10pt #1A1A1A
        '<font><sz val="10"/><color rgb="FF1A1A1A"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 1: Title Banner 13pt Bold #FFFFFF
        '<font><b/><sz val="13"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 2: Subtitle Bar 9.5pt #E6E6E6
        '<font><sz val="9.5"/><color rgb="FFE6E6E6"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 3: Meta Bar 9.5pt Bold #1A1A1A
        '<font><b/><sz val="9.5"/><color rgb="FF1A1A1A"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 4: Summary Bar 9.5pt #262626
        '<font><sz val="9.5"/><color rgb="FF262626"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 5: Section Header 11pt Bold #1F1F1F
        '<font><b/><sz val="11"/><color rgb="FF1F1F1F"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 6: Table Header & Inverted Badge 10pt Bold #FFFFFF
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 7: ERROR Row 10pt Bold #000000
        '<font><b/><sz val="10"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 8: WARN Row 10pt Bold #1A1A1A
        '<font><b/><sz val="10"/><color rgb="FF1A1A1A"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 9: Muted / DEBUG / Cont INFO 10pt #595959
        '<font><sz val="10"/><color rgb="FF595959"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
        # 10: Cont WARN/ERR 10pt #262626
        '<font><sz val="10"/><color rgb="FF262626"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
    ]

    # モノトーンカラースケール塗りつぶし定義 (fillId 0, 1 は Excel 予約)
    fills = [
        '<fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="gray125"/></fill>',
        # 2: Deep Off-Black #1F1F1F (Title Banner / ERROR Badge)
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F1F1F"/><bgColor indexed="64"/></patternFill></fill>',
        # 3: Dark Charcoal #333333 (Subtitle Bar)
        '<fill><patternFill patternType="solid"><fgColor rgb="FF333333"/><bgColor indexed="64"/></patternFill></fill>',
        # 4: Table Header #262626
        '<fill><patternFill patternType="solid"><fgColor rgb="FF262626"/><bgColor indexed="64"/></patternFill></fill>',
        # 5: WARN Badge #595959
        '<fill><patternFill patternType="solid"><fgColor rgb="FF595959"/><bgColor indexed="64"/></patternFill></fill>',
        # 6: ERROR Row & Meta Bar #D9D9D9
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>',
        # 7: ERROR Continuation Row #E6E6E6
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE6E6E6"/><bgColor indexed="64"/></patternFill></fill>',
        # 8: WARN Row #ECECEC
        '<fill><patternFill patternType="solid"><fgColor rgb="FFECECEC"/><bgColor indexed="64"/></patternFill></fill>',
        # 9: WARN Continuation Row #F4F4F4
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF4F4F4"/><bgColor indexed="64"/></patternFill></fill>',
        # 10: Pure White #FFFFFF (INFO Even Row)
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>',
        # 11: Ultra-Light Gray #F8F8F8 (INFO Odd Row / Summary Bar)
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF8F8F8"/><bgColor indexed="64"/></patternFill></fill>',
    ]

    borders = [
        # 0: None
        "<border><left/><right/><top/><bottom/><diagonal/></border>",
        # 1: Table Header Border
        '<border>'
        '<left style="thin"><color rgb="FF595959"/></left>'
        '<right style="thin"><color rgb="FF595959"/></right>'
        '<top style="medium"><color rgb="FF1F1F1F"/></top>'
        '<bottom style="medium"><color rgb="FF000000"/></bottom>'
        '<diagonal/></border>',
        # 2: Standard Data Row Border
        '<border>'
        '<left style="thin"><color rgb="FFD0D0D0"/></left>'
        '<right style="thin"><color rgb="FFD0D0D0"/></right>'
        '<top style="thin"><color rgb="FFBFBFBF"/></top>'
        '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom>'
        '<diagonal/></border>',
        # 3: Continuation Row ("->") Border (top dotted to group with parent entry)
        '<border>'
        '<left style="thin"><color rgb="FFD0D0D0"/></left>'
        '<right style="thin"><color rgb="FFD0D0D0"/></right>'
        '<top style="dotted"><color rgb="FFBFBFBF"/></top>'
        '<bottom style="thin"><color rgb="FFBFBFBF"/></bottom>'
        '<diagonal/></border>',
        # 4: Banner / Metadata Box Border
        '<border>'
        '<left style="thin"><color rgb="FF808080"/></left>'
        '<right style="thin"><color rgb="FF808080"/></right>'
        '<top style="thin"><color rgb="FF808080"/></top>'
        '<bottom style="thin"><color rgb="FF808080"/></bottom>'
        '<diagonal/></border>',
        # 5: Section Header Bottom Border
        '<border><left/><right/><top/><bottom style="medium"><color rgb="FF262626"/></bottom><diagonal/></border>',
    ]

    def xf(num_fmt: int, font: int, fill: int, border: int, h_align: str, wrap: int = 0, indent: int = 0) -> str:
        ind_attr = f' indent="{indent}"' if indent > 0 else ""
        return (
            f'<xf numFmtId="{num_fmt}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" '
            f'applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
            f'<alignment horizontal="{h_align}" vertical="center" wrapText="{wrap}"{ind_attr}/></xf>'
        )

    cell_xfs = [
        # 0: Default
        xf(49, 0, 0, 0, "left"),
        # 1..5: Banners & Section Header
        xf(49, 1, 2, 4, "left", 0, 1),   # 1: STYLE_TITLE_BANNER
        xf(49, 2, 3, 4, "left", 0, 1),   # 2: STYLE_SUBTITLE_BAR
        xf(49, 3, 6, 4, "left", 0, 1),   # 3: STYLE_META_BAR
        xf(49, 4, 11, 4, "left", 0, 1),  # 4: STYLE_SUMMARY_BAR
        xf(49, 5, 0, 5, "left"),         # 5: STYLE_SECTION_HDR
        # 6..8: Table Column Headers
        xf(49, 6, 4, 1, "center"),       # 6: STYLE_TBL_HEADER_C
        xf(49, 6, 4, 1, "left"),         # 7: STYLE_TBL_HEADER_L
        xf(49, 6, 4, 1, "right"),        # 8: STYLE_TBL_HEADER_R
        # 9..14: INFO Even / Odd
        xf(49, 0, 10, 2, "center"),      # 9:  STYLE_INFO_EVEN_C
        xf(49, 0, 10, 2, "left"),        # 10: STYLE_INFO_EVEN_L
        xf(49, 0, 10, 2, "right"),       # 11: STYLE_INFO_EVEN_R
        xf(49, 0, 11, 2, "center"),      # 12: STYLE_INFO_ODD_C
        xf(49, 0, 11, 2, "left"),        # 13: STYLE_INFO_ODD_L
        xf(49, 0, 11, 2, "right"),       # 14: STYLE_INFO_ODD_R
        # 15..20: DEBUG Even / Odd
        xf(49, 9, 10, 2, "center"),      # 15: STYLE_DBG_EVEN_C
        xf(49, 9, 10, 2, "left"),        # 16: STYLE_DBG_EVEN_L
        xf(49, 9, 10, 2, "right"),       # 17: STYLE_DBG_EVEN_R
        xf(49, 9, 11, 2, "center"),      # 18: STYLE_DBG_ODD_C
        xf(49, 9, 11, 2, "left"),        # 19: STYLE_DBG_ODD_L
        xf(49, 9, 11, 2, "right"),       # 20: STYLE_DBG_ODD_R
        # 21..25: WARN Row & Badge
        xf(49, 8, 8, 2, "center"),       # 21: STYLE_WARN_C
        xf(49, 8, 8, 2, "left"),         # 22: STYLE_WARN_L
        xf(49, 8, 8, 2, "right"),        # 23: STYLE_WARN_R
        xf(49, 6, 5, 2, "center"),       # 24: STYLE_WARN_BADGE_C
        xf(49, 6, 5, 2, "right"),        # 25: STYLE_WARN_BADGE_R
        # 26..30: ERROR Row & Badge
        xf(49, 7, 6, 2, "center"),       # 26: STYLE_ERR_C
        xf(49, 7, 6, 2, "left"),         # 27: STYLE_ERR_L
        xf(49, 7, 6, 2, "right"),        # 28: STYLE_ERR_R
        xf(49, 6, 2, 2, "center"),       # 29: STYLE_ERR_BADGE_C
        xf(49, 6, 2, 2, "right"),        # 30: STYLE_ERR_BADGE_R
        # 31..39: Continuation Rows ("->")
        xf(49, 9, 11, 3, "center"),      # 31: STYLE_CONT_INFO_C
        xf(49, 9, 11, 3, "left"),        # 32: STYLE_CONT_INFO_L
        xf(49, 9, 11, 3, "right"),       # 33: STYLE_CONT_INFO_R
        xf(49, 10, 9, 3, "center"),      # 34: STYLE_CONT_WARN_C
        xf(49, 10, 9, 3, "left"),        # 35: STYLE_CONT_WARN_L
        xf(49, 10, 9, 3, "right"),       # 36: STYLE_CONT_WARN_R
        xf(49, 10, 7, 3, "center"),      # 37: STYLE_CONT_ERR_C
        xf(49, 10, 7, 3, "left"),        # 38: STYLE_CONT_ERR_L
        xf(49, 10, 7, 3, "right"),       # 39: STYLE_CONT_ERR_R
        # 40..43: Numeric Summary Cells (#,##0 = numFmtId 3)
        xf(3, 0, 10, 2, "right"),        # 40: STYLE_NUM_EVEN
        xf(3, 0, 11, 2, "right"),        # 41: STYLE_NUM_ODD
        xf(3, 7, 6, 2, "right"),         # 42: STYLE_NUM_ERR
        xf(3, 8, 8, 2, "right"),         # 43: STYLE_NUM_WARN
    ]

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<fonts count="{len(fonts)}">{"".join(fonts)}</fonts>'
        f'<fills count="{len(fills)}">{"".join(fills)}</fills>'
        f'<borders count="{len(borders)}">{"".join(borders)}</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>'
        f'<cellXfs count="{len(cell_xfs)}">{"".join(cell_xfs)}</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '<dxfs count="0"/>'
        '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        '</styleSheet>'
    )


def make_str_cell(ref: str, style_id: int, text: str, sst: SharedStringTable) -> str:
    idx = sst.get_index(text)
    return f'<c r="{ref}" s="{style_id}" t="s"><v>{idx}</v></c>'


def make_num_cell(ref: str, style_id: int, val) -> str:
    return f'<c r="{ref}" s="{style_id}"><v>{val}</v></c>'


def build_detail_sheet_xml(bucket: LogBucket, sst: SharedStringTable, tab_selected: int = 0) -> str:
    headers = bucket.headers
    num_cols = len(headers)
    last_col_let = col_letter(num_cols)

    # カラム幅計算
    col_widths = []
    for c_idx, h_name in enumerate(headers):
        min_w = MIN_COL_WIDTHS.get(h_name, 14)
        max_w = MAX_COL_WIDTHS.get(h_name, 120)
        # AutoFilter ドロップダウン矢印分の余白 (+5)
        w = max(min_w, display_width(h_name) + 5)
        col_widths.append(min(max_w, w))

    for r_obj in bucket.rows:
        cols = r_obj["cols"]
        for c_idx in range(min(num_cols, len(cols))):
            h_name = headers[c_idx]
            max_w = MAX_COL_WIDTHS.get(h_name, 120)
            cw = display_width(cols[c_idx]) + 3
            if cw > col_widths[c_idx]:
                col_widths[c_idx] = min(max_w, cw)

    # メタデータ文字列の作成
    range_str = " / ".join(bucket.ranges) if bucket.ranges else "All Ranges (JST)"
    cnt_str = ", ".join(bucket.containers) if bucket.containers else "front|back"
    svc_str = ", ".join(bucket.services) if bucket.services else "intra-web|intra-api|inter-api|sf-api"
    th_part = f"   |   Thresholds: {bucket.thresholds}" if bucket.thresholds else ""
    meta_text = (
        f"Range (JST): {range_str}   |   Target: Container=[{cnt_str}]  "
        f"ECS-Service=[{svc_str}]  ScannedFiles={bucket.scanned_files}{th_part}"
    )

    max_p_part = f"   |   MaxPause={bucket.max_pause_ms:.1f}ms" if bucket.log_type == "gc.log" else ""
    summary_text = (
        f"Summary: Total={bucket.entry_count}  "
        f"[ ERROR={bucket.err_count} / WARN={bucket.warn_count} / INFO={bucket.info_count} ]   "
        f"(Output Detail Lines={len(bucket.rows)}){max_p_part}"
    )

    rows_xml = []
    merges = []

    # Row 1: Title Banner
    r1_cells = [make_str_cell("A1", STYLE_TITLE_BANNER, REPORT_TITLES[bucket.log_type], sst)]
    for c_i in range(2, num_cols + 1):
        r1_cells.append(make_str_cell(f"{col_letter(c_i)}1", STYLE_TITLE_BANNER, "", sst))
    rows_xml.append(f'<row r="1" ht="30" customHeight="1">{"".join(r1_cells)}</row>')
    merges.append(f"A1:{last_col_let}1")

    # Row 2: Metadata Bar
    r2_cells = [make_str_cell("A2", STYLE_META_BAR, meta_text, sst)]
    for c_i in range(2, num_cols + 1):
        r2_cells.append(make_str_cell(f"{col_letter(c_i)}2", STYLE_META_BAR, "", sst))
    rows_xml.append(f'<row r="2" ht="20" customHeight="1">{"".join(r2_cells)}</row>')
    merges.append(f"A2:{last_col_let}2")

    # Row 3: Summary Bar
    r3_cells = [make_str_cell("A3", STYLE_SUMMARY_BAR, summary_text, sst)]
    for c_i in range(2, num_cols + 1):
        r3_cells.append(make_str_cell(f"{col_letter(c_i)}3", STYLE_SUMMARY_BAR, "", sst))
    rows_xml.append(f'<row r="3" ht="20" customHeight="1">{"".join(r3_cells)}</row>')
    merges.append(f"A3:{last_col_let}3")

    # Row 4: Blank Spacer
    rows_xml.append('<row r="4" ht="8" customHeight="1"/>')

    # Row 5: Table Column Header Row (各出力項目をヘッダに設定)
    r5_cells = []
    for c_i, h_name in enumerate(headers, start=1):
        if c_i == num_cols:
            hdr_style = STYLE_TBL_HEADER_L
        elif h_name == "Pause":
            hdr_style = STYLE_TBL_HEADER_R
        else:
            hdr_style = STYLE_TBL_HEADER_C
        r5_cells.append(make_str_cell(f"{col_letter(c_i)}5", hdr_style, h_name, sst))
    rows_xml.append(f'<row r="5" ht="26" customHeight="1">{"".join(r5_cells)}</row>')

    # Rows 6+: データ行 (1行ずつわかりやすく出力)
    cur_row_num = 6
    if not bucket.rows:
        empty_cells = [
            make_str_cell("A6", STYLE_INFO_EVEN_C, "-", sst)
        ]
        for c_i in range(2, num_cols):
            empty_cells.append(make_str_cell(f"{col_letter(c_i)}6", STYLE_INFO_EVEN_C, "-", sst))
        empty_cells.append(
            make_str_cell(
                f"{last_col_let}6",
                STYLE_INFO_EVEN_L,
                "[!] 指定されたJST時間範囲に該当するログエントリはありません。",
                sst,
            )
        )
        rows_xml.append(f'<row r="6" ht="20" customHeight="1">{"".join(empty_cells)}</row>')
        cur_row_num = 6
    else:
        for r_idx, r_obj in enumerate(bucket.rows):
            cur_row_num = 6 + r_idx
            cols = r_obj["cols"]
            is_cont = r_obj["is_cont"]
            sev = r_obj["severity"]
            even_entry = (r_obj["entry_idx"] % 2 == 1)

            if is_cont:
                if sev == "ERROR":
                    st_c, st_l, st_r = STYLE_CONT_ERR_C, STYLE_CONT_ERR_L, STYLE_CONT_ERR_R
                elif sev == "WARN":
                    st_c, st_l, st_r = STYLE_CONT_WARN_C, STYLE_CONT_WARN_L, STYLE_CONT_WARN_R
                else:
                    st_c, st_l, st_r = STYLE_CONT_INFO_C, STYLE_CONT_INFO_L, STYLE_CONT_INFO_R
            else:
                if sev == "ERROR":
                    st_c, st_l, st_r = STYLE_ERR_C, STYLE_ERR_L, STYLE_ERR_R
                elif sev == "WARN":
                    st_c, st_l, st_r = STYLE_WARN_C, STYLE_WARN_L, STYLE_WARN_R
                elif sev == "DEBUG":
                    if even_entry:
                        st_c, st_l, st_r = STYLE_DBG_EVEN_C, STYLE_DBG_EVEN_L, STYLE_DBG_EVEN_R
                    else:
                        st_c, st_l, st_r = STYLE_DBG_ODD_C, STYLE_DBG_ODD_L, STYLE_DBG_ODD_R
                else:
                    if even_entry:
                        st_c, st_l, st_r = STYLE_INFO_EVEN_C, STYLE_INFO_EVEN_L, STYLE_INFO_EVEN_R
                    else:
                        st_c, st_l, st_r = STYLE_INFO_ODD_C, STYLE_INFO_ODD_L, STYLE_INFO_ODD_R

            row_cells = []
            for c_i in range(1, num_cols + 1):
                val = cols[c_i - 1] if (c_i - 1) < len(cols) else ""
                h_name = headers[c_i - 1]
                cell_ref = f"{col_letter(c_i)}{cur_row_num}"

                if c_i == num_cols:
                    c_style = st_l
                elif h_name == "Level" and not is_cont:
                    if sev == "ERROR":
                        c_style = STYLE_ERR_BADGE_C
                    elif sev == "WARN":
                        c_style = STYLE_WARN_BADGE_C
                    else:
                        c_style = st_c
                elif h_name == "Sts" and not is_cont:
                    if sev == "ERROR" or str(val).startswith("5"):
                        c_style = STYLE_ERR_BADGE_C
                    elif sev == "WARN" or str(val).startswith("4"):
                        c_style = STYLE_WARN_BADGE_C
                    else:
                        c_style = st_c
                elif h_name == "Pause":
                    if not is_cont and val and val != "-":
                        if sev == "ERROR":
                            c_style = STYLE_ERR_BADGE_R
                        elif sev == "WARN":
                            c_style = STYLE_WARN_BADGE_R
                        else:
                            c_style = st_r
                    else:
                        c_style = st_c if val == "-" else st_r
                else:
                    c_style = st_c

                row_cells.append(make_str_cell(cell_ref, c_style, val, sst))

            rows_xml.append(f'<row r="{cur_row_num}" ht="20" customHeight="1">{"".join(row_cells)}</row>')

    cols_xml_parts = ["<cols>"]
    for c_i, w in enumerate(col_widths, start=1):
        cols_xml_parts.append(f'<col min="{c_i}" max="{c_i}" width="{w}" customWidth="1"/>')
    cols_xml_parts.append("</cols>")

    merge_xml = (
        f'<mergeCells count="{len(merges)}">'
        + "".join(f'<mergeCell ref="{m}"/>' for m in merges)
        + "</mergeCells>"
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<dimension ref="A1:{last_col_let}{cur_row_num}"/>'
        f'<sheetViews><sheetView tabSelected="{tab_selected}" showGridLines="1" workbookViewId="0">'
        '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A6" sqref="A6"/>'
        '</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>'
        f'{"".join(cols_xml_parts)}'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        f'<autoFilter ref="A5:{last_col_let}{cur_row_num}"/>'
        f'{merge_xml}'
        '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
        '</worksheet>'
    )


def build_summary_sheet_xml(buckets: dict, source_files: list, sst: SharedStringTable, tab_selected: int = 1) -> str:
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    src_names = ", ".join(
        "stdin" if os.path.basename(p).startswith("jboss_teraterm_stdin.") else os.path.basename(p)
        for p in source_files
    )
    if len(src_names) > 140:
        src_names = src_names[:137] + "..."

    rows_xml = []
    merges = []
    num_cols = 9
    last_col_let = col_letter(num_cols)

    # Row 1: Title Banner
    r1 = [make_str_cell("A1", STYLE_TITLE_BANNER, "JBoss EAP ログ解析レポート サマリー (server.log / access_log.log / gc.log)", sst)]
    for c_i in range(2, num_cols + 1):
        r1.append(make_str_cell(f"{col_letter(c_i)}1", STYLE_TITLE_BANNER, "", sst))
    rows_xml.append(f'<row r="1" ht="30" customHeight="1">{"".join(r1)}</row>')
    merges.append(f"A1:{last_col_let}1")

    # Row 2: Subtitle Bar
    sub_text = f"Generated: {now_str} (JST)   |   Source Log Files: {src_names}"
    r2 = [make_str_cell("A2", STYLE_SUBTITLE_BAR, sub_text, sst)]
    for c_i in range(2, num_cols + 1):
        r2.append(make_str_cell(f"{col_letter(c_i)}2", STYLE_SUBTITLE_BAR, "", sst))
    rows_xml.append(f'<row r="2" ht="20" customHeight="1">{"".join(r2)}</row>')
    merges.append(f"A2:{last_col_let}2")

    # Row 3: Spacer
    rows_xml.append('<row r="3" ht="10" customHeight="1"/>')

    # Row 4: Section 1 Header
    r4 = [make_str_cell("A4", STYLE_SECTION_HDR, "■ 1. ログ種別ごと集計サマリー (Summary by Log Type)", sst)]
    for c_i in range(2, num_cols + 1):
        r4.append(make_str_cell(f"{col_letter(c_i)}4", STYLE_SECTION_HDR, "", sst))
    rows_xml.append(f'<row r="4" ht="24" customHeight="1">{"".join(r4)}</row>')
    merges.append(f"A4:{last_col_let}4")

    # Row 5: Summary Table Header
    sec1_headers = [
        "ログ種別 (Sheet)",
        "抽出対象期間 (JST Range)",
        "対象コンテナ",
        "対象ECSサービス",
        "総イベント数",
        "出力行数 (含スタックトレース)",
        "ERROR 件数",
        "WARN 件数",
        "INFO/他 件数 / 最大GC停止",
    ]
    r5 = []
    for c_i, h in enumerate(sec1_headers, start=1):
        r5.append(make_str_cell(f"{col_letter(c_i)}5", STYLE_TBL_HEADER_C, h, sst))
    rows_xml.append(f'<row r="5" ht="25" customHeight="1">{"".join(r5)}</row>')

    cur_r = 6
    for idx, lt in enumerate(LOG_TYPES_ORDER):
        b = buckets[lt]
        even = (idx % 2 == 0)
        st_c = STYLE_INFO_EVEN_C if even else STYLE_INFO_ODD_C
        st_l = STYLE_INFO_EVEN_L if even else STYLE_INFO_ODD_L
        st_num = STYLE_NUM_EVEN if even else STYLE_NUM_ODD
        st_err = STYLE_NUM_ERR if b.err_count > 0 else st_num
        st_warn = STYLE_NUM_WARN if b.warn_count > 0 else st_num

        rng_s = " / ".join(b.ranges) if b.ranges else ("-" if not b.seen_section else "All (JST)")
        cnt_s = ", ".join(b.containers) if b.containers else ("-" if not b.seen_section else "front|back")
        svc_s = ", ".join(b.services) if b.services else ("-" if not b.seen_section else "all")
        extra_s = f"{b.info_count:,} 件"
        if lt == "gc.log" and b.seen_section:
            extra_s += f" (MaxPause: {b.max_pause_ms:.1f}ms)"

        cells = [
            make_str_cell(f"A{cur_r}", st_c, lt, sst),
            make_str_cell(f"B{cur_r}", st_c, rng_s, sst),
            make_str_cell(f"C{cur_r}", st_c, cnt_s, sst),
            make_str_cell(f"D{cur_r}", st_l, svc_s, sst),
            make_num_cell(f"E{cur_r}", st_num, b.entry_count),
            make_num_cell(f"F{cur_r}", st_num, len(b.rows)),
            make_num_cell(f"G{cur_r}", st_err, b.err_count),
            make_num_cell(f"H{cur_r}", st_warn, b.warn_count),
            make_str_cell(f"I{cur_r}", st_c, extra_s, sst),
        ]
        rows_xml.append(f'<row r="{cur_r}" ht="21" customHeight="1">{"".join(cells)}</row>')
        cur_r += 1

    # Spacer
    rows_xml.append(f'<row r="{cur_r}" ht="12" customHeight="1"/>')
    cur_r += 1

    # Section 2 Header: Container / ECS Service Breakdown
    r_sec2 = [make_str_cell(f"A{cur_r}", STYLE_SECTION_HDR, "■ 2. コンテナ / ECSサービス別 内訳 (Breakdown by Container / ECS Service)", sst)]
    for c_i in range(2, num_cols + 1):
        r_sec2.append(make_str_cell(f"{col_letter(c_i)}{cur_r}", STYLE_SECTION_HDR, "", sst))
    rows_xml.append(f'<row r="{cur_r}" ht="24" customHeight="1">{"".join(r_sec2)}</row>')
    merges.append(f"A{cur_r}:{last_col_let}{cur_r}")
    cur_r += 1

    sec2_headers = [
        "ログ種別 (Sheet)",
        "Cntnr (コンテナ)",
        "ECS Svc (サービス)",
        "Total 件数",
        "ERROR 件数",
        "WARN 件数",
        "INFO/他 件数",
        "最大GC停止時間 (MaxPause)",
        "ステータス判定",
    ]
    r_s2h = []
    for c_i, h in enumerate(sec2_headers, start=1):
        r_s2h.append(make_str_cell(f"{col_letter(c_i)}{cur_r}", STYLE_TBL_HEADER_C, h, sst))
    rows_xml.append(f'<row r="{cur_r}" ht="25" customHeight="1">{"".join(r_s2h)}</row>')
    cur_r += 1

    bd_idx = 0
    for lt in LOG_TYPES_ORDER:
        b = buckets[lt]
        for (cnt_k, svc_k) in sorted(b.cs_stats.keys()):
            st_data = b.cs_stats[(cnt_k, svc_k)]
            even = (bd_idx % 2 == 0)
            bd_idx += 1
            st_c = STYLE_INFO_EVEN_C if even else STYLE_INFO_ODD_C
            st_num = STYLE_NUM_EVEN if even else STYLE_NUM_ODD
            st_err = STYLE_NUM_ERR if st_data["err"] > 0 else st_num
            st_warn = STYLE_NUM_WARN if st_data["warn"] > 0 else st_num

            if st_data["err"] > 0:
                status_label = "ERROR 検出"
                st_badge = STYLE_ERR_BADGE_C
            elif st_data["warn"] > 0:
                status_label = "WARN 検出"
                st_badge = STYLE_WARN_BADGE_C
            else:
                status_label = "NORMAL"
                st_badge = st_c

            mp_str = f"{st_data['max_pause']:.1f}ms" if lt == "gc.log" and st_data["max_pause"] > 0 else "-"

            cells = [
                make_str_cell(f"A{cur_r}", st_c, lt, sst),
                make_str_cell(f"B{cur_r}", st_c, cnt_k, sst),
                make_str_cell(f"C{cur_r}", st_c, svc_k, sst),
                make_num_cell(f"D{cur_r}", st_num, st_data["total"]),
                make_num_cell(f"E{cur_r}", st_err, st_data["err"]),
                make_num_cell(f"F{cur_r}", st_warn, st_data["warn"]),
                make_num_cell(f"G{cur_r}", st_num, st_data["info"]),
                make_str_cell(f"H{cur_r}", st_c, mp_str, sst),
                make_str_cell(f"I{cur_r}", st_badge, status_label, sst),
            ]
            rows_xml.append(f'<row r="{cur_r}" ht="20" customHeight="1">{"".join(cells)}</row>')
            cur_r += 1

    if bd_idx == 0:
        cells = [make_str_cell(f"A{cur_r}", STYLE_INFO_EVEN_C, "-", sst) for _ in range(num_cols)]
        rows_xml.append(f'<row r="{cur_r}" ht="20" customHeight="1">{"".join(cells)}</row>')
        cur_r += 1

    col_widths = [20, 44, 18, 36, 16, 26, 16, 24, 30]
    cols_xml = "<cols>" + "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(col_widths, start=1)
    ) + "</cols>"

    merge_xml = (
        f'<mergeCells count="{len(merges)}">'
        + "".join(f'<mergeCell ref="{m}"/>' for m in merges)
        + "</mergeCells>"
    )

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<dimension ref="A1:{last_col_let}{cur_r - 1}"/>'
        f'<sheetViews><sheetView tabSelected="{tab_selected}" showGridLines="1" workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>'
        f'{cols_xml}'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        f'{merge_xml}'
        '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
        '</worksheet>'
    )


def write_xlsx_report(output_path: str, buckets: dict, source_files: list):
    sst = SharedStringTable()

    # 出力対象シートの構成: Summary + 検出されたログ種別シート (または全3種別)
    active_types = [lt for lt in LOG_TYPES_ORDER if buckets[lt].seen_section or buckets[lt].rows]
    if not active_types:
        active_types = list(LOG_TYPES_ORDER)

    sheets_info = []  # (sheet_name, xml_content)
    summary_xml = build_summary_sheet_xml(buckets, source_files, sst, tab_selected=1)
    sheets_info.append(("Summary", summary_xml))

    for lt in active_types:
        detail_xml = build_detail_sheet_xml(buckets[lt], sst, tab_selected=0)
        sheets_info.append((lt, detail_xml))

    # [Content_Types].xml
    ct_overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>',
    ]
    for idx in range(1, len(sheets_info) + 1):
        ct_overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{idx}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'{"".join(ct_overrides)}'
        '</Types>'
    )

    # _rels/.rels
    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    # xl/workbook.xml & xl/_rels/workbook.xml.rels
    wb_sheets = []
    wb_rels = []
    for idx, (s_name, _) in enumerate(sheets_info, start=1):
        esc_name = sanitize_xml_text(s_name)
        wb_sheets.append(f'<sheet name="{esc_name}" sheetId="{idx}" r:id="rId{idx}"/>')
        wb_rels.append(
            f'<Relationship Id="rId{idx}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{idx}.xml"/>'
        )

    style_rid = len(sheets_info) + 1
    sst_rid = len(sheets_info) + 2
    wb_rels.append(
        f'<Relationship Id="rId{style_rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    wb_rels.append(
        f'<Relationship Id="rId{sst_rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
        'Target="sharedStrings.xml"/>'
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16200"/></bookViews>'
        f'<sheets>{"".join(wb_sheets)}</sheets>'
        '</workbook>'
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(wb_rels)}'
        '</Relationships>'
    )

    styles_xml = build_styles_xml()
    sst_xml = sst.to_xml()

    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml.encode("utf-8"))
        zf.writestr("_rels/.rels", root_rels_xml.encode("utf-8"))
        zf.writestr("xl/workbook.xml", workbook_xml.encode("utf-8"))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml.encode("utf-8"))
        zf.writestr("xl/styles.xml", styles_xml.encode("utf-8"))
        zf.writestr("xl/sharedStrings.xml", sst_xml.encode("utf-8"))
        for idx, (_, s_xml) in enumerate(sheets_info, start=1):
            zf.writestr(f"xl/worksheets/sheet{idx}.xml", s_xml.encode("utf-8"))

    return sheets_info


def main():
    output_path = sys.argv[1]
    raw_inputs = sys.argv[2:]

    resolved_files = []
    for item in raw_inputs:
        if os.path.isdir(item):
            for entry in sorted(os.listdir(item)):
                full_p = os.path.join(item, entry)
                if os.path.isfile(full_p) and entry.lower().endswith((".log", ".txt")):
                    resolved_files.append(full_p)
        elif os.path.isfile(item):
            resolved_files.append(item)
        else:
            print(f"[WARN] 指定されたパスが見つかりません: {item}", file=sys.stderr)

    if not resolved_files:
        print("[ERROR] 有効な入力ログファイルが見つかりませんでした。", file=sys.stderr)
        sys.exit(1)

    buckets = parse_teraterm_files(resolved_files)
    sheets_info = write_xlsx_report(output_path, buckets, resolved_files)

    print("================================================================================")
    print(" JBoss EAP Log Excel Report Generator (Meiryo UI / Monotone)")
    print("================================================================================")
    print(f" Output Excel : {os.path.abspath(output_path)}")
    print(f" Source Files : {len(resolved_files)} file(s)")
    for lt in LOG_TYPES_ORDER:
        b = buckets[lt]
        if b.seen_section or b.rows:
            print(
                f"   - {lt:<14} : Events={b.entry_count:<5} Lines={len(b.rows):<5} "
                f"(ERROR={b.err_count}, WARN={b.warn_count}, INFO={b.info_count})"
            )
    print("================================================================================")


if __name__ == "__main__":
    main()
PYEOF
