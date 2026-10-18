#!/usr/bin/env bash
# ==============================================================================
# RHEL 9.8 JBoss EAP アクセスログ横断検索・JST時系列ソート・カラー整形出力コマンド
# 対象パス:
#   /mnt/logs/<コンテナ>/logs/<ECSサービス>/mid/<タスクID(32文字) または current>/access_log.log
#   /mnt/logs/<コンテナ>/logs/<ECSサービス>/mid/<タスクID(32文字) または current>/access_log.YYYY-MM-DD.log
# ==============================================================================

export TZ=Asia/Tokyo; \
BASE_DIR="${BASE_DIR:-/mnt/logs}"; \
CONTAINERS="${CONTAINERS:-front|back}"; \
SERVICES="${SERVICES:-intra-web|intra-api|inter-api|sf-api}"; \
START_JST="${START_JST:-2026-08-20 00:00:00}"; \
END_JST="${END_JST:-2026-08-20 23:59:59}"; \
LOG_TZ="${LOG_TZ:-JST}"; \
LEVEL_FILTER="${LEVEL_FILTER:-}"; \
STATUS_FILTER="${STATUS_FILTER:-}"; \
KEYWORD_FILTER="${KEYWORD_FILTER:-}"; \
find -L "${BASE_DIR%/}" -mindepth 6 -maxdepth 6 -type f \
  -regextype posix-extended \
  -regex "^${BASE_DIR%/}/(${CONTAINERS})/logs/(${SERVICES})/mid/(current|[0-9a-zA-Z]{32})/access_log(\.[0-9]{4}-[0-9]{2}-[0-9]{2})?\.log(\.[0-9]{4}-[0-9]{2}-[0-9]{2})?$" \
  -exec stat -L -c "%d:%i|%Y|%n" {} + 2>/dev/null \
| LC_ALL=C gawk \
  -v start_jst="$START_JST" \
  -v end_jst="$END_JST" \
  -v containers="$CONTAINERS" \
  -v services="$SERVICES" \
  -v log_tz="$LOG_TZ" \
  -v level_filter="$LEVEL_FILTER" \
  -v status_filter="$STATUS_FILTER" \
  -v keyword_filter="$KEYWORD_FILTER" \
'function norm_ts(ts, is_end,   s, a, Y, M, D, h, m, sec, ms) { \
  s = ts; \
  gsub(/^[[:space:]]+|[[:space:]]+$/, "", s); \
  gsub(/\//, "-", s); \
  sub(/T/, " ", s); \
  sub(/\./, ",", s); \
  if (s == "") return is_end ? "9999-12-31 23:59:59,999" : "0000-01-01 00:00:00,000"; \
  if (match(s, /^([0-9]{4})-([0-9]{1,2})-([0-9]{1,2})([[:space:]]+([0-9]{1,2}):([0-9]{1,2})(:([0-9]{1,2})(,([0-9]{1,6}))?)?)?$/, a)) { \
    Y = a[1]+0; M = a[2]+0; D = a[3]+0; \
    if (a[5] == "") { \
      h = is_end ? 23 : 0; m = is_end ? 59 : 0; sec = is_end ? 59 : 0; ms = is_end ? "999" : "000"; \
    } else if (a[8] == "") { \
      h = a[5]+0; m = a[6]+0; sec = is_end ? 59 : 0; ms = is_end ? "999" : "000"; \
    } else if (a[10] == "") { \
      h = a[5]+0; m = a[6]+0; sec = a[8]+0; ms = is_end ? "999" : "000"; \
    } else { \
      h = a[5]+0; m = a[6]+0; sec = a[8]+0; ms = substr(a[10] "000", 1, 3); \
    } \
    return sprintf("%04d-%02d-%02d %02d:%02d:%02d,%s", Y, M, D, h, m, sec, ms); \
  } \
  return s; \
} \
function date_shift(d_str, days,   dt, ep) { \
  dt = d_str " 12 00 00"; \
  gsub(/-/, " ", dt); \
  ep = mktime(dt); \
  return (ep == -1) ? d_str : strftime("%Y-%m-%d", ep + (days * 86400)); \
} \
function to_jst(d_str, t_str, ms_str, tz_str,   ms, dt, ep, sgn, tz_off, tzm) { \
  ms = substr(ms_str "000", 1, 3); \
  if (tz_str == "+0900" || tz_str == "+09:00" || (tz_str == "" && toupper(log_tz) != "UTC")) { \
    return d_str " " t_str "," ms; \
  } \
  dt = d_str " " t_str; \
  gsub(/[-:]/, " ", dt); \
  ep = mktime(dt); \
  if (ep == -1) return d_str " " t_str "," ms; \
  if (tz_str == "Z" || tz_str == "+00:00" || tz_str == "+0000" || (tz_str == "" && toupper(log_tz) == "UTC")) { \
    tz_off = 0; \
  } else if (match(tz_str, /^([+-])([0-9]{2}):?([0-9]{2})$/, tzm)) { \
    sgn = (tzm[1] == "-") ? -1 : 1; \
    tz_off = sgn * ((tzm[2] + 0) * 3600 + (tzm[3] + 0) * 60); \
  } else { \
    return d_str " " t_str "," ms; \
  } \
  return strftime("%Y-%m-%d %H:%M:%S", ep + (32400 - tz_off)) "," ms; \
} \
function parse_access_details(prefix, suffix, raw_line,   cpart, rest, rm, ip, usr, req, sts, tail, lvl) { \
  cpart = prefix; \
  rest  = suffix; \
  gsub(/^[[:space:]]+|[[:space:]]+$/, "", cpart); \
  gsub(/^[[:space:]]+|[[:space:]]+$/, "", rest); \
  ip = "-"; usr = ""; req = ""; sts = "---"; tail = ""; lvl = "INFO"; \
  if (match(rest, /^([^"]*)[[:space:]]*("([^"\\]|\\.)*")[[:space:]]+([1-5][0-9]{2})([[:space:]]+(.*))?$/, rm)) { \
    if (cpart == "" && rm[1] != "") { \
      cpart = rm[1]; \
      gsub(/^[[:space:]]+|[[:space:]]+$/, "", cpart); \
    } \
    req  = rm[2]; \
    sts  = rm[4]; \
    tail = rm[6]; \
  } else if (match(rest, /^([^A-Z]*)(GET|POST|PUT|DELETE|PATCH|HEAD|OPTIONS|TRACE|CONNECT)[[:space:]]+([^[:space:]]+)([[:space:]]+HTTP\/[0-9.]+)?[[:space:]]+([1-5][0-9]{2})([[:space:]]+(.*))?$/, rm)) { \
    if (cpart == "" && rm[1] != "") { \
      cpart = rm[1]; \
      gsub(/^[[:space:]]+|[[:space:]]+$/, "", cpart); \
    } \
    req  = "\"" rm[2] " " rm[3] rm[4] "\""; \
    sts  = rm[5]; \
    tail = rm[7]; \
  } else { \
    req = (cpart != "" && rest != "") ? (cpart " " rest) : (rest != "" ? rest : raw_line); \
    cpart = ""; \
  } \
  if (cpart != "") { \
    if (match(cpart, /^(.+)[[:space:]]+-[[:space:]]+([^[:space:]]+)$/, cm)) { \
      ip  = cm[1]; \
      usr = (cm[2] != "-") ? cm[2] : ""; \
    } else { \
      ip = cpart; \
    } \
    gsub(/^[[:space:]]+|[[:space:]]+$/, "", ip); \
    if (index(ip, ",") > 0) { \
      split(ip, ipa, /,[[:space:]]*/); \
      ip = ipa[1]; \
    } \
  } \
  if (sts ~ /^5[0-9]{2}$/ || raw_line ~ /(^|[^A-Za-z])(ERROR|FATAL|SEVERE|Exception)([^A-Za-z]|$)/) { \
    lvl = "ERROR"; \
  } else if (sts ~ /^4[0-9]{2}$/ || raw_line ~ /(^|[^A-Za-z])(WARN|WARNING)([^A-Za-z]|$)/) { \
    lvl = "WARN"; \
  } else { \
    lvl = "INFO"; \
  } \
  cur_ip   = (ip != "") ? ip : "-"; \
  cur_sts  = sts; \
  cur_lvl  = lvl; \
  cur_body = (usr != "" ? "[user:" usr "] " : "") req (tail != "" ? " " tail : ""); \
} \
function flush_rec(   key, cs_key) { \
  if (cur_ts == "") return; \
  if (cur_ts >= n_start && cur_ts <= n_end) { \
    if (level_filter == "" || cur_lvl ~ "^(" level_filter ")$") { \
      if (status_filter == "" || cur_sts ~ "^(" status_filter ")$") { \
        if (keyword_filter == "" || cur_raw ~ keyword_filter || cur_lvl ~ keyword_filter || cur_sts ~ keyword_filter) { \
          key = sprintf("%s\034%09d", cur_ts, ++seq); \
          records[key] = cur_cnt "\034" cur_svc "\034" cur_tsk "\034" cur_lvl "\034" cur_sts "\034" cur_ip "\034" cur_ts "\034" cur_body; \
          total_cnt++; \
          cs_key = sprintf("%-5s / %-9s", cur_cnt, cur_svc); \
          cs_total[cs_key]++; \
          if (cur_lvl == "ERROR") { err_cnt++; cs_err[cs_key]++; } \
          else if (cur_lvl == "WARN") { warn_cnt++; cs_warn[cs_key]++; } \
          else { info_cnt++; cs_info[cs_key]++; } \
        } \
      } \
    } \
  } \
  cur_ts = ""; \
} \
BEGIN { \
  split("JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC", mnames, " "); \
  for (i = 1; i <= 12; i++) mon_map[mnames[i]] = sprintf("%02d", i); \
  n_start  = norm_ts(start_jst, 0); \
  n_end    = norm_ts(end_jst, 1); \
  min_date = date_shift(substr(n_start, 1, 10), -1); \
  max_date = date_shift(substr(n_end, 1, 10), 1); \
  C_RST    = "\033[0m"; \
  C_BOLD   = "\033[1m"; \
  C_DIM    = "\033[0;90m"; \
  C_CYN    = "\033[0;36m"; \
  C_BCYN   = "\033[1;36m"; \
  C_BMAG   = "\033[1;35m"; \
  C_BBLU   = "\033[1;34m"; \
  C_BGRN   = "\033[1;32m"; \
  C_ERR_BG = "\033[1;97;41m"; \
  C_ERR_FG = "\033[1;91m"; \
  C_ERR_ST = "\033[0;91m"; \
  C_WRN_BG = "\033[1;30;43m"; \
  C_WRN_FG = "\033[1;93m"; \
  C_WRN_ST = "\033[0;33m"; \
  C_INF_FG = "\033[1;32m"; \
} \
{ \
  split($0, meta, "|"); \
  ino = meta[1]; mtime = meta[2]; fpath = meta[3]; \
  if (match(fpath, /\/([^\/]+)\/logs\/([^\/]+)\/mid\/([^\/]+)\/(access_log(\.([0-9]{4}-[0-9]{2}-[0-9]{2}))?\.log(\.([0-9]{4}-[0-9]{2}-[0-9]{2}))?)$/, m)) { \
    cnt = m[1]; svc = m[2]; tsk = m[3]; fname = m[4]; \
    rdate = (m[6] != "") ? m[6] : m[8]; \
    if (rdate != "" && (rdate < min_date || rdate > max_date)) next; \
    if (!(ino in seen_ino) || (seen_tsk[ino] == "current" && tsk != "current")) { \
      seen_ino[ino]   = fpath; \
      seen_mtime[ino] = mtime; \
      seen_cnt[ino]   = cnt; \
      seen_svc[ino]   = svc; \
      seen_tsk[ino]   = tsk; \
      seen_rdate[ino] = rdate; \
    } \
  } \
} \
END { \
  PROCINFO["sorted_in"] = "@ind_str_asc"; \
  for (ino in seen_ino) { \
    fpath = seen_ino[ino]; \
    mtime = seen_mtime[ino]; \
    cnt   = seen_cnt[ino]; \
    svc   = seen_svc[ino]; \
    tsk   = (seen_tsk[ino] == "current") ? "current" : substr(seen_tsk[ino], 1, 8); \
    rdate = seen_rdate[ino]; \
    fdate = (rdate != "") ? rdate : strftime("%Y-%m-%d", mtime); \
    files_scanned++; \
    cur_ts = ""; cur_cnt = cnt; cur_svc = svc; cur_tsk = tsk; \
    while ((getline line < fpath) > 0) { \
      sub(/\r$/, "", line); \
      if (match(line, /\[([0-9]{1,2})\/([A-Za-z]{3})\/([0-9]{4}):([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?[[:space:]]*([+-][0-9]{2}:?[0-9]{2}|Z)?\]/, tm)) { \
        flush_rec(); \
        p_str = substr(line, 1, RSTART - 1); \
        s_str = substr(line, RSTART + RLENGTH); \
        m_num = mon_map[toupper(tm[2])]; \
        if (m_num == "") m_num = "01"; \
        d_s = sprintf("%04d-%02d-%02d", tm[3]+0, m_num+0, tm[1]+0); \
        cur_ts  = to_jst(d_s, tm[4], tm[6], tm[7]); \
        cur_raw = line; \
        parse_access_details(p_str, s_str, line); \
      } else if (match(line, /\[([0-9]{4}[-\/][0-9]{2}[-\/][0-9]{2})[T :]([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?[[:space:]]*(Z|[+-][0-9]{2}:?[0-9]{2})?\]/, tm)) { \
        flush_rec(); \
        p_str = substr(line, 1, RSTART - 1); \
        s_str = substr(line, RSTART + RLENGTH); \
        d_s = tm[1]; gsub(/\//, "-", d_s); \
        cur_ts  = to_jst(d_s, tm[2], tm[4], tm[5]); \
        cur_raw = line; \
        parse_access_details(p_str, s_str, line); \
      } else if (match(line, /^([0-9]{4}[-\/][0-9]{2}[-\/][0-9]{2})[T ]([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?(Z|[+-][0-9]{2}:?[0-9]{2})?[[:space:]]+(.*)$/, tm)) { \
        flush_rec(); \
        d_s = tm[1]; gsub(/\//, "-", d_s); \
        cur_ts  = to_jst(d_s, tm[2], tm[4], tm[5]); \
        cur_raw = line; \
        parse_access_details("", tm[6], line); \
      } else if (match(line, /^([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?[[:space:]]+(.*)$/, tm)) { \
        flush_rec(); \
        cur_ts  = to_jst(fdate, tm[1], tm[3], ""); \
        cur_raw = line; \
        parse_access_details("", tm[4], line); \
      } else if (cur_ts != "") { \
        cur_body = cur_body "\036" line; \
        cur_raw  = cur_raw "\n" line; \
      } \
    } \
    flush_rec(); \
    close(fpath); \
  } \
  sep = "========================================================================================================================"; \
  div = "------------------------+-------+-----------+----------+---------+-----+-----------------+------------------------------"; \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
  printf " %sJBoss EAP Access Log Viewer (JST Sorted)%s | Range: %s%s%s to %s%s%s (JST)\n", C_BOLD, C_RST, C_BCYN, n_start, C_RST, C_BCYN, n_end, C_RST; \
  printf " Target: Container=[%s]  ECS-Service=[%s]  ScannedFiles=%d\n", containers, services, files_scanned+0; \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
  printf "%s%-23s | %-5s | %-9s | %-8s | %-7s | %-3s | %-15s | %s%s\n", C_BOLD, "JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Sts", "Client IP", "Request / Access Log Details", C_RST; \
  printf "%s%s%s\n", C_DIM, div, C_RST; \
  if (total_cnt + 0 == 0) { \
    printf " %s[!] No matching access log entries found in the specified JST time range.%s\n", C_WRN_FG, C_RST; \
  } \
  for (k in records) { \
    split(records[k], r, "\034"); \
    cnt = r[1]; svc = r[2]; tsk = r[3]; lvl = r[4]; sts = r[5]; ip = r[6]; ts = r[7]; body = r[8]; \
    n_lines = split(body, blines, "\036"); \
    msg1 = blines[1]; \
    if (lvl == "ERROR") { \
      b_lvl = sprintf("%s %-5s %s", C_ERR_BG, "ERROR", C_RST); \
      b_sts = sprintf("%s%-3s%s", C_ERR_BG, sts, C_RST); \
      printf "%s%-23s%s | %s%-5s%s | %s%-9s%s | %s%-8s%s | %s | %s | %s%-15s%s | %s%s%s\n", C_ERR_FG, ts, C_RST, C_ERR_FG, cnt, C_RST, C_ERR_FG, svc, C_RST, C_ERR_FG, tsk, C_RST, b_lvl, b_sts, C_ERR_FG, ip, C_RST, C_ERR_FG, msg1, C_RST; \
      for (i = 2; i <= n_lines; i++) { \
        st_color = (blines[i] ~ /(Caused by:|Exception|Error)/) ? C_ERR_FG : C_ERR_ST; \
        printf "%s%-23s | %-5s | %-9s | %-8s |   ->    |     | %-15s |%s %s%s%s\n", C_DIM, "", cnt, svc, "", "", C_RST, st_color, blines[i], C_RST; \
      } \
    } else if (lvl == "WARN") { \
      b_lvl = sprintf("%s %-5s %s", C_WRN_BG, "WARN", C_RST); \
      b_sts = sprintf("%s%-3s%s", C_WRN_BG, sts, C_RST); \
      printf "%s%-23s%s | %s%-5s%s | %s%-9s%s | %s%-8s%s | %s | %s | %s%-15s%s | %s%s%s\n", C_WRN_FG, ts, C_RST, C_WRN_FG, cnt, C_RST, C_WRN_FG, svc, C_RST, C_WRN_FG, tsk, C_RST, b_lvl, b_sts, C_WRN_FG, ip, C_RST, C_WRN_FG, msg1, C_RST; \
      for (i = 2; i <= n_lines; i++) { \
        printf "%s%-23s | %-5s | %-9s | %-8s |   ->    |     | %-15s |%s %s%s%s\n", C_DIM, "", cnt, svc, "", "", C_RST, C_WRN_ST, blines[i], C_RST; \
      } \
    } else { \
      c_cnt = (cnt == "front") ? C_BCYN : C_BMAG; \
      c_svc = (svc == "intra-web") ? C_BGRN : (svc == "intra-api") ? C_BBLU : (svc == "inter-api") ? C_BCYN : C_BMAG; \
      c_sts = (sts ~ /^3/) ? C_BCYN : C_BGRN; \
      b_lvl = sprintf("%s %-5s %s", C_INF_FG, "INFO", C_RST); \
      b_sts = sprintf("%s%-3s%s", c_sts, sts, C_RST); \
      printf "%s%-23s%s | %s%-5s%s | %s%-9s%s | %s%-8s%s | %s | %s | %-15s | %s\n", C_CYN, ts, C_RST, c_cnt, cnt, C_RST, c_svc, svc, C_RST, C_DIM, tsk, C_RST, b_lvl, b_sts, ip, msg1; \
      for (i = 2; i <= n_lines; i++) { \
        printf "%s%-23s | %-5s | %-9s | %-8s |   ->    |     | %-15s | %s%s\n", C_DIM, "", cnt, svc, "", "", blines[i], C_RST; \
      } \
    } \
  } \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
  printf " %sSummary:%s Total=%s%d%s  [ %sERROR(5xx)=%d%s  %sWARN(4xx)=%d%s  %sINFO(2xx/3xx)=%d%s ]\n", C_BOLD, C_RST, C_BOLD, total_cnt+0, C_RST, C_ERR_FG, err_cnt+0, C_RST, C_WRN_FG, warn_cnt+0, C_RST, C_BGRN, info_cnt+0, C_RST; \
  if (total_cnt + 0 > 0) { \
    printf " %sBreakdown by Container / ECS Service:%s\n", C_BOLD, C_RST; \
    for (cs in cs_total) { \
      printf "   - %-17s : Total=%-4d ( %sERROR=%d%s, %sWARN=%d%s, %sINFO=%d%s )\n", cs, cs_total[cs], C_ERR_FG, cs_err[cs]+0, C_RST, C_WRN_FG, cs_warn[cs]+0, C_RST, C_BGRN, cs_info[cs]+0, C_RST; \
    } \
  } \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
}'
