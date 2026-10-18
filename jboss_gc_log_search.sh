#!/usr/bin/env bash
# ==============================================================================
# RHEL 9.8 JBoss EAP GCログ横断検索・JST時系列ソート・カラー整形出力コマンド
# 対象パス:
#   /mnt/logs/<コンテナ>/logs/<ECSサービス>/mid/<タスクID(32文字) または current>/gc.log
#   /mnt/logs/<コンテナ>/logs/<ECSサービス>/mid/<タスクID(32文字) または current>/gc.log.<連番>
# ==============================================================================

export TZ=Asia/Tokyo; \
BASE_DIR="${BASE_DIR:-/mnt/logs}"; \
CONTAINERS="${CONTAINERS:-front|back}"; \
SERVICES="${SERVICES:-intra-web|intra-api|inter-api|sf-api}"; \
START_JST="${START_JST:-2026-08-20 00:00:00}"; \
END_JST="${END_JST:-2026-08-20 23:59:59}"; \
LOG_TZ="${LOG_TZ:-JST}"; \
WARN_PAUSE_MS="${WARN_PAUSE_MS:-200}"; \
ERR_PAUSE_MS="${ERR_PAUSE_MS:-1000}"; \
LEVEL_FILTER="${LEVEL_FILTER:-}"; \
KEYWORD_FILTER="${KEYWORD_FILTER:-}"; \
find -L "${BASE_DIR%/}" -mindepth 6 -maxdepth 6 -type f \
  -regextype posix-extended \
  -regex "^${BASE_DIR%/}/(${CONTAINERS})/logs/(${SERVICES})/mid/(current|[0-9a-zA-Z]{32})/gc\.log(\.[0-9]+(\.current)?)?$" \
  -exec stat -L -c "%d:%i|%Y|%n" {} + 2>/dev/null \
| LC_ALL=C gawk \
  -v start_jst="$START_JST" \
  -v end_jst="$END_JST" \
  -v containers="$CONTAINERS" \
  -v services="$SERVICES" \
  -v log_tz="$LOG_TZ" \
  -v warn_pause_ms="$WARN_PAUSE_MS" \
  -v err_pause_ms="$ERR_PAUSE_MS" \
  -v level_filter="$LEVEL_FILTER" \
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
function parse_gc_details(body_str, dec_lvl, dec_tags, up_str,   msg, p_ms, pm, is_stw, lvl, tag_prefix) { \
  msg = body_str; \
  gsub(/^[[:space:]]+|[[:space:]]+$/, "", msg); \
  p_ms = ""; \
  if (match(msg, /([0-9]+(\.[0-9]+)?)ms([[:space:]]|$)/, pm)) { \
    p_ms = pm[1] + 0; \
  } else if (match(msg, /,[[:space:]]*([0-9]+(\.[0-9]+)?)[[:space:]]*secs\]/, pm)) { \
    p_ms = (pm[1] + 0) * 1000; \
  } else if (match(msg, /real=([0-9]+(\.[0-9]+)?)[[:space:]]*secs/, pm)) { \
    p_ms = (pm[1] + 0) * 1000; \
  } \
  is_stw = (msg ~ /(Pause|\[GC|\[Full GC|Safepoint|Allocation Stall)/ && !(msg ~ /Concurrent (Mark|Rebuild|Clear|Reset|Create|Preclean|Sweep|Undo)/)); \
  if (dec_lvl ~ /^(ERROR|FATAL|SEVERE)$/ || msg ~ /(Pause Full|\[Full GC|OutOfMemoryError|Out of Memory|To-space exhausted|Evacuation Failure|Promotion failed|concurrent mode failure|Allocation Stall|GC overhead limit exceeded)/ || (p_ms != "" && is_stw && p_ms >= (err_pause_ms + 0))) { \
    lvl = "ERROR"; \
  } else if (dec_lvl == "WARN" || msg ~ /(System\.gc\(\)|Metadata GC Threshold|GCLocker Initiated GC|Humongous|To-space overflow|Allocation Failure)/ || (p_ms != "" && is_stw && p_ms >= (warn_pause_ms + 0))) { \
    lvl = "WARN"; \
  } else if (dec_lvl != "") { \
    lvl = dec_lvl; \
  } else { \
    lvl = "INFO"; \
  } \
  tag_prefix = (dec_tags != "" && dec_tags != "gc") ? "[" dec_tags "] " : ""; \
  cur_lvl   = lvl; \
  cur_pause = (p_ms != "" && is_stw) ? sprintf("%.1f", p_ms) : ""; \
  cur_body  = tag_prefix msg; \
} \
function flush_rec(   key, cs_key, p_val) { \
  if (cur_ts == "") return; \
  if (cur_ts >= n_start && cur_ts <= n_end) { \
    if (level_filter == "" || cur_lvl ~ "^(" level_filter ")$") { \
      if (keyword_filter == "" || cur_raw ~ keyword_filter || cur_lvl ~ keyword_filter) { \
        key = sprintf("%s\034%09d", cur_ts, ++seq); \
        records[key] = cur_cnt "\034" cur_svc "\034" cur_tsk "\034" cur_lvl "\034" cur_pause "\034" cur_ts "\034" cur_body; \
        total_cnt++; \
        cs_key = sprintf("%-5s / %-9s", cur_cnt, cur_svc); \
        cs_total[cs_key]++; \
        if (cur_lvl ~ /^(ERROR|FATAL|SEVERE)$/) { err_cnt++; cs_err[cs_key]++; } \
        else if (cur_lvl ~ /^(WARN|WARNING)$/) { warn_cnt++; cs_warn[cs_key]++; } \
        else { info_cnt++; cs_info[cs_key]++; } \
        if (cur_pause != "") { \
          p_val = cur_pause + 0; \
          if (p_val > max_pause) max_pause = p_val; \
          if (p_val > cs_max_pause[cs_key]) cs_max_pause[cs_key] = p_val; \
        } \
      } \
    } \
  } \
  cur_ts = ""; \
} \
BEGIN { \
  n_start  = norm_ts(start_jst, 0); \
  n_end    = norm_ts(end_jst, 1); \
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
  if (match(fpath, /\/([^\/]+)\/logs\/([^\/]+)\/mid\/([^\/]+)\/(gc\.log(\.([0-9]+)(\.current)?)?)$/, m)) { \
    cnt = m[1]; svc = m[2]; tsk = m[3]; fname = m[4]; rnum = m[6]; \
    if (!(ino in seen_ino) || (seen_tsk[ino] == "current" && tsk != "current")) { \
      seen_ino[ino]   = fpath; \
      seen_mtime[ino] = mtime; \
      seen_cnt[ino]   = cnt; \
      seen_svc[ino]   = svc; \
      seen_tsk[ino]   = tsk; \
      seen_rnum[ino]  = rnum; \
    } \
  } \
} \
END { \
  PROCINFO["sorted_in"] = "@ind_str_asc"; \
  for (ino in seen_ino) { \
    r_ord = (seen_rnum[ino] == "") ? 999999 : (seen_rnum[ino] + 0); \
    scan_key = sprintf("%012d\034%06d\034%s", seen_mtime[ino] + 0, r_ord, ino); \
    scan_order[scan_key] = ino; \
  } \
  for (sk in scan_order) { \
    ino   = scan_order[sk]; \
    fpath = seen_ino[ino]; \
    mtime = seen_mtime[ino]; \
    cnt   = seen_cnt[ino]; \
    svc   = seen_svc[ino]; \
    tsk   = (seen_tsk[ino] == "current") ? "current" : substr(seen_tsk[ino], 1, 8); \
    fdate = strftime("%Y-%m-%d", mtime); \
    files_scanned++; \
    cur_ts = ""; cur_cnt = cnt; cur_svc = svc; cur_tsk = tsk; \
    while ((getline line < fpath) > 0) { \
      sub(/\r$/, "", line); \
      if (substr(line, 1, 1) == "[") { \
        rest = line; \
        ts_found = ""; d_s = ""; t_s = ""; ms_s = ""; tz_s = ""; \
        up_s = ""; dec_lvl = ""; dec_tags = ""; \
        while (match(rest, /^\[([^\]]+)\][[:space:]]*/, dm)) { \
          rlen = RLENGTH; \
          tok  = dm[1]; \
          gsub(/^[[:space:]]+|[[:space:]]+$/, "", tok); \
          if (ts_found == "" && match(tok, /^([0-9]{4}[-\/][0-9]{2}[-\/][0-9]{2})[T ]([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?(Z|[+-][0-9]{2}:?[0-9]{2})?$/, tm)) { \
            d_s = tm[1]; gsub(/\//, "-", d_s); \
            t_s = tm[2]; ms_s = tm[4]; tz_s = tm[5]; \
            ts_found = "1"; \
          } else if (match(tok, /^([0-9]+([.,][0-9]+)?)(s|ms)$/, um)) { \
            up_s = tok; \
          } else if (tolower(tok) ~ /^(error|warning|warn|info|debug|trace)$/) { \
            dec_lvl = toupper(tok); \
            if (dec_lvl == "WARNING") dec_lvl = "WARN"; \
          } else if (tok ~ /^[a-z0-9_,[:space:]-]+$/) { \
            if (tok ~ /[a-z]/) { \
              dec_tags = tok; \
              gsub(/[[:space:]]+/, "", dec_tags); \
            } \
          } else { \
            break; \
          } \
          rest = substr(rest, rlen + 1); \
        } \
        if (ts_found == "1") { \
          flush_rec(); \
          cur_ts  = to_jst(d_s, t_s, ms_s, tz_s); \
          cur_raw = line; \
          parse_gc_details(rest, dec_lvl, dec_tags, up_s); \
          continue; \
        } \
      } \
      if (match(line, /^([0-9]{4}[-\/][0-9]{2}[-\/][0-9]{2})[T ]([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?(Z|[+-][0-9]{2}:?[0-9]{2})?:?[[:space:]]+(([0-9]+[.,][0-9]+):[[:space:]]+)?(.*)$/, lm)) { \
        flush_rec(); \
        d_s = lm[1]; gsub(/\//, "-", d_s); \
        cur_ts  = to_jst(d_s, lm[2], lm[4], lm[5]); \
        cur_raw = line; \
        parse_gc_details(lm[8], "", "", lm[7]); \
      } else if (match(line, /^([0-9]{2}:[0-9]{2}:[0-9]{2})([.,]([0-9]{1,6}))?:?[[:space:]]+(.*)$/, lm)) { \
        flush_rec(); \
        cur_ts  = to_jst(fdate, lm[1], lm[3], ""); \
        cur_raw = line; \
        parse_gc_details(lm[4], "", "", ""); \
      } else if (cur_ts != "") { \
        cur_body = cur_body "\036" line; \
        cur_raw  = cur_raw "\n" line; \
      } \
    } \
    flush_rec(); \
    close(fpath); \
  } \
  sep = "========================================================================================================================"; \
  div = "------------------------+-------+-----------+----------+---------+-----------+------------------------------------------"; \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
  printf " %sJBoss EAP GC Log Viewer (JST Sorted)%s | Range: %s%s%s to %s%s%s (JST)\n", C_BOLD, C_RST, C_BCYN, n_start, C_RST, C_BCYN, n_end, C_RST; \
  printf " Target: Container=[%s]  ECS-Service=[%s]  ScannedFiles=%d\n", containers, services, files_scanned+0; \
  printf " Thresholds: %sWARN >= %dms%s | %sERROR >= %dms or Full GC / Evacuation Failure / OOM%s\n", C_WRN_FG, warn_pause_ms+0, C_RST, C_ERR_FG, err_pause_ms+0, C_RST; \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
  printf "%s%-23s | %-5s | %-9s | %-8s | %-7s | %-9s | %s%s\n", C_BOLD, "JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Pause", "GC Log Message", C_RST; \
  printf "%s%s%s\n", C_DIM, div, C_RST; \
  if (total_cnt + 0 == 0) { \
    printf " %s[!] No matching GC log entries found in the specified JST time range.%s\n", C_WRN_FG, C_RST; \
  } \
  for (k in records) { \
    split(records[k], r, "\034"); \
    cnt = r[1]; svc = r[2]; tsk = r[3]; lvl = r[4]; pause = r[5]; ts = r[6]; body = r[7]; \
    n_lines = split(body, blines, "\036"); \
    msg1 = blines[1]; \
    p_str = (pause != "") ? sprintf("%7sms", pause) : "      -  "; \
    if (lvl ~ /^(ERROR|FATAL|SEVERE)$/) { \
      b_lvl = sprintf("%s %-5s %s", C_ERR_BG, substr(lvl, 1, 5), C_RST); \
      b_pau = (pause != "") ? sprintf("%s%9s%s", C_ERR_BG, p_str, C_RST) : sprintf("%s%9s%s", C_ERR_FG, p_str, C_RST); \
      printf "%s%-23s%s | %s%-5s%s | %s%-9s%s | %s%-8s%s | %s | %s | %s%s%s\n", C_ERR_FG, ts, C_RST, C_ERR_FG, cnt, C_RST, C_ERR_FG, svc, C_RST, C_ERR_FG, tsk, C_RST, b_lvl, b_pau, C_ERR_FG, msg1, C_RST; \
      for (i = 2; i <= n_lines; i++) { \
        st_color = (blines[i] ~ /(OutOfMemory|Exception|Error|exhausted|Failure)/) ? C_ERR_FG : C_ERR_ST; \
        printf "%s%-23s | %-5s | %-9s | %-8s |   ->    |           |%s %s%s%s\n", C_DIM, "", cnt, svc, "", C_RST, st_color, blines[i], C_RST; \
      } \
    } else if (lvl ~ /^(WARN|WARNING)$/) { \
      b_lvl = sprintf("%s %-5s %s", C_WRN_BG, substr(lvl, 1, 5), C_RST); \
      b_pau = (pause != "") ? sprintf("%s%9s%s", C_WRN_BG, p_str, C_RST) : sprintf("%s%9s%s", C_WRN_FG, p_str, C_RST); \
      printf "%s%-23s%s | %s%-5s%s | %s%-9s%s | %s%-8s%s | %s | %s | %s%s%s\n", C_WRN_FG, ts, C_RST, C_WRN_FG, cnt, C_RST, C_WRN_FG, svc, C_RST, C_WRN_FG, tsk, C_RST, b_lvl, b_pau, C_WRN_FG, msg1, C_RST; \
      for (i = 2; i <= n_lines; i++) { \
        printf "%s%-23s | %-5s | %-9s | %-8s |   ->    |           |%s %s%s%s\n", C_DIM, "", cnt, svc, "", C_RST, C_WRN_ST, blines[i], C_RST; \
      } \
    } else { \
      c_cnt = (cnt == "front") ? C_BCYN : C_BMAG; \
      c_svc = (svc == "intra-web") ? C_BGRN : (svc == "intra-api") ? C_BBLU : (svc == "inter-api") ? C_BCYN : C_BMAG; \
      c_lvl = (lvl ~ /^(DEBUG|TRACE)$/) ? C_DIM : C_INF_FG; \
      c_msg = (lvl ~ /^(DEBUG|TRACE)$/) ? C_DIM : C_RST; \
      c_pau = (pause != "") ? C_BGRN : C_DIM; \
      b_lvl = sprintf("%s %-5s %s", c_lvl, substr(lvl, 1, 5), C_RST); \
      b_pau = sprintf("%s%9s%s", c_pau, p_str, C_RST); \
      printf "%s%-23s%s | %s%-5s%s | %s%-9s%s | %s%-8s%s | %s | %s | %s%s%s\n", C_CYN, ts, C_RST, c_cnt, cnt, C_RST, c_svc, svc, C_RST, C_DIM, tsk, C_RST, b_lvl, b_pau, c_msg, msg1, C_RST; \
      for (i = 2; i <= n_lines; i++) { \
        printf "%s%-23s | %-5s | %-9s | %-8s |   ->    |           | %s%s\n", C_DIM, "", cnt, svc, "", blines[i], C_RST; \
      } \
    } \
  } \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
  printf " %sSummary:%s Total=%s%d%s  [ %sERROR=%d%s  %sWARN=%d%s  %sINFO/OTHER=%d%s ]  MaxPause=%s%.1fms%s\n", C_BOLD, C_RST, C_BOLD, total_cnt+0, C_RST, C_ERR_FG, err_cnt+0, C_RST, C_WRN_FG, warn_cnt+0, C_RST, C_BGRN, info_cnt+0, C_RST, (max_pause >= err_pause_ms ? C_ERR_FG : (max_pause >= warn_pause_ms ? C_WRN_FG : C_BGRN)), max_pause+0, C_RST; \
  if (total_cnt + 0 > 0) { \
    printf " %sBreakdown by Container / ECS Service:%s\n", C_BOLD, C_RST; \
    for (cs in cs_total) { \
      cp = cs_max_pause[cs] + 0; \
      cp_col = (cp >= err_pause_ms) ? C_ERR_FG : (cp >= warn_pause_ms) ? C_WRN_FG : C_BGRN; \
      printf "   - %-17s : Total=%-4d ( %sERROR=%d%s, %sWARN=%d%s, %sINFO=%d%s, MaxPause=%s%.1fms%s )\n", cs, cs_total[cs], C_ERR_FG, cs_err[cs]+0, C_RST, C_WRN_FG, cs_warn[cs]+0, C_RST, C_BGRN, cs_info[cs]+0, C_RST, cp_col, cp, C_RST; \
    } \
  } \
  printf "%s%s%s\n", C_BCYN, sep, C_RST; \
}'
