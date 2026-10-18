#Requires -Version 5.1
<#
.SYNOPSIS
    Windows 11 JBoss EAP TeraTerm Log (server.log / access_log.log / gc.log) to Excel (.xlsx) Report Converter
    Font: Meiryo UI / Color Theme: Monotone Base

.DESCRIPTION
    Reads TeraTerm log files containing the execution results of:
      - jboss_log_search.sh        (server.log)
      - jboss_access_log_search.sh (access_log.log)
      - jboss_gc_log_search.sh     (gc.log)
    Sets each output column item as the Excel table header and generates a clean,
    line-by-line readable Excel (.xlsx) report styled with Meiryo UI font and a monotone color scheme.
    Runs natively on Windows 11 (PowerShell 5.1 / PowerShell 7+) using standard .NET libraries only.

.PARAMETER InputPath
    One or more TeraTerm log files or directories containing log files.

.PARAMETER OutputPath
    Destination Excel (.xlsx) file path. Defaults to .\jboss_log_report_<YYYYMMDD_HHMMSS>.xlsx.

.EXAMPLE
    .\jboss_log_to_excel.ps1 -InputPath .\teraterm.log -OutputPath .\jboss_log_report.xlsx
    .\jboss_log_to_excel.ps1 .\server_tt.log, .\access_tt.log, .\gc_tt.log
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromPipeline = $true)]
    [Alias("i", "LogPath", "Path")]
    [string[]]$InputPath = @(),

    [Parameter(Position = 1)]
    [Alias("o", "OutFile")]
    [string]$OutputPath = "",

    [Parameter()]
    [Alias("h")]
    [switch]$Help,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$RemainingArgs = @()
)

begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = "Stop"

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function ConvertFrom-UnicodeEscapes {
        param([string]$Text)
        return [System.Text.RegularExpressions.Regex]::Unescape($Text)
    }

    function Show-Usage {
        Write-Host "Usage:"
        Write-Host "  .\jboss_log_to_excel.ps1 [-InputPath] <logfile_or_dir...> [[-OutputPath] <path.xlsx>]"
        Write-Host ""
        Write-Host "Options:"
        Write-Host "  -InputPath, -i <path[]>    Input TeraTerm log file(s) or directory"
        Write-Host "  -OutputPath, -o <xlsx>     Output Excel (.xlsx) file path"
        Write-Host "  -Help, -h, --help          Show this help message"
    }

    $CollectedInputs = New-Object System.Collections.Generic.List[string]
}

process {
    if ($null -ne $InputPath) {
        foreach ($p in $InputPath) {
            if (-not [string]::IsNullOrWhiteSpace($p)) {
                $CollectedInputs.Add($p)
            }
        }
    }
}

end {
    if ($Help.IsPresent -or ($InputPath -contains "--help") -or ($RemainingArgs -contains "--help")) {
        Show-Usage
        exit 0
    }

    # Ensure command-line InputPath is collected even when stdin is closed under powershell.exe -File
    if ($CollectedInputs.Count -eq 0 -and $null -ne $InputPath) {
        foreach ($p in $InputPath) {
            if (-not [string]::IsNullOrWhiteSpace($p)) {
                $CollectedInputs.Add($p)
            }
        }
    }

    # Handle any extra positional arguments or -o / --output flags passed via powershell.exe -File
    if ($null -ne $RemainingArgs -and $RemainingArgs.Count -gt 0) {
        for ($idx = 0; $idx -lt $RemainingArgs.Count; $idx++) {
            $argVal = $RemainingArgs[$idx]
            if ($argVal -in @("-o", "--output", "-OutputPath", "-OutFile")) {
                if (($idx + 1) -lt $RemainingArgs.Count) {
                    $OutputPath = $RemainingArgs[$idx + 1]
                    $idx++
                }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($argVal)) {
                $CollectedInputs.Add($argVal)
            }
        }
    }

    # If OutputPath was bound positionally to a second .log/.txt input file, treat it as an input file
    if (-not [string]::IsNullOrWhiteSpace($OutputPath) -and ($OutputPath -match '\.(log|txt)$') -and (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
        $CollectedInputs.Add($OutputPath)
        $OutputPath = ""
    }
    # Fallback to environment variables if no arguments passed
    if ($CollectedInputs.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($env:INPUT_LOGS)) {
        foreach ($part in ($env:INPUT_LOGS -split '[;:]')) {
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $CollectedInputs.Add($part.Trim())
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath) -and -not [string]::IsNullOrWhiteSpace($env:OUTPUT_EXCEL)) {
        $OutputPath = $env:OUTPUT_EXCEL
    }

    # Auto-discover .log / .txt files in current directory if none specified
    if ($CollectedInputs.Count -eq 0) {
        $autoFiles = Get-ChildItem -Path (Get-Location).Path -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @(".log", ".txt") } |
            Sort-Object Name
        foreach ($af in $autoFiles) {
            $CollectedInputs.Add($af.FullName)
        }
    }

    if ($CollectedInputs.Count -eq 0) {
        $msgNoInput = ConvertFrom-UnicodeEscapes "[ERROR] \u5165\u529b\u3068\u306a\u308bTeraTerm\u30ed\u30b0\u30d5\u30a1\u30a4\u30eb\u304c\u6307\u5b9a\u3055\u308c\u3066\u3044\u307e\u305b\u3093\u3002-InputPath \u306b\u30ed\u30b0\u30d5\u30a1\u30a4\u30eb\u3092\u6307\u5b9a\u3057\u3066\u304f\u3060\u3055\u3044\u3002"
        Write-Error $msgNoInput
        exit 1
    }

    $ResolvedFiles = New-Object System.Collections.Generic.List[string]
    foreach ($rawItem in $CollectedInputs) {
        if (Test-Path -LiteralPath $rawItem -PathType Container) {
            $dirFiles = Get-ChildItem -LiteralPath $rawItem -File |
                Where-Object { $_.Extension -in @(".log", ".txt") } |
                Sort-Object Name
            foreach ($df in $dirFiles) {
                $ResolvedFiles.Add($df.FullName)
            }
        }
        elseif (Test-Path -LiteralPath $rawItem -PathType Leaf) {
            $ResolvedFiles.Add((Resolve-Path -LiteralPath $rawItem).Path)
        }
        elseif (Test-Path -Path $rawItem) {
            foreach ($rp in (Resolve-Path -Path $rawItem)) {
                if (Test-Path -LiteralPath $rp.Path -PathType Leaf) {
                    $ResolvedFiles.Add($rp.Path)
                }
            }
        }
        else {
            Write-Warning "Path not found: $rawItem"
        }
    }

    if ($ResolvedFiles.Count -eq 0) {
        Write-Error "No valid input log files found."
        exit 1
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $tsNow = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath "jboss_log_report_${tsNow}.xlsx"
    }
    elseif (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
        $OutputPath = [System.IO.Path]::GetFullPath((Join-Path -Path (Get-Location).Path -ChildPath $OutputPath))
    }

    # ==========================================================================
    # 1. Constants & Regex Patterns
    # ==========================================================================
    $LogTypesOrder = @("server.log", "access_log.log", "gc.log")

    $DefaultHeaders = @{
        "server.log"     = @("JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Log Message")
        "access_log.log" = @("JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Sts", "Client IP", "Request / Access Log Details")
        "gc.log"         = @("JST Timestamp", "Cntnr", "ECS Svc", "Task/Dir", "Level", "Pause", "GC Log Message")
    }

    $ReportTitles = @{
        "server.log"     = "JBoss EAP Server Log Report (server.log)"
        "access_log.log" = "JBoss EAP Access Log Report (access_log.log)"
        "gc.log"         = "JBoss EAP GC Log Report (gc.log)"
    }

    $MinColWidths = @{
        "JST Timestamp"                = 25
        "Cntnr"                        = 10
        "ECS Svc"                      = 13
        "Task/Dir"                     = 12
        "Level"                        = 10
        "Sts"                          = 8
        "Client IP"                    = 17
        "Pause"                        = 13
        "Log Message"                  = 85
        "Request / Access Log Details" = 85
        "GC Log Message"               = 85
    }

    $MaxColWidths = @{
        "JST Timestamp"                = 30
        "Cntnr"                        = 18
        "ECS Svc"                      = 22
        "Task/Dir"                     = 22
        "Level"                        = 14
        "Sts"                          = 12
        "Client IP"                    = 28
        "Pause"                        = 18
        "Log Message"                  = 220
        "Request / Access Log Details" = 220
        "GC Log Message"               = 220
    }

    $ReAnsiOsc    = [regex]::new("\x1b\][^\x07\x1b]*(?:\x07|\x1b\\)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReAnsiCsi    = [regex]::new("\x1b\[[0-9;?]*[ -/]*[@-~]", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReAnsiEsc    = [regex]::new("\x1b[@-Z\\-_]|\x1b", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReCtrlChars  = [regex]::new("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReTeraTermTs = [regex]::new("^\[(?:\d{4}[-/]\d{2}[-/]\d{2}|[A-Z][a-z]{2}\s+[A-Z][a-z]{2}\s+\d{1,2})\s+\d{2}:\d{2}:\d{2}(?:[.,]\d{1,6})?(?:\s+\d{4})?\]\s+(?!\[(?:info|warn|warning|error|debug|trace|gc)\b)", [System.Text.RegularExpressions.RegexOptions]::Compiled)

    $ReBannerTitle      = [regex]::new("^\s*JBoss EAP (Log|Access Log|GC Log) Viewer \(JST Sorted\)\s*\|\s*Range:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\d{3})\s+to\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\d{3})\s*\(JST\)\s*$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReBannerTarget     = [regex]::new("^\s*Target:\s*Container=\[(.*?)\]\s+ECS-Service=\[(.*?)\]\s+ScannedFiles=(\d+)\s*$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReBannerThresholds = [regex]::new("^\s*Thresholds:\s*(WARN\s*>=\s*\d+ms\s*\|\s*ERROR\s*>=\s*\d+ms.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReTableHeader      = [regex]::new("^\s*(JST Timestamp\s*\|\s*Cntnr\s*\|\s*ECS Svc\s*\|\s*Task/Dir\s*\|\s*Level\s*\|.+)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReTableDivider     = [regex]::new("^\s*-{15,}\+-{5,}\+-{5,}\+", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReSepLine          = [regex]::new("^\s*={40,}\s*$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReNoMatch          = [regex]::new("^\s*\[!\]\s*No matching .*entries found in the specified JST time range\.", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReSummaryLine      = [regex]::new("^\s*Summary:\s*Total=(\d+)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReBreakdownItem    = [regex]::new("^\s*-\s*([A-Za-z0-9_.-]+)\s*/\s*([A-Za-z0-9_.-]+)\s*:\s*Total=(\d+)", [System.Text.RegularExpressions.RegexOptions]::Compiled)

    $ReMainRow          = [regex]::new("^\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2},\d{3})\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|\s*([A-Z0-9_-]+)\s*\|(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReContRow          = [regex]::new("^(\s*)\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]*?)\s*\|\s*(->)\s*\|(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)

    $ReAccessRest       = [regex]::new("^\s*([0-9]{3}|---)\s*\|\s*([^|]*?)\s*\|(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReAccessContRest   = [regex]::new("^\s*([^|]*?)\s*\|\s*([^|]*?)\s*\|(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReGcRest           = [regex]::new("^\s*([+-]?[0-9]+(?:\.[0-9]+)?ms|-)\s*\|(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReGcContRest       = [regex]::new("^\s*([^|]*?)\s*\|(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $RePauseMs          = [regex]::new("^([+-]?[0-9]+(?:\.[0-9]+)?)ms$", [System.Text.RegularExpressions.RegexOptions]::Compiled)

    # ==========================================================================
    # 2. Helper Functions
    # ==========================================================================
    function Read-LinesAutoEncoding {
        param([string]$FilePath)
        [byte[]]$bytes = [System.IO.File]::ReadAllBytes($FilePath)
        if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
            $text = [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
            return $text -split "\r?\n|\r"
        }
        try {
            $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
            $text = $utf8Strict.GetString($bytes)
            return $text -split "\r?\n|\r"
        }
        catch {
            try {
                $cp932 = [System.Text.Encoding]::GetEncoding(932)
                $text = $cp932.GetString($bytes)
                return $text -split "\r?\n|\r"
            }
            catch {
                $text = [System.Text.Encoding]::UTF8.GetString($bytes)
                return $text -split "\r?\n|\r"
            }
        }
    }

    function Clean-LogLine {
        param([string]$RawLine)
        if ([string]::IsNullOrEmpty($RawLine)) { return "" }
        $s = $RawLine.TrimEnd([char]"`r", [char]"`n")
        while ($s.IndexOf([char]8) -ge 0) {
            $next = [regex]::Replace($s, "[^\x08]\x08", "")
            if ($next -eq $s) {
                $s = $s.Replace([string][char]8, "")
                break
            }
            $s = $next
        }
        $s = $ReAnsiOsc.Replace($s, "")
        $s = $ReAnsiCsi.Replace($s, "")
        $s = $ReAnsiEsc.Replace($s, "")
        $s = $ReCtrlChars.Replace($s, "")
        $s = $ReTeraTermTs.Replace($s, "")
        return $s
    }

    function Remove-SingleLeadingSpace {
        param([string]$Text)
        if ($null -eq $Text) { return "" }
        if ($Text.StartsWith(" ")) { return $Text.Substring(1) }
        return $Text
    }

    function Get-InferredLogType {
        param([string]$Rest, [bool]$IsCont)
        if (-not $IsCont) {
            if ($ReAccessRest.IsMatch($Rest)) { return "access_log.log" }
            if ($ReGcRest.IsMatch($Rest)) { return "gc.log" }
            return "server.log"
        }
        else {
            $mAcc = $ReAccessContRest.Match($Rest)
            if ($mAcc.Success -and $mAcc.Groups[1].Value -eq "" -and $mAcc.Groups[2].Value -eq "") {
                return "access_log.log"
            }
            $mGc = $ReGcContRest.Match($Rest)
            if ($mGc.Success -and $mGc.Groups[1].Value -eq "") {
                return "gc.log"
            }
            return "server.log"
        }
    }

    function Split-RowColumns {
        param(
            [string]$LogType,
            [string]$Ts,
            [string]$Cnt,
            [string]$Svc,
            [string]$Tsk,
            [string]$Lvl,
            [string]$Rest,
            [bool]$IsCont
        )
        if ($LogType -eq "access_log.log") {
            $m = if ($IsCont) { $ReAccessContRest.Match($Rest) } else { $ReAccessRest.Match($Rest) }
            if (-not $m.Success) { $m = $ReAccessContRest.Match($Rest) }
            if ($m.Success) {
                $sts = $m.Groups[1].Value.Trim()
                $ip  = $m.Groups[2].Value.Trim()
                $msg = Remove-SingleLeadingSpace $m.Groups[3].Value
                return ,([string[]]@($Ts, $Cnt, $Svc, $Tsk, $Lvl, $sts, $ip, $msg))
            }
            return ,([string[]]@($Ts, $Cnt, $Svc, $Tsk, $Lvl, "", "", (Remove-SingleLeadingSpace $Rest)))
        }
        elseif ($LogType -eq "gc.log") {
            $m = if ($IsCont) { $ReGcContRest.Match($Rest) } else { $ReGcRest.Match($Rest) }
            if (-not $m.Success) { $m = $ReGcContRest.Match($Rest) }
            if ($m.Success) {
                $pause = $m.Groups[1].Value.Trim()
                $msg   = Remove-SingleLeadingSpace $m.Groups[2].Value
                return ,([string[]]@($Ts, $Cnt, $Svc, $Tsk, $Lvl, $pause, $msg))
            }
            return ,([string[]]@($Ts, $Cnt, $Svc, $Tsk, $Lvl, "", (Remove-SingleLeadingSpace $Rest)))
        }
        else {
            $msg = Remove-SingleLeadingSpace $Rest
            return ,([string[]]@($Ts, $Cnt, $Svc, $Tsk, $Lvl, $msg))
        }
    }

    function Get-DisplayWidth {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return 0 }
        $w = 0
        foreach ($ch in $Text.ToCharArray()) {
            $code = [int]$ch
            if ($code -gt 0x7F) { $w += 2 } else { $w += 1 }
        }
        return $w
    }

    function Get-ColumnLetter {
        param([int]$ColIndex1Based)
        $res = ""
        $n = $ColIndex1Based
        while ($n -gt 0) {
            $rem = ($n - 1) % 26
            $res = [char](65 + $rem) + $res
            $n = [Math]::Floor(($n - 1) / 26)
        }
        return $res
    }

    function Escape-XmlText {
        param([string]$Text)
        if ($null -eq $Text) { return "" }
        $cleaned = $ReCtrlChars.Replace($Text, "")
        if ($cleaned.Length -gt 32760) {
            $cleaned = $cleaned.Substring(0, 32757) + "..."
        }
        return [System.Security.SecurityElement]::Escape($cleaned)
    }

    # ==========================================================================
    # 3. Parse Input TeraTerm Log Files
    # ==========================================================================
    $Buckets = @{}
    foreach ($lt in $LogTypesOrder) {
        $Buckets[$lt] = [PSCustomObject]@{
            LogType      = $lt
            Headers      = [string[]]$DefaultHeaders[$lt].Clone()
            Ranges       = New-Object System.Collections.Generic.List[string]
            Containers   = New-Object System.Collections.Generic.List[string]
            Services     = New-Object System.Collections.Generic.List[string]
            ScannedFiles = 0
            Thresholds   = ""
            Rows         = New-Object System.Collections.Generic.List[object]
            SeenSection  = $false
            EntryCount   = 0
            ErrCount     = 0
            WarnCount    = 0
            InfoCount    = 0
            MaxPauseMs   = 0.0
            CsStats      = @{}
        }
    }

    function Add-BucketRow {
        param($Bucket, [string[]]$Cols, [bool]$IsCont)
        $Bucket.SeenSection = $true
        if (-not $IsCont) {
            $Bucket.EntryCount++
            $lvl = $Cols[4].Trim().ToUpperInvariant()
            if ($lvl -in @("ERROR", "FATAL", "SEVERE")) {
                $sev = "ERROR"
                $Bucket.ErrCount++
            }
            elseif ($lvl -in @("WARN", "WARNING")) {
                $sev = "WARN"
                $Bucket.WarnCount++
            }
            elseif ($lvl -in @("DEBUG", "TRACE")) {
                $sev = "DEBUG"
                $Bucket.InfoCount++
            }
            else {
                $sev = "INFO"
                $Bucket.InfoCount++
            }

            $cnt = if ([string]::IsNullOrWhiteSpace($Cols[1])) { "-" } else { $Cols[1].Trim() }
            $svc = if ([string]::IsNullOrWhiteSpace($Cols[2])) { "-" } else { $Cols[2].Trim() }
            $csKey = "${cnt}|${svc}"
            if (-not $Bucket.CsStats.ContainsKey($csKey)) {
                $Bucket.CsStats[$csKey] = @{
                    Cnt      = $cnt
                    Svc      = $svc
                    Total    = 0
                    Err      = 0
                    Warn     = 0
                    Info     = 0
                    MaxPause = 0.0
                }
            }
            $Bucket.CsStats[$csKey].Total++
            if ($sev -eq "ERROR") { $Bucket.CsStats[$csKey].Err++ }
            elseif ($sev -eq "WARN") { $Bucket.CsStats[$csKey].Warn++ }
            else { $Bucket.CsStats[$csKey].Info++ }

            if ($Bucket.LogType -eq "gc.log" -and $Cols.Count -ge 6) {
                $pStr = $Cols[5].Trim()
                $mP = $RePauseMs.Match($pStr)
                if ($mP.Success) {
                    $pVal = [double]::Parse($mP.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture)
                    if ($pVal -gt $Bucket.MaxPauseMs) { $Bucket.MaxPauseMs = $pVal }
                    if ($pVal -gt $Bucket.CsStats[$csKey].MaxPause) { $Bucket.CsStats[$csKey].MaxPause = $pVal }
                }
            }

            $Bucket.Rows.Add([PSCustomObject]@{
                Cols     = $Cols
                IsCont   = $false
                Severity = $sev
                EntryIdx = $Bucket.EntryCount
            })
        }
        else {
            $parentSev = if ($Bucket.Rows.Count -gt 0) { $Bucket.Rows[$Bucket.Rows.Count - 1].Severity } else { "INFO" }
            $parentIdx = if ($Bucket.Rows.Count -gt 0) { $Bucket.Rows[$Bucket.Rows.Count - 1].EntryIdx } else { 1 }
            $Bucket.Rows.Add([PSCustomObject]@{
                Cols     = $Cols
                IsCont   = $true
                Severity = $parentSev
                EntryIdx = $parentIdx
            })
        }
    }

    $ReRawServer    = [regex]::new("^(\d{4}[-/]\d{2}[-/]\d{2})[T ](\d{2}:\d{2}:\d{2})(?:[,.](\d{1,6}))?(?:Z|[+-]\d{2}:?\d{2})?\s+\[?([A-Z]+)\]?\s+(.*)$", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReRawAccess    = [regex]::new('^(\S+)\s+\S+\s+(\S+)\s+\[(\d{1,2})/([A-Za-z]{3})/(\d{4}):(\d{2}:\d{2}:\d{2})(?:[.,](\d{1,6}))?\s*([+-]\d{4}|Z)?\]\s+("[^"]*")\s+([1-5]\d{2})(?:\s+(.*))?$', [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $ReRawGcUnified = [regex]::new("^\[(\d{4}[-/]\d{2}[-/]\d{2})[T ](\d{2}:\d{2}:\d{2})(?:[.,](\d{1,6}))?(?:Z|[+-]\d{2}:?\d{2})?\](?:\[[^\]]*\])*\[?(info|warning|warn|error|debug|trace)\]?(?:\[([^\]]*)\])?\s*(.*)$", ([System.Text.RegularExpressions.RegexOptions]::Compiled -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase))
    $ReRawGcPause   = [regex]::new("([0-9]+(?:\.[0-9]+)?)ms(?:\s|$)", [System.Text.RegularExpressions.RegexOptions]::Compiled)
    $MonthMap = @{
        "JAN"="01"; "FEB"="02"; "MAR"="03"; "APR"="04"; "MAY"="05"; "JUN"="06";
        "JUL"="07"; "AUG"="08"; "SEP"="09"; "OCT"="10"; "NOV"="11"; "DEC"="12"
    }

    foreach ($fp in $ResolvedFiles) {
        $CurrentType = $null
        $PendingTarget = $null
        $PendingThresholds = $null
        $fileMatchedTableRows = 0
        $lines = Read-LinesAutoEncoding -FilePath $fp

        foreach ($rawLine in $lines) {
            $line = Clean-LogLine -RawLine $rawLine
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            # 1. Banner Title
            $mTitle = $ReBannerTitle.Match($line)
            if ($mTitle.Success) {
                $vKind = $mTitle.Groups[1].Value
                if ($vKind -eq "Log") { $CurrentType = "server.log" }
                elseif ($vKind -eq "Access Log") { $CurrentType = "access_log.log" }
                elseif ($vKind -eq "GC Log") { $CurrentType = "gc.log" }

                $b = $Buckets[$CurrentType]
                $b.SeenSection = $true
                $rStr = "$($mTitle.Groups[2].Value) to $($mTitle.Groups[3].Value)"
                if (-not $b.Ranges.Contains($rStr)) { $b.Ranges.Add($rStr) }
                continue
            }

            # 2. Banner Target
            $mTarget = $ReBannerTarget.Match($line)
            if ($mTarget.Success) {
                $cntT = $mTarget.Groups[1].Value
                $svcT = $mTarget.Groups[2].Value
                $scCnt = [int]$mTarget.Groups[3].Value
                if ($null -ne $CurrentType) {
                    $b = $Buckets[$CurrentType]
                    if ($cntT -and -not $b.Containers.Contains($cntT)) { $b.Containers.Add($cntT) }
                    if ($svcT -and -not $b.Services.Contains($svcT)) { $b.Services.Add($svcT) }
                    $b.ScannedFiles += $scCnt
                }
                else {
                    $PendingTarget = @($cntT, $svcT, $scCnt)
                }
                continue
            }

            # 3. Banner Thresholds
            $mTh = $ReBannerThresholds.Match($line)
            if ($mTh.Success) {
                $thStr = $mTh.Groups[1].Value.Trim()
                if ($CurrentType -eq "gc.log") {
                    $Buckets["gc.log"].Thresholds = $thStr
                }
                else {
                    $PendingThresholds = $thStr
                }
                continue
            }

            # 4. Table Header Row
            $mHdr = $ReTableHeader.Match($line)
            if ($mHdr.Success) {
                if ($line.Contains("printf ") -or $line.Contains("%s") -or $line.Contains("\n")) {
                    continue
                }
                $hdrParts = @($mHdr.Groups[1].Value -split "\|" | ForEach-Object { $_.Trim() })
                if ($hdrParts.Count -eq 8 -or ($hdrParts -contains "Client IP")) {
                    $CurrentType = "access_log.log"
                }
                elseif ($hdrParts.Count -eq 7 -or ($hdrParts -contains "Pause")) {
                    $CurrentType = "gc.log"
                }
                else {
                    $CurrentType = "server.log"
                }

                $b = $Buckets[$CurrentType]
                $b.SeenSection = $true
                if ($hdrParts.Count -eq $DefaultHeaders[$CurrentType].Count) {
                    $b.Headers = [string[]]$hdrParts
                }
                if ($null -ne $PendingTarget) {
                    if ($PendingTarget[0] -and -not $b.Containers.Contains($PendingTarget[0])) { $b.Containers.Add($PendingTarget[0]) }
                    if ($PendingTarget[1] -and -not $b.Services.Contains($PendingTarget[1])) { $b.Services.Add($PendingTarget[1]) }
                    $b.ScannedFiles += [int]$PendingTarget[2]
                    $PendingTarget = $null
                }
                if ($null -ne $PendingThresholds -and $CurrentType -eq "gc.log") {
                    $Buckets["gc.log"].Thresholds = $PendingThresholds
                    $PendingThresholds = $null
                }
                continue
            }

            # 5. Dividers / Separators / Summary Footer
            if ($ReTableDivider.IsMatch($line) -or $ReSepLine.IsMatch($line) -or $ReNoMatch.IsMatch($line)) {
                continue
            }
            if ($ReSummaryLine.IsMatch($line) -or $ReBreakdownItem.IsMatch($line) -or $line.TrimStart().StartsWith("Breakdown by Container")) {
                $CurrentType = $null
                continue
            }

            # 6. Main Data Row
            $mMain = $ReMainRow.Match($line)
            if ($mMain.Success) {
                $ts   = $mMain.Groups[1].Value.Trim()
                $cnt  = $mMain.Groups[2].Value.Trim()
                $svc  = $mMain.Groups[3].Value.Trim()
                $tsk  = $mMain.Groups[4].Value.Trim()
                $lvl  = $mMain.Groups[5].Value.Trim()
                $rest = $mMain.Groups[6].Value

                $lt = if ($null -ne $CurrentType) { $CurrentType } else { Get-InferredLogType -Rest $rest -IsCont $false }
                $CurrentType = $lt
                $cols = Split-RowColumns -LogType $lt -Ts $ts -Cnt $cnt -Svc $svc -Tsk $tsk -Lvl $lvl -Rest $rest -IsCont $false
                Add-BucketRow -Bucket $Buckets[$lt] -Cols $cols -IsCont $false
                $fileMatchedTableRows++
                continue
            }

            # 7. Continuation Data Row ("->")
            $mCont = $ReContRow.Match($line)
            if ($mCont.Success) {
                $cnt  = $mCont.Groups[2].Value.Trim()
                $svc  = $mCont.Groups[3].Value.Trim()
                $tsk  = $mCont.Groups[4].Value.Trim()
                $lvl  = $mCont.Groups[5].Value.Trim()
                $rest = $mCont.Groups[6].Value

                $lt = if ($null -ne $CurrentType) { $CurrentType } else { Get-InferredLogType -Rest $rest -IsCont $true }
                $CurrentType = $lt
                $cols = Split-RowColumns -LogType $lt -Ts "" -Cnt $cnt -Svc $svc -Tsk $tsk -Lvl $lvl -Rest $rest -IsCont $true
                Add-BucketRow -Bucket $Buckets[$lt] -Cols $cols -IsCont $true
                $fileMatchedTableRows++
                continue
            }
        }

        # Fallback: Raw log lines if file had 0 pipe-formatted table rows
        if ($fileMatchedTableRows -eq 0) {
            $lastRawLt = $null
            foreach ($rawLine in $lines) {
                $line = Clean-LogLine -RawLine $rawLine
                if ([string]::IsNullOrWhiteSpace($line)) { continue }

                $mGc = $ReRawGcUnified.Match($line)
                if ($mGc.Success) {
                    $dS = $mGc.Groups[1].Value.Replace("/", "-")
                    $msS = ($mGc.Groups[3].Value + "000").Substring(0, 3)
                    $ts = "$dS $($mGc.Groups[2].Value),$msS"
                    $rawLvl = if ($mGc.Groups[4].Value) { $mGc.Groups[4].Value.ToUpperInvariant() } else { "INFO" }
                    if ($rawLvl -eq "WARNING") { $rawLvl = "WARN" }
                    $msg = $mGc.Groups[6].Value.Trim()
                    $mPMs = $ReRawGcPause.Match($msg)
                    $pVal = if ($mPMs.Success) { [double]::Parse($mPMs.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) } else { -1.0 }
                    $pauseS = if ($pVal -ge 0) { ("{0:F1}ms" -f $pVal) } else { "-" }
                    if ($msg.Contains("Pause Full") -or $msg.Contains("OutOfMemory") -or $pVal -ge 1000.0) { $rawLvl = "ERROR" }
                    elseif ($pVal -ge 200.0) { $rawLvl = "WARN" }
                    Add-BucketRow -Bucket $Buckets["gc.log"] -Cols @($ts, "-", "-", "-", $rawLvl, $pauseS, $msg) -IsCont $false
                    $lastRawLt = "gc.log"
                    continue
                }

                $mAcc = $ReRawAccess.Match($line)
                if ($mAcc.Success) {
                    $ip  = $mAcc.Groups[1].Value
                    $usr = $mAcc.Groups[2].Value
                    $dd  = [int]$mAcc.Groups[3].Value
                    $mon = $mAcc.Groups[4].Value.ToUpperInvariant()
                    $yr  = [int]$mAcc.Groups[5].Value
                    $tS  = $mAcc.Groups[6].Value
                    $msS = ($mAcc.Groups[7].Value + "000").Substring(0, 3)
                    $req = $mAcc.Groups[9].Value
                    $sts = $mAcc.Groups[10].Value
                    $tl  = $mAcc.Groups[11].Value
                    $mm  = if ($MonthMap.ContainsKey($mon)) { $MonthMap[$mon] } else { "01" }
                    $ts  = ("{0:D4}-{1}-{2:D2} {3},{4}" -f $yr, $mm, $dd, $tS, $msS)
                    $lvl = if ($sts.StartsWith("5")) { "ERROR" } elseif ($sts.StartsWith("4")) { "WARN" } else { "INFO" }
                    $uPart = if ($usr -and $usr -ne "-") { "[user:$usr] " } else { "" }
                    $tPart = if ($tl) { " $tl" } else { "" }
                    Add-BucketRow -Bucket $Buckets["access_log.log"] -Cols @($ts, "-", "-", "-", $lvl, $sts, $ip, "${uPart}${req}${tPart}") -IsCont $false
                    $lastRawLt = "access_log.log"
                    continue
                }

                $mSrv = $ReRawServer.Match($line)
                if ($mSrv.Success) {
                    $dS = $mSrv.Groups[1].Value.Replace("/", "-")
                    $msS = ($mSrv.Groups[3].Value + "000").Substring(0, 3)
                    $ts = "$dS $($mSrv.Groups[2].Value),$msS"
                    $lvl = $mSrv.Groups[4].Value.ToUpperInvariant()
                    $msg = $mSrv.Groups[5].Value
                    Add-BucketRow -Bucket $Buckets["server.log"] -Cols @($ts, "-", "-", "-", $lvl, $msg) -IsCont $false
                    $lastRawLt = "server.log"
                    continue
                }

                if ($lastRawLt -eq "server.log" -and $Buckets["server.log"].Rows.Count -gt 0) {
                    Add-BucketRow -Bucket $Buckets["server.log"] -Cols @("", "-", "-", "", "->", $line) -IsCont $true
                }
            }
        }
    }

    # ==========================================================================
    # 4. Shared String Table & Style Definitions (Meiryo UI / Monotone)
    # ==========================================================================
    $SstStrings = New-Object System.Collections.Generic.List[string]
    $SstMap = New-Object 'System.Collections.Generic.Dictionary[string,int]'
    $script:SstTotalRefs = 0

    function Get-SharedStringIndex {
        param([AllowNull()][string]$Text)
        $script:SstTotalRefs++
        $s = if ($null -eq $Text) { "" } else { $Text }
        $s = $ReCtrlChars.Replace($s, "")
        if ($s.Length -gt 32760) {
            $s = $s.Substring(0, 32757) + "..."
        }
        $idx = 0
        if (-not $SstMap.TryGetValue($s, [ref]$idx)) {
            $idx = $SstStrings.Count
            $SstStrings.Add($s)
            $SstMap[$s] = $idx
        }
        return $idx
    }

    function New-StrCell {
        param([string]$Ref, [int]$StyleId, [AllowNull()][string]$Text)
        $idx = Get-SharedStringIndex -Text $Text
        return "<c r=`"$Ref`" s=`"$StyleId`" t=`"s`"><v>$idx</v></c>"
    }

    function New-NumCell {
        param([string]$Ref, [int]$StyleId, $Value)
        return "<c r=`"$Ref`" s=`"$StyleId`"><v>$Value</v></c>"
    }

    # Style Constants
    $STYLE_TITLE_BANNER = 1
    $STYLE_SUBTITLE_BAR = 2
    $STYLE_META_BAR     = 3
    $STYLE_SUMMARY_BAR  = 4
    $STYLE_SECTION_HDR  = 5
    $STYLE_TBL_HEADER_C = 6
    $STYLE_TBL_HEADER_L = 7
    $STYLE_TBL_HEADER_R = 8

    $STYLE_INFO_EVEN_C  = 9
    $STYLE_INFO_EVEN_L  = 10
    $STYLE_INFO_EVEN_R  = 11
    $STYLE_INFO_ODD_C   = 12
    $STYLE_INFO_ODD_L   = 13
    $STYLE_INFO_ODD_R   = 14

    $STYLE_DBG_EVEN_C   = 15
    $STYLE_DBG_EVEN_L   = 16
    $STYLE_DBG_EVEN_R   = 17
    $STYLE_DBG_ODD_C    = 18
    $STYLE_DBG_ODD_L    = 19
    $STYLE_DBG_ODD_R    = 20

    $STYLE_WARN_C       = 21
    $STYLE_WARN_L       = 22
    $STYLE_WARN_R       = 23
    $STYLE_WARN_BADGE_C = 24
    $STYLE_WARN_BADGE_R = 25

    $STYLE_ERR_C        = 26
    $STYLE_ERR_L        = 27
    $STYLE_ERR_R        = 28
    $STYLE_ERR_BADGE_C  = 29
    $STYLE_ERR_BADGE_R  = 30

    $STYLE_CONT_INFO_C  = 31
    $STYLE_CONT_INFO_L  = 32
    $STYLE_CONT_INFO_R  = 33
    $STYLE_CONT_WARN_C  = 34
    $STYLE_CONT_WARN_L  = 35
    $STYLE_CONT_WARN_R  = 36
    $STYLE_CONT_ERR_C   = 37
    $STYLE_CONT_ERR_L   = 38
    $STYLE_CONT_ERR_R   = 39

    $STYLE_NUM_EVEN     = 40
    $STYLE_NUM_ODD      = 41
    $STYLE_NUM_ERR      = 42
    $STYLE_NUM_WARN     = 43

    function Build-StylesXml {
        $fonts = @(
            '<font><sz val="10"/><color rgb="FF1A1A1A"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><b/><sz val="13"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><sz val="9.5"/><color rgb="FFE6E6E6"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><b/><sz val="9.5"/><color rgb="FF1A1A1A"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><sz val="9.5"/><color rgb="FF262626"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><b/><sz val="11"/><color rgb="FF1F1F1F"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><b/><sz val="10"/><color rgb="FF000000"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><b/><sz val="10"/><color rgb="FF1A1A1A"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><sz val="10"/><color rgb="FF595959"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>',
            '<font><sz val="10"/><color rgb="FF262626"/><name val="Meiryo UI"/><family val="2"/><charset val="128"/></font>'
        )

        $fills = @(
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FF1F1F1F"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FF333333"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FF262626"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FF595959"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFE6E6E6"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFECECEC"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFF4F4F4"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid"><fgColor rgb="FFF8F8F8"/><bgColor indexed="64"/></patternFill></fill>'
        )

        $borders = @(
            '<border><left/><right/><top/><bottom/><diagonal/></border>',
            '<border><left style="thin"><color rgb="FF595959"/></left><right style="thin"><color rgb="FF595959"/></right><top style="medium"><color rgb="FF1F1F1F"/></top><bottom style="medium"><color rgb="FF000000"/></bottom><diagonal/></border>',
            '<border><left style="thin"><color rgb="FFD0D0D0"/></left><right style="thin"><color rgb="FFD0D0D0"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>',
            '<border><left style="thin"><color rgb="FFD0D0D0"/></left><right style="thin"><color rgb="FFD0D0D0"/></right><top style="dotted"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>',
            '<border><left style="thin"><color rgb="FF808080"/></left><right style="thin"><color rgb="FF808080"/></right><top style="thin"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>',
            '<border><left/><right/><top/><bottom style="medium"><color rgb="FF262626"/></bottom><diagonal/></border>'
        )

        function New-Xf([int]$nf, [int]$fn, [int]$fl, [int]$bd, [string]$ha, [int]$wrap = 0, [int]$ind = 0) {
            $indAttr = if ($ind -gt 0) { " indent=`"$ind`"" } else { "" }
            return "<xf numFmtId=`"$nf`" fontId=`"$fn`" fillId=`"$fl`" borderId=`"$bd`" xfId=`"0`" applyNumberFormat=`"1`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"$ha`" vertical=`"center`" wrapText=`"$wrap`"$indAttr/></xf>"
        }

        $cellXfs = @(
            (New-Xf 49 0 0 0 "left"),
            (New-Xf 49 1 2 4 "left" 0 1),
            (New-Xf 49 2 3 4 "left" 0 1),
            (New-Xf 49 3 6 4 "left" 0 1),
            (New-Xf 49 4 11 4 "left" 0 1),
            (New-Xf 49 5 0 5 "left"),
            (New-Xf 49 6 4 1 "center"),
            (New-Xf 49 6 4 1 "left"),
            (New-Xf 49 6 4 1 "right"),
            (New-Xf 49 0 10 2 "center"),
            (New-Xf 49 0 10 2 "left"),
            (New-Xf 49 0 10 2 "right"),
            (New-Xf 49 0 11 2 "center"),
            (New-Xf 49 0 11 2 "left"),
            (New-Xf 49 0 11 2 "right"),
            (New-Xf 49 9 10 2 "center"),
            (New-Xf 49 9 10 2 "left"),
            (New-Xf 49 9 10 2 "right"),
            (New-Xf 49 9 11 2 "center"),
            (New-Xf 49 9 11 2 "left"),
            (New-Xf 49 9 11 2 "right"),
            (New-Xf 49 8 8 2 "center"),
            (New-Xf 49 8 8 2 "left"),
            (New-Xf 49 8 8 2 "right"),
            (New-Xf 49 6 5 2 "center"),
            (New-Xf 49 6 5 2 "right"),
            (New-Xf 49 7 6 2 "center"),
            (New-Xf 49 7 6 2 "left"),
            (New-Xf 49 7 6 2 "right"),
            (New-Xf 49 6 2 2 "center"),
            (New-Xf 49 6 2 2 "right"),
            (New-Xf 49 9 11 3 "center"),
            (New-Xf 49 9 11 3 "left"),
            (New-Xf 49 9 11 3 "right"),
            (New-Xf 49 10 9 3 "center"),
            (New-Xf 49 10 9 3 "left"),
            (New-Xf 49 10 9 3 "right"),
            (New-Xf 49 10 7 3 "center"),
            (New-Xf 49 10 7 3 "left"),
            (New-Xf 49 10 7 3 "right"),
            (New-Xf 3 0 10 2 "right"),
            (New-Xf 3 0 11 2 "right"),
            (New-Xf 3 7 6 2 "right"),
            (New-Xf 3 8 8 2 "right")
        )

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
            "<fonts count=`"$($fonts.Count)`">$($fonts -join '')</fonts>" +
            "<fills count=`"$($fills.Count)`">$($fills -join '')</fills>" +
            "<borders count=`"$($borders.Count)`">$($borders -join '')</borders>" +
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>' +
            "<cellXfs count=`"$($cellXfs.Count)`">$($cellXfs -join '')</cellXfs>" +
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
            '<dxfs count="0"/>' +
            '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
            '</styleSheet>'
        )
    }

    function Build-DetailSheetXml {
        param($Bucket, [int]$TabSelected = 0)
        $headers = $Bucket.Headers
        $numCols = $headers.Count
        $lastColLet = Get-ColumnLetter -ColIndex1Based $numCols

        $colWidths = New-Object int[] $numCols
        for ($i = 0; $i -lt $numCols; $i++) {
            $hName = $headers[$i]
            $minW = if ($MinColWidths.ContainsKey($hName)) { $MinColWidths[$hName] } else { 14 }
            $maxW = if ($MaxColWidths.ContainsKey($hName)) { $MaxColWidths[$hName] } else { 120 }
            $w = [Math]::Max($minW, (Get-DisplayWidth -Text $hName) + 5)
            $colWidths[$i] = [Math]::Min($maxW, $w)
        }

        foreach ($rObj in $Bucket.Rows) {
            $cols = $rObj.Cols
            $limit = [Math]::Min($numCols, $cols.Count)
            for ($i = 0; $i -lt $limit; $i++) {
                $hName = $headers[$i]
                $maxW = if ($MaxColWidths.ContainsKey($hName)) { $MaxColWidths[$hName] } else { 120 }
                $cw = (Get-DisplayWidth -Text $cols[$i]) + 3
                if ($cw -gt $colWidths[$i]) {
                    $colWidths[$i] = [Math]::Min($maxW, $cw)
                }
            }
        }

        $rangeStr = if ($Bucket.Ranges.Count -gt 0) { $Bucket.Ranges -join " / " } else { "All Ranges (JST)" }
        $cntStr   = if ($Bucket.Containers.Count -gt 0) { $Bucket.Containers -join ", " } else { "front|back" }
        $svcStr   = if ($Bucket.Services.Count -gt 0) { $Bucket.Services -join ", " } else { "intra-web|intra-api|inter-api|sf-api" }
        $thPart   = if (-not [string]::IsNullOrWhiteSpace($Bucket.Thresholds)) { "   |   Thresholds: $($Bucket.Thresholds)" } else { "" }
        $metaText = "Range (JST): $rangeStr   |   Target: Container=[$cntStr]  ECS-Service=[$svcStr]  ScannedFiles=$($Bucket.ScannedFiles)$thPart"

        $maxPPart = if ($Bucket.LogType -eq "gc.log") { ("   |   MaxPause={0:F1}ms" -f $Bucket.MaxPauseMs) } else { "" }
        $summaryText = "Summary: Total=$($Bucket.EntryCount)  [ ERROR=$($Bucket.ErrCount) / WARN=$($Bucket.WarnCount) / INFO=$($Bucket.InfoCount) ]   (Output Detail Lines=$($Bucket.Rows.Count))$maxPPart"

        $sbRows = New-Object System.Text.StringBuilder
        $merges = New-Object System.Collections.Generic.List[string]

        # Row 1: Title Banner
        [void]$sbRows.Append('<row r="1" ht="30" customHeight="1">')
        [void]$sbRows.Append((New-StrCell -Ref "A1" -StyleId $STYLE_TITLE_BANNER -Text $ReportTitles[$Bucket.LogType]))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)1" -StyleId $STYLE_TITLE_BANNER -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A1:${lastColLet}1")

        # Row 2: Metadata Bar
        [void]$sbRows.Append('<row r="2" ht="20" customHeight="1">')
        [void]$sbRows.Append((New-StrCell -Ref "A2" -StyleId $STYLE_META_BAR -Text $metaText))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)2" -StyleId $STYLE_META_BAR -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A2:${lastColLet}2")

        # Row 3: Summary Bar
        [void]$sbRows.Append('<row r="3" ht="20" customHeight="1">')
        [void]$sbRows.Append((New-StrCell -Ref "A3" -StyleId $STYLE_SUMMARY_BAR -Text $summaryText))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)3" -StyleId $STYLE_SUMMARY_BAR -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A3:${lastColLet}3")

        # Row 4: Spacer
        [void]$sbRows.Append('<row r="4" ht="8" customHeight="1"/>')

        # Row 5: Table Header
        [void]$sbRows.Append('<row r="5" ht="26" customHeight="1">')
        for ($c = 1; $c -le $numCols; $c++) {
            $hName = $headers[$c - 1]
            $hdrStyle = if ($c -eq $numCols) { $STYLE_TBL_HEADER_L } elseif ($hName -eq "Pause") { $STYLE_TBL_HEADER_R } else { $STYLE_TBL_HEADER_C }
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)5" -StyleId $hdrStyle -Text $hName))
        }
        [void]$sbRows.Append('</row>')

        # Rows 6+: Data Rows
        $curRowNum = 6
        if ($Bucket.Rows.Count -eq 0) {
            $noMatchMsg = ConvertFrom-UnicodeEscapes "[!] \u6307\u5b9a\u3055\u308c\u305fJST\u6642\u9593\u7bc4\u56f2\u306b\u8a72\u5f53\u3059\u308b\u30ed\u30b0\u30a8\u30f3\u30c8\u30ea\u306f\u3042\u308a\u307e\u305b\u3093\u3002"
            [void]$sbRows.Append('<row r="6" ht="20" customHeight="1">')
            for ($c = 1; $c -lt $numCols; $c++) {
                [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)6" -StyleId $STYLE_INFO_EVEN_C -Text "-"))
            }
            [void]$sbRows.Append((New-StrCell -Ref "${lastColLet}6" -StyleId $STYLE_INFO_EVEN_L -Text $noMatchMsg))
            [void]$sbRows.Append('</row>')
        }
        else {
            for ($rIdx = 0; $rIdx -lt $Bucket.Rows.Count; $rIdx++) {
                $curRowNum = 6 + $rIdx
                $rObj = $Bucket.Rows[$rIdx]
                $cols = $rObj.Cols
                $isCont = $rObj.IsCont
                $sev = $rObj.Severity
                $evenEntry = (($rObj.EntryIdx % 2) -eq 1)

                if ($isCont) {
                    if ($sev -eq "ERROR") {
                        $stC = $STYLE_CONT_ERR_C; $stL = $STYLE_CONT_ERR_L; $stR = $STYLE_CONT_ERR_R
                    }
                    elseif ($sev -eq "WARN") {
                        $stC = $STYLE_CONT_WARN_C; $stL = $STYLE_CONT_WARN_L; $stR = $STYLE_CONT_WARN_R
                    }
                    else {
                        $stC = $STYLE_CONT_INFO_C; $stL = $STYLE_CONT_INFO_L; $stR = $STYLE_CONT_INFO_R
                    }
                }
                else {
                    if ($sev -eq "ERROR") {
                        $stC = $STYLE_ERR_C; $stL = $STYLE_ERR_L; $stR = $STYLE_ERR_R
                    }
                    elseif ($sev -eq "WARN") {
                        $stC = $STYLE_WARN_C; $stL = $STYLE_WARN_L; $stR = $STYLE_WARN_R
                    }
                    elseif ($sev -eq "DEBUG") {
                        if ($evenEntry) {
                            $stC = $STYLE_DBG_EVEN_C; $stL = $STYLE_DBG_EVEN_L; $stR = $STYLE_DBG_EVEN_R
                        }
                        else {
                            $stC = $STYLE_DBG_ODD_C; $stL = $STYLE_DBG_ODD_L; $stR = $STYLE_DBG_ODD_R
                        }
                    }
                    else {
                        if ($evenEntry) {
                            $stC = $STYLE_INFO_EVEN_C; $stL = $STYLE_INFO_EVEN_L; $stR = $STYLE_INFO_EVEN_R
                        }
                        else {
                            $stC = $STYLE_INFO_ODD_C; $stL = $STYLE_INFO_ODD_L; $stR = $STYLE_INFO_ODD_R
                        }
                    }
                }

                [void]$sbRows.Append("<row r=`"$curRowNum`" ht=`"20`" customHeight=`"1`">")
                for ($c = 1; $c -le $numCols; $c++) {
                    $val = if (($c - 1) -lt $cols.Count) { $cols[$c - 1] } else { "" }
                    $hName = $headers[$c - 1]
                    $cellRef = "$(Get-ColumnLetter $c)$curRowNum"

                    if ($c -eq $numCols) {
                        $cStyle = $stL
                    }
                    elseif ($hName -eq "Level" -and (-not $isCont)) {
                        if ($sev -eq "ERROR") { $cStyle = $STYLE_ERR_BADGE_C }
                        elseif ($sev -eq "WARN") { $cStyle = $STYLE_WARN_BADGE_C }
                        else { $cStyle = $stC }
                    }
                    elseif ($hName -eq "Sts" -and (-not $isCont)) {
                        if ($sev -eq "ERROR" -or $val.StartsWith("5")) { $cStyle = $STYLE_ERR_BADGE_C }
                        elseif ($sev -eq "WARN" -or $val.StartsWith("4")) { $cStyle = $STYLE_WARN_BADGE_C }
                        else { $cStyle = $stC }
                    }
                    elseif ($hName -eq "Pause") {
                        if ((-not $isCont) -and (-not [string]::IsNullOrEmpty($val)) -and ($val -ne "-")) {
                            if ($sev -eq "ERROR") { $cStyle = $STYLE_ERR_BADGE_R }
                            elseif ($sev -eq "WARN") { $cStyle = $STYLE_WARN_BADGE_R }
                            else { $cStyle = $stR }
                        }
                        else {
                            $cStyle = if ($val -eq "-") { $stC } else { $stR }
                        }
                    }
                    else {
                        $cStyle = $stC
                    }

                    [void]$sbRows.Append((New-StrCell -Ref $cellRef -StyleId $cStyle -Text $val))
                }
                [void]$sbRows.Append('</row>')
            }
        }

        $sbCols = New-Object System.Text.StringBuilder
        [void]$sbCols.Append('<cols>')
        for ($c = 1; $c -le $numCols; $c++) {
            [void]$sbCols.Append("<col min=`"$c`" max=`"$c`" width=`"$($colWidths[$c - 1])`" customWidth=`"1`"/>")
        }
        [void]$sbCols.Append('</cols>')

        $mergeXml = "<mergeCells count=`"$($merges.Count)`">" + (($merges | ForEach-Object { "<mergeCell ref=`"$_`"/>" }) -join "") + "</mergeCells>"

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
            "<dimension ref=`"A1:${lastColLet}${curRowNum}`"/>" +
            "<sheetViews><sheetView tabSelected=`"$TabSelected`" showGridLines=`"1`" workbookViewId=`"0`">" +
            '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/>' +
            '<selection pane="bottomLeft" activeCell="A6" sqref="A6"/>' +
            '</sheetView></sheetViews>' +
            '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>' +
            $sbCols.ToString() +
            "<sheetData>$($sbRows.ToString())</sheetData>" +
            "<autoFilter ref=`"A5:${lastColLet}${curRowNum}`"/>" +
            $mergeXml +
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
            '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
            '</worksheet>'
        )
    }

    function Build-SummarySheetXml {
        param($Buckets, [string[]]$SourceFiles, [int]$TabSelected = 1)
        $nowStr = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        $srcNames = ($SourceFiles | ForEach-Object { [System.IO.Path]::GetFileName($_) }) -join ", "
        if ($srcNames.Length -gt 140) { $srcNames = $srcNames.Substring(0, 137) + "..." }

        $sbRows = New-Object System.Text.StringBuilder
        $merges = New-Object System.Collections.Generic.List[string]
        $numCols = 9
        $lastColLet = Get-ColumnLetter -ColIndex1Based $numCols

        $titleStr = ConvertFrom-UnicodeEscapes "JBoss EAP \u30ed\u30b0\u89e3\u6790\u30ec\u30dd\u30fc\u30c8 \u30b5\u30de\u30ea\u30fc (server.log / access_log.log / gc.log)"
        [void]$sbRows.Append('<row r="1" ht="30" customHeight="1">')
        [void]$sbRows.Append((New-StrCell -Ref "A1" -StyleId $STYLE_TITLE_BANNER -Text $titleStr))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)1" -StyleId $STYLE_TITLE_BANNER -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A1:${lastColLet}1")

        $subText = "Generated: $nowStr (JST)   |   Source Log Files: $srcNames"
        [void]$sbRows.Append('<row r="2" ht="20" customHeight="1">')
        [void]$sbRows.Append((New-StrCell -Ref "A2" -StyleId $STYLE_SUBTITLE_BAR -Text $subText))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)2" -StyleId $STYLE_SUBTITLE_BAR -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A2:${lastColLet}2")

        [void]$sbRows.Append('<row r="3" ht="10" customHeight="1"/>')

        $sec1Title = ConvertFrom-UnicodeEscapes "\u25a0 1. \u30ed\u30b0\u7a2e\u5225\u3054\u3068\u96c6\u8a08\u30b5\u30de\u30ea\u30fc (Summary by Log Type)"
        [void]$sbRows.Append('<row r="4" ht="24" customHeight="1">')
        [void]$sbRows.Append((New-StrCell -Ref "A4" -StyleId $STYLE_SECTION_HDR -Text $sec1Title))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)4" -StyleId $STYLE_SECTION_HDR -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A4:${lastColLet}4")

        $sec1Headers = @(
            (ConvertFrom-UnicodeEscapes "\u30ed\u30b0\u7a2e\u5225 (Sheet)"),
            (ConvertFrom-UnicodeEscapes "\u62bd\u51fa\u5bfe\u8c61\u671f\u9593 (JST Range)"),
            (ConvertFrom-UnicodeEscapes "\u5bfe\u8c61\u30b3\u30f3\u30c6\u30ca"),
            (ConvertFrom-UnicodeEscapes "\u5bfe\u8c61ECS\u30b5\u30fc\u30d3\u30b9"),
            (ConvertFrom-UnicodeEscapes "\u7dcf\u30a4\u30d9\u30f3\u30c8\u6570"),
            (ConvertFrom-UnicodeEscapes "\u51fa\u529b\u884c\u6570 (\u542b\u30b9\u30bf\u30c3\u30af\u30c8\u30ec\u30fc\u30b9)"),
            (ConvertFrom-UnicodeEscapes "ERROR \u4ef6\u6570"),
            (ConvertFrom-UnicodeEscapes "WARN \u4ef6\u6570"),
            (ConvertFrom-UnicodeEscapes "INFO/\u4ed6 \u4ef6\u6570 / \u6700\u5927GC\u505c\u6b62")
        )
        [void]$sbRows.Append('<row r="5" ht="25" customHeight="1">')
        for ($c = 1; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)5" -StyleId $STYLE_TBL_HEADER_C -Text $sec1Headers[$c - 1]))
        }
        [void]$sbRows.Append('</row>')

        $curR = 6
        $unitKen = ConvertFrom-UnicodeEscapes "\u4ef6"
        for ($idx = 0; $idx -lt $LogTypesOrder.Count; $idx++) {
            $lt = $LogTypesOrder[$idx]
            $b = $Buckets[$lt]
            $even = (($idx % 2) -eq 0)
            $stC = if ($even) { $STYLE_INFO_EVEN_C } else { $STYLE_INFO_ODD_C }
            $stL = if ($even) { $STYLE_INFO_EVEN_L } else { $STYLE_INFO_ODD_L }
            $stNum = if ($even) { $STYLE_NUM_EVEN } else { $STYLE_NUM_ODD }
            $stErr = if ($b.ErrCount -gt 0) { $STYLE_NUM_ERR } else { $stNum }
            $stWarn = if ($b.WarnCount -gt 0) { $STYLE_NUM_WARN } else { $stNum }

            $rngS = if ($b.Ranges.Count -gt 0) { $b.Ranges -join " / " } elseif (-not $b.SeenSection) { "-" } else { "All (JST)" }
            $cntS = if ($b.Containers.Count -gt 0) { $b.Containers -join ", " } elseif (-not $b.SeenSection) { "-" } else { "front|back" }
            $svcS = if ($b.Services.Count -gt 0) { $b.Services -join ", " } elseif (-not $b.SeenSection) { "-" } else { "all" }
            $extraS = "{0:N0} $unitKen" -f $b.InfoCount
            if ($lt -eq "gc.log" -and $b.SeenSection) {
                $extraS += (" (MaxPause: {0:F1}ms)" -f $b.MaxPauseMs)
            }

            [void]$sbRows.Append("<row r=`"$curR`" ht=`"21`" customHeight=`"1`">")
            [void]$sbRows.Append((New-StrCell -Ref "A$curR" -StyleId $stC -Text $lt))
            [void]$sbRows.Append((New-StrCell -Ref "B$curR" -StyleId $stC -Text $rngS))
            [void]$sbRows.Append((New-StrCell -Ref "C$curR" -StyleId $stC -Text $cntS))
            [void]$sbRows.Append((New-StrCell -Ref "D$curR" -StyleId $stL -Text $svcS))
            [void]$sbRows.Append((New-NumCell -Ref "E$curR" -StyleId $stNum -Value $b.EntryCount))
            [void]$sbRows.Append((New-NumCell -Ref "F$curR" -StyleId $stNum -Value $b.Rows.Count))
            [void]$sbRows.Append((New-NumCell -Ref "G$curR" -StyleId $stErr -Value $b.ErrCount))
            [void]$sbRows.Append((New-NumCell -Ref "H$curR" -StyleId $stWarn -Value $b.WarnCount))
            [void]$sbRows.Append((New-StrCell -Ref "I$curR" -StyleId $stC -Text $extraS))
            [void]$sbRows.Append('</row>')
            $curR++
        }

        [void]$sbRows.Append("<row r=`"$curR`" ht=`"12`" customHeight=`"1`"/>")
        $curR++

        $sec2Title = ConvertFrom-UnicodeEscapes "\u25a0 2. \u30b3\u30f3\u30c6\u30ca / ECS\u30b5\u30fc\u30d3\u30b9\u5225 \u5185\u8a33 (Breakdown by Container / ECS Service)"
        [void]$sbRows.Append("<row r=`"$curR`" ht=`"24`" customHeight=`"1`">")
        [void]$sbRows.Append((New-StrCell -Ref "A$curR" -StyleId $STYLE_SECTION_HDR -Text $sec2Title))
        for ($c = 2; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)$curR" -StyleId $STYLE_SECTION_HDR -Text ""))
        }
        [void]$sbRows.Append('</row>')
        $merges.Add("A${curR}:${lastColLet}${curR}")
        $curR++

        $sec2Headers = @(
            (ConvertFrom-UnicodeEscapes "\u30ed\u30b0\u7a2e\u5225 (Sheet)"),
            (ConvertFrom-UnicodeEscapes "Cntnr (\u30b3\u30f3\u30c6\u30ca)"),
            (ConvertFrom-UnicodeEscapes "ECS Svc (\u30b5\u30fc\u30d3\u30b9)"),
            (ConvertFrom-UnicodeEscapes "Total \u4ef6\u6570"),
            (ConvertFrom-UnicodeEscapes "ERROR \u4ef6\u6570"),
            (ConvertFrom-UnicodeEscapes "WARN \u4ef6\u6570"),
            (ConvertFrom-UnicodeEscapes "INFO/\u4ed6 \u4ef6\u6570"),
            (ConvertFrom-UnicodeEscapes "\u6700\u5927GC\u505c\u6b62\u6642\u9593 (MaxPause)"),
            (ConvertFrom-UnicodeEscapes "\u30b9\u30c6\u30fc\u30bf\u30b9\u5224\u5b9a")
        )
        [void]$sbRows.Append("<row r=`"$curR`" ht=`"25`" customHeight=`"1`">")
        for ($c = 1; $c -le $numCols; $c++) {
            [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)$curR" -StyleId $STYLE_TBL_HEADER_C -Text $sec2Headers[$c - 1]))
        }
        [void]$sbRows.Append('</row>')
        $curR++

        $bdIdx = 0
        $lblErrDet = ConvertFrom-UnicodeEscapes "ERROR \u691c\u51fa"
        $lblWrnDet = ConvertFrom-UnicodeEscapes "WARN \u691c\u51fa"
        foreach ($lt in $LogTypesOrder) {
            $b = $Buckets[$lt]
            $sortedKeys = @($b.CsStats.Keys | Sort-Object)
            foreach ($k in $sortedKeys) {
                $stData = $b.CsStats[$k]
                $even = (($bdIdx % 2) -eq 0)
                $bdIdx++
                $stC = if ($even) { $STYLE_INFO_EVEN_C } else { $STYLE_INFO_ODD_C }
                $stNum = if ($even) { $STYLE_NUM_EVEN } else { $STYLE_NUM_ODD }
                $stErr = if ($stData.Err -gt 0) { $STYLE_NUM_ERR } else { $stNum }
                $stWarn = if ($stData.Warn -gt 0) { $STYLE_NUM_WARN } else { $stNum }

                if ($stData.Err -gt 0) {
                    $statusLabel = $lblErrDet
                    $stBadge = $STYLE_ERR_BADGE_C
                }
                elseif ($stData.Warn -gt 0) {
                    $statusLabel = $lblWrnDet
                    $stBadge = $STYLE_WARN_BADGE_C
                }
                else {
                    $statusLabel = "NORMAL"
                    $stBadge = $stC
                }

                $mpStr = if ($lt -eq "gc.log" -and $stData.MaxPause -gt 0) { "{0:F1}ms" -f $stData.MaxPause } else { "-" }

                [void]$sbRows.Append("<row r=`"$curR`" ht=`"20`" customHeight=`"1`">")
                [void]$sbRows.Append((New-StrCell -Ref "A$curR" -StyleId $stC -Text $lt))
                [void]$sbRows.Append((New-StrCell -Ref "B$curR" -StyleId $stC -Text $stData.Cnt))
                [void]$sbRows.Append((New-StrCell -Ref "C$curR" -StyleId $stC -Text $stData.Svc))
                [void]$sbRows.Append((New-NumCell -Ref "D$curR" -StyleId $stNum -Value $stData.Total))
                [void]$sbRows.Append((New-NumCell -Ref "E$curR" -StyleId $stErr -Value $stData.Err))
                [void]$sbRows.Append((New-NumCell -Ref "F$curR" -StyleId $stWarn -Value $stData.Warn))
                [void]$sbRows.Append((New-NumCell -Ref "G$curR" -StyleId $stNum -Value $stData.Info))
                [void]$sbRows.Append((New-StrCell -Ref "H$curR" -StyleId $stC -Text $mpStr))
                [void]$sbRows.Append((New-StrCell -Ref "I$curR" -StyleId $stBadge -Text $statusLabel))
                [void]$sbRows.Append('</row>')
                $curR++
            }
        }

        if ($bdIdx -eq 0) {
            [void]$sbRows.Append("<row r=`"$curR`" ht=`"20`" customHeight=`"1`">")
            for ($c = 1; $c -le $numCols; $c++) {
                [void]$sbRows.Append((New-StrCell -Ref "$(Get-ColumnLetter $c)$curR" -StyleId $STYLE_INFO_EVEN_C -Text "-"))
            }
            [void]$sbRows.Append('</row>')
            $curR++
        }

        $colWidths = @(20, 44, 18, 36, 16, 26, 16, 24, 30)
        $colsXml = "<cols>" + ((1..$numCols | ForEach-Object { "<col min=`"$_`" max=`"$_`" width=`"$($colWidths[$_ - 1])`" customWidth=`"1`"/>" }) -join "") + "</cols>"
        $mergeXml = "<mergeCells count=`"$($merges.Count)`">" + (($merges | ForEach-Object { "<mergeCell ref=`"$_`"/>" }) -join "") + "</mergeCells>"

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
            '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
            "<dimension ref=`"A1:${lastColLet}$($curR - 1)`"/>" +
            "<sheetViews><sheetView tabSelected=`"$TabSelected`" showGridLines=`"1`" workbookViewId=`"0`"/></sheetViews>" +
            '<sheetFormatPr defaultRowHeight="20" customHeight="1"/>' +
            $colsXml +
            "<sheetData>$($sbRows.ToString())</sheetData>" +
            $mergeXml +
            '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
            '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
            '</worksheet>'
        )
    }

    function Build-SharedStringsXml {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append("<sst xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" count=`"$script:SstTotalRefs`" uniqueCount=`"$($SstStrings.Count)`">")
        foreach ($s in $SstStrings) {
            $esc = Escape-XmlText -Text $s
            if ($s.StartsWith(" ") -or $s.EndsWith(" ") -or $s.Contains("`n") -or $s.Contains("`t") -or $s.Contains("  ")) {
                [void]$sb.Append("<si><t xml:space=`"preserve`">$esc</t></si>")
            }
            else {
                [void]$sb.Append("<si><t>$esc</t></si>")
            }
        }
        [void]$sb.Append('</sst>')
        return $sb.ToString()
    }

    # ==========================================================================
    # 5. Assemble OpenXML (.xlsx) Package
    # ==========================================================================
    $ActiveTypes = @($LogTypesOrder | Where-Object { $Buckets[$_].SeenSection -or ($Buckets[$_].Rows.Count -gt 0) })
    if ($ActiveTypes.Count -eq 0) {
        $ActiveTypes = $LogTypesOrder
    }

    $SheetsInfo = New-Object System.Collections.Generic.List[object]
    $summaryXml = Build-SummarySheetXml -Buckets $Buckets -SourceFiles $ResolvedFiles -TabSelected 1
    $SheetsInfo.Add([PSCustomObject]@{ Name = "Summary"; Xml = $summaryXml })

    foreach ($lt in $ActiveTypes) {
        $detailXml = Build-DetailSheetXml -Bucket $Buckets[$lt] -TabSelected 0
        $SheetsInfo.Add([PSCustomObject]@{ Name = $lt; Xml = $detailXml })
    }

    $ctOverrides = New-Object System.Collections.Generic.List[string]
    $ctOverrides.Add('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    $ctOverrides.Add('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    $ctOverrides.Add('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    for ($idx = 1; $idx -le $SheetsInfo.Count; $idx++) {
        $ctOverrides.Add("<Override PartName=`"/xl/worksheets/sheet${idx}.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
    }
    $contentTypesXml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        ($ctOverrides -join '') +
        '</Types>'
    )

    $rootRelsXml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>'
    )

    $wbSheets = New-Object System.Collections.Generic.List[string]
    $wbRels   = New-Object System.Collections.Generic.List[string]
    for ($idx = 1; $idx -le $SheetsInfo.Count; $idx++) {
        $escName = Escape-XmlText -Text $SheetsInfo[$idx - 1].Name
        $wbSheets.Add("<sheet name=`"$escName`" sheetId=`"$idx`" r:id=`"rId$idx`"/>")
        $wbRels.Add("<Relationship Id=`"rId$idx`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet${idx}.xml`"/>")
    }
    $styleRid = $SheetsInfo.Count + 1
    $sstRid   = $SheetsInfo.Count + 2
    $wbRels.Add("<Relationship Id=`"rId$styleRid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>")
    $wbRels.Add("<Relationship Id=`"rId$sstRid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings`" Target=`"sharedStrings.xml`"/>")

    $workbookXml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16200"/></bookViews>' +
        "<sheets>$($wbSheets -join '')</sheets>" +
        '</workbook>'
    )

    $workbookRelsXml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        ($wbRels -join '') +
        '</Relationships>'
    )

    $stylesXml = Build-StylesXml
    $sstXml    = Build-SharedStringsXml

    $outDir = [System.IO.Path]::GetDirectoryName($OutputPath)
    if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
        [void](New-Item -ItemType Directory -Path $outDir -Force)
    }
    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $fs = New-Object System.IO.FileStream($OutputPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            function Write-ZipEntry {
                param($ZipArchive, [string]$EntryPath, [string]$Content, $Encoding)
                $entry = $ZipArchive.CreateEntry($EntryPath, [System.IO.Compression.CompressionLevel]::Optimal)
                $stream = $entry.Open()
                try {
                    $bytes = $Encoding.GetBytes($Content)
                    $stream.Write($bytes, 0, $bytes.Length)
                }
                finally {
                    $stream.Dispose()
                }
            }

            Write-ZipEntry -ZipArchive $zip -EntryPath "[Content_Types].xml"        -Content $contentTypesXml -Encoding $utf8NoBom
            Write-ZipEntry -ZipArchive $zip -EntryPath "_rels/.rels"                -Content $rootRelsXml    -Encoding $utf8NoBom
            Write-ZipEntry -ZipArchive $zip -EntryPath "xl/workbook.xml"            -Content $workbookXml    -Encoding $utf8NoBom
            Write-ZipEntry -ZipArchive $zip -EntryPath "xl/_rels/workbook.xml.rels" -Content $workbookRelsXml -Encoding $utf8NoBom
            Write-ZipEntry -ZipArchive $zip -EntryPath "xl/styles.xml"              -Content $stylesXml      -Encoding $utf8NoBom
            Write-ZipEntry -ZipArchive $zip -EntryPath "xl/sharedStrings.xml"       -Content $sstXml         -Encoding $utf8NoBom

            for ($idx = 1; $idx -le $SheetsInfo.Count; $idx++) {
                Write-ZipEntry -ZipArchive $zip -EntryPath "xl/worksheets/sheet${idx}.xml" -Content $SheetsInfo[$idx - 1].Xml -Encoding $utf8NoBom
            }
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $fs.Dispose()
    }

    Write-Host "================================================================================"
    Write-Host " JBoss EAP Log Excel Report Generator (Meiryo UI / Monotone)"
    Write-Host "================================================================================"
    Write-Host " Output Excel : $OutputPath"
    Write-Host " Source Files : $($ResolvedFiles.Count) file(s)"
    foreach ($lt in $LogTypesOrder) {
        $b = $Buckets[$lt]
        if ($b.SeenSection -or ($b.Rows.Count -gt 0)) {
            $lineInfo = ("   - {0,-14} : Events={1,-5} Lines={2,-5} (ERROR={3}, WARN={4}, INFO={5})" -f $lt, $b.EntryCount, $b.Rows.Count, $b.ErrCount, $b.WarnCount, $b.InfoCount)
            Write-Host $lineInfo
        }
    }
    Write-Host "================================================================================"
}
