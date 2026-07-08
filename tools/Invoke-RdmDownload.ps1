param(
    [string]$Engineers = "",
    [string]$EngineersFile = "",
    [string]$EngineerIdsCsv = "",
    [ValidateSet("Stop", "Skip")]
    [string]$OnMissing = "Stop",
    [string]$StartDate = "2023/10/01",
    [string]$EndDate = "",
    [int]$PageSize = 200,
    [int]$TimeoutSec = 90,
    [int]$MaxRetries = 3,
    [int]$RetryDelaySec = 5,
    [int]$MaxPages = 0,
    [string]$OutputDir = "",
    [switch]$AllFromIds,
    [switch]$DryRun,
    [switch]$NoRawJson,
    [switch]$NoAutoResolveMissing,
    [switch]$UseSystemProxy
)

$ErrorActionPreference = "Stop"

if (-not $EngineersFile) {
    $EngineersFile = Join-Path $PSScriptRoot "rdm_engineers.txt"
}
if (-not $EngineerIdsCsv) {
    $EngineerIdsCsv = Join-Path $PSScriptRoot "rdm_engineer_ids.csv"
}
if (-not $EndDate) {
    $EndDate = (Get-Date).ToString("yyyy/MM/dd")
}
if (-not $OutputDir) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputDir = Join-Path $PSScriptRoot "..\outputs\rdm_download_$stamp"
}

if ($PageSize -le 0) {
    throw "PageSize must be greater than 0."
}
if ($TimeoutSec -le 0) {
    throw "TimeoutSec must be greater than 0."
}
if ($MaxRetries -le 0) {
    throw "MaxRetries must be greater than 0."
}

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }

function Normalize-Text {
    param([object]$Value)
    if ($null -eq $Value) {
        return ""
    }
    return ([string]$Value) -replace "<[^>]*>", ""
}

function Encode-QueryValue {
    param([string]$Value)
    return [System.Uri]::EscapeDataString($Value)
}

function Load-EngineerIdMap {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Engineer ID CSV not found: $Path"
    }

    $map = @{}
    $rows = Import-Csv -LiteralPath $Path -Encoding UTF8
    foreach ($row in $rows) {
        $name = ([string]$row.name).Trim()
        $idText = ([string]$row.engineer_id).Trim()
        if (-not $name -or -not $idText) {
            continue
        }
        $id = 0
        if (-not [int]::TryParse($idText, [ref]$id)) {
            throw "Invalid engineer_id for $name in $Path`: $idText"
        }
        $map[$name] = $id
    }
    return $map
}

function Split-EngineerTokens {
    param([string]$Text)
    if (-not $Text) {
        return @()
    }
    return @($Text -split "[\s,，;；]+" | Where-Object { $_ })
}

function Resolve-EngineerInput {
    param(
        [string]$EngineersText,
        [string]$File,
        [hashtable]$IdMap,
        [switch]$All
    )

    if ($All) {
        return @($IdMap.Keys | Sort-Object)
    }

    $tokens = @()
    if ($EngineersText) {
        $tokens = Split-EngineerTokens $EngineersText
    } else {
        if (-not (Test-Path -LiteralPath $File)) {
            throw "Engineers file not found: $File"
        }
        $raw = (Get-Content -LiteralPath $File -Raw -Encoding UTF8) -replace ([string][char]0xFEFF), ""
        $tokens = Split-EngineerTokens $raw
    }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($token in $tokens) {
        if ($token -match "^(.+?)[:=](\d+)$") {
            $name = $Matches[1].Trim()
            $id = [int]$Matches[2]
            if ($name) {
                $IdMap[$name] = $id
                $names.Add($name)
            }
        } else {
            $names.Add($token.Trim())
        }
    }

    return @($names | Where-Object { $_ } | Select-Object -Unique)
}

function Normalize-DateRange {
    param(
        [string]$FromDate,
        [string]$ToDate
    )

    $dateRegex = "^\d{4}/\d{2}/\d{2}$"
    if ($FromDate -eq "0") {
        return [pscustomobject]@{
            StartDate = "0"
            EndDate = $ToDate
            QueryValue = ""
            Display = "all"
        }
    }
    if ($ToDate -eq "0") {
        $ToDate = (Get-Date).ToString("yyyy/MM/dd")
    }
    if ($FromDate -notmatch $dateRegex) {
        throw "StartDate must be yyyy/MM/dd or 0. Actual: $FromDate"
    }
    if ($ToDate -notmatch $dateRegex) {
        throw "EndDate must be yyyy/MM/dd or 0. Actual: $ToDate"
    }

    $encodedStart = Encode-QueryValue $FromDate
    $encodedEnd = Encode-QueryValue $ToDate
    return [pscustomobject]@{
        StartDate = $FromDate
        EndDate = $ToDate
        QueryValue = "$encodedStart+-+$encodedEnd"
        Display = "$FromDate to $ToDate"
    }
}

function Build-RdmUrl {
    param(
        [int]$EngineerId,
        [string]$DateQueryValue,
        [int]$Size,
        [int]$Offset
    )

    $baseUrl = [Environment]::GetEnvironmentVariable("FIRE_TV_RDM_BASE_URL", "Process")
    if (-not $baseUrl) { $baseUrl = [Environment]::GetEnvironmentVariable("FIRE_TV_RDM_BASE_URL", "User") }
    if (-not $baseUrl) { throw "FIRE_TV_RDM_BASE_URL is required." }
    $ts = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()
    $parts = @(
        "pageSize=$Size",
        "pageOffset=$Offset",
        "pageSearch=",
        "query%5B0%5D%5Bname%5D=engineer_id",
        "query%5B0%5D%5Bvalue%5D=$EngineerId",
        "query%5B1%5D%5Bname%5D=start_time",
        "query%5B1%5D%5Bvalue%5D=$DateQueryValue",
        "query%5B2%5D%5Bname%5D=estimated_time_range",
        "query%5B2%5D%5Bvalue%5D=",
        "query%5B3%5D%5Bname%5D=projectID",
        "query%5B3%5D%5Bvalue%5D=",
        "_=$ts"
    )
    return $baseUrl + "?" + ($parts -join "&")
}

function Build-RdmEngineerSearchUrl {
    param(
        [string]$EngineerName,
        [int]$Size = 50
    )

    $baseUrl = [Environment]::GetEnvironmentVariable("FIRE_TV_RDM_BASE_URL", "Process")
    if (-not $baseUrl) { $baseUrl = [Environment]::GetEnvironmentVariable("FIRE_TV_RDM_BASE_URL", "User") }
    if (-not $baseUrl) { throw "FIRE_TV_RDM_BASE_URL is required." }
    $ts = [DateTimeOffset]::Now.ToUnixTimeMilliseconds()
    $parts = @(
        "pageSize=$Size",
        "pageOffset=0",
        "pageSearch=$(Encode-QueryValue $EngineerName)",
        "query%5B0%5D%5Bname%5D=engineer_id",
        "query%5B0%5D%5Bvalue%5D=",
        "query%5B1%5D%5Bname%5D=start_time",
        "query%5B1%5D%5Bvalue%5D=",
        "query%5B2%5D%5Bname%5D=estimated_time_range",
        "query%5B2%5D%5Bvalue%5D=",
        "query%5B3%5D%5Bname%5D=projectID",
        "query%5B3%5D%5Bvalue%5D=",
        "_=$ts"
    )
    return $baseUrl + "?" + ($parts -join "&")
}

function Invoke-DirectGet {
    param(
        [string]$Url,
        [int]$TimeoutSeconds,
        [switch]$UseProxy
    )

    $request = [System.Net.HttpWebRequest] [System.Net.WebRequest]::Create($Url)
    $request.Method = "GET"
    $request.Timeout = $TimeoutSeconds * 1000
    $request.ReadWriteTimeout = $TimeoutSeconds * 1000
    $request.UserAgent = "RdmDownload/1.0"
    $request.Accept = "application/json,text/javascript,*/*;q=0.01"
    $request.Headers["X-Requested-With"] = "XMLHttpRequest"
    $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    if (-not $UseProxy) {
        $request.Proxy = $null
    }

    $response = $null
    $reader = $null
    try {
        $response = [System.Net.HttpWebResponse] $request.GetResponse()
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream(), [System.Text.Encoding]::UTF8)
        $content = $reader.ReadToEnd()
        return [pscustomobject]@{
            StatusCode = [int]$response.StatusCode
            Content = $content
        }
    } finally {
        if ($reader) {
            $reader.Dispose()
        }
        if ($response) {
            $response.Dispose()
        }
    }
}

function Get-RdmPage {
    param(
        [string]$Url,
        [string]$EngineerName,
        [int]$Offset
    )

    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $result = Invoke-DirectGet -Url $Url -TimeoutSeconds $TimeoutSec -UseProxy:$UseSystemProxy
            $parsed = $result.Content | ConvertFrom-Json
            if ($null -eq $parsed.total -or $null -eq $parsed.rows) {
                throw "Unexpected response shape. HTTP $($result.StatusCode), content length $($result.Content.Length)."
            }
            return [pscustomobject]@{
                Total = [int]$parsed.total
                Rows = @($parsed.rows)
                RawContent = $result.Content
                HttpStatus = $result.StatusCode
            }
        } catch {
            if ($attempt -ge $MaxRetries) {
                throw "Download failed for $EngineerName offset $Offset after $MaxRetries attempts: $($_.Exception.Message)"
            }
            $wait = $RetryDelaySec * $attempt
            Write-Warning "Request failed for $EngineerName offset $Offset, attempt $attempt/$MaxRetries. Retry in $wait seconds. $($_.Exception.Message)"
            Start-Sleep -Seconds $wait
        }
    }
}

function Find-RdmEngineerId {
    param([string]$EngineerName)

    $url = Build-RdmEngineerSearchUrl -EngineerName $EngineerName
    $page = Get-RdmPage -Url $url -EngineerName $EngineerName -Offset 0
    $candidateIds = New-Object System.Collections.Generic.HashSet[int]

    foreach ($row in @($page.Rows)) {
        $names = @((Normalize-Text $row.engineer) -split "\s+" | Where-Object { $_ })
        $ids = @($row.engineer_id | ForEach-Object { [int]$_ })

        if ($names.Count -eq 1 -and $names[0] -eq $EngineerName -and $ids.Count -eq 1) {
            [void]$candidateIds.Add($ids[0])
            continue
        }

        if ($names.Count -eq $ids.Count) {
            for ($i = 0; $i -lt $names.Count; $i++) {
                if ($names[$i] -eq $EngineerName) {
                    [void]$candidateIds.Add($ids[$i])
                }
            }
        }
    }

    $resolved = @($candidateIds)
    if ($resolved.Count -eq 1) {
        return [int]$resolved[0]
    }
    if ($resolved.Count -gt 1) {
        throw "Multiple engineer_id values found for $EngineerName`: $($resolved -join ', ')"
    }
    return $null
}

function Save-EngineerIdMapEntry {
    param(
        [string]$Path,
        [string]$EngineerName,
        [int]$EngineerId
    )

    $rows = @(Import-Csv -LiteralPath $Path -Encoding UTF8)
    if (@($rows | Where-Object { ([string]$_.name).Trim() -eq $EngineerName }).Count -gt 0) {
        return
    }
    $rows += [pscustomobject]@{
        name = $EngineerName
        engineer_id = $EngineerId
    }
    Export-CsvUtf8Bom -Path $Path -Rows $rows
}

function Convert-RdmRowToCsvObject {
    param([object]$Row)
    $props = $Row.PSObject.Properties
    return [pscustomobject][ordered]@{
        "订单编号" = Normalize-Text ($props["number"].Value)
        "订单信息" = Normalize-Text ($props["software"].Value)
        "项目编号" = Normalize-Text ($props["items"].Value)
        "业务" = Normalize-Text ($props["business"].Value)
        "数量" = Normalize-Text ($props["order_count"].Value)
        "订单类" = Normalize-Text ($props["order"].Value)
        "BOM" = Normalize-Text ($props["bom_number"].Value)
        "机芯" = Normalize-Text ($props["movement_name"].Value)
        "FORMAT" = Normalize-Text ($props["order_format"].Value)
        "EMMC" = Normalize-Text ($props["memc"].Value)
        "需求时间" = Normalize-Text ($props["demand_time"].Value)
        "发布时间" = Normalize-Text ($props["estimated_time"].Value)
        "状态" = Normalize-Text ($props["status"].Value)
        "SPM" = Normalize-Text ($props["engineer"].Value)
        "版本" = Normalize-Text ($props["_versions"].Value)
        "ID" = Normalize-Text ($props["projectID"].Value)
        "备注" = Normalize-Text ($props["remark"].Value)
    }
}

function Initialize-Csv {
    param([string]$Path)
    $headers = @("订单编号", "订单信息", "项目编号", "业务", "数量", "订单类", "BOM", "机芯", "FORMAT", "EMMC", "需求时间", "发布时间", "状态", "SPM", "版本", "ID", "备注")
    $encoding = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, ($headers -join ",") + [Environment]::NewLine, $encoding)
}

function Export-CsvUtf8Bom {
    param(
        [string]$Path,
        [object[]]$Rows
    )
    $encoding = New-Object System.Text.UTF8Encoding($true)
    if ($Rows.Count -eq 0) {
        [System.IO.File]::WriteAllText($Path, "", $encoding)
        return
    }
    $csv = $Rows | ConvertTo-Csv -NoTypeInformation
    [System.IO.File]::WriteAllText($Path, (($csv -join [Environment]::NewLine) + [Environment]::NewLine), $encoding)
}

function Export-CsvForExcel {
    param(
        [string]$Path,
        [object[]]$Rows
    )
    if ($Rows.Count -eq 0) {
        [System.IO.File]::WriteAllText($Path, "", [System.Text.Encoding]::Default)
        return
    }
    $csv = $Rows | ConvertTo-Csv -NoTypeInformation
    [System.IO.File]::WriteAllText($Path, (($csv -join [Environment]::NewLine) + [Environment]::NewLine), [System.Text.Encoding]::Default)
}

function Append-CsvRows {
    param(
        [string]$Path,
        [object[]]$Rows
    )
    if ($Rows.Count -eq 0) {
        return
    }
    $Rows | Export-Csv -LiteralPath $Path -Append -NoTypeInformation -Encoding UTF8
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$csvPath = Join-Path $OutputDir "data.csv"
$rawJsonlPath = Join-Path $OutputDir "raw_rows.jsonl"
$summaryPath = Join-Path $OutputDir "summary.csv"
$missingPath = Join-Path $OutputDir "missing_engineer_ids.csv"
$missingExcelPath = Join-Path $OutputDir "missing_engineer_ids_excel.csv"
$urlLogPath = Join-Path $OutputDir "urls.txt"
$runLogPath = Join-Path $OutputDir "run.log"
$latestPath = Join-Path $PSScriptRoot "..\outputs\latest_rdm_download_path.txt"

$resolvedOutputDir = (Resolve-Path -LiteralPath $OutputDir).Path
$latestEncoding = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText((Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\outputs")).Path "latest_rdm_download_path.txt"), $resolvedOutputDir, $latestEncoding)

Start-Transcript -LiteralPath $runLogPath -Force | Out-Null
try {
    $idMap = Load-EngineerIdMap -Path $EngineerIdsCsv
    $selectedNames = Resolve-EngineerInput -EngineersText $Engineers -File $EngineersFile -IdMap $idMap -All:$AllFromIds
    if ($selectedNames.Count -eq 0) {
        throw "No engineers selected."
    }

    $missing = @($selectedNames | Where-Object { -not $idMap.ContainsKey($_) })
    if ($missing.Count -gt 0 -and -not $NoAutoResolveMissing) {
        foreach ($name in $missing) {
            try {
                Write-Host "Resolve engineer_id from RDM history: $name"
                $resolvedId = Find-RdmEngineerId -EngineerName $name
                if ($null -ne $resolvedId) {
                    $idMap[$name] = $resolvedId
                    Save-EngineerIdMapEntry -Path $EngineerIdsCsv -EngineerName $name -EngineerId $resolvedId
                    Write-Host "  resolved: $name = $resolvedId"
                } else {
                    Write-Warning "No RDM history found for $name; engineer_id remains unresolved."
                }
            } catch {
                Write-Warning "Could not auto-resolve engineer_id for $name`: $($_.Exception.Message)"
            }
        }
        $missing = @($selectedNames | Where-Object { -not $idMap.ContainsKey($_) })
    }
    if ($missing.Count -gt 0) {
        $missingRows = $missing | ForEach-Object {
            [pscustomobject]@{
                name = $_
                reason = "engineer_id not found. Add it to tools\rdm_engineer_ids.csv or input as Name:Id."
            }
        }
        $missingExcelRows = $missing | ForEach-Object {
            [pscustomobject]@{
                name = $_
                reason = "缺少 engineer_id，请补到 tools\rdm_engineer_ids.csv，或用 姓名:ID 临时输入"
            }
        }
        Export-CsvUtf8Bom -Path $missingPath -Rows @($missingRows)
        Export-CsvForExcel -Path $missingExcelPath -Rows @($missingExcelRows)
        if ($OnMissing -eq "Stop") {
            throw "Missing engineer_id for: $($missing -join ', '). Details: $missingPath"
        }
        Write-Warning "Skipping engineers without engineer_id: $($missing -join ', '). Details: $missingPath"
        Write-Warning "Excel-readable missing list: $missingExcelPath"
        $selectedNames = @($selectedNames | Where-Object { $idMap.ContainsKey($_) })
    }

    if ($selectedNames.Count -eq 0) {
        throw "No engineers with valid engineer_id."
    }

    $dateRange = Normalize-DateRange -FromDate $StartDate -ToDate $EndDate
    Initialize-Csv -Path $csvPath
    if (-not $NoRawJson) {
        Set-Content -LiteralPath $rawJsonlPath -Value "" -Encoding UTF8
    }
    Set-Content -LiteralPath $urlLogPath -Value "" -Encoding UTF8

    Write-Host "RDM downloader"
    Write-Host "Output dir: $resolvedOutputDir"
    Write-Host "Latest output path file: $latestPath"
    Write-Host "Date range: $($dateRange.Display)"
    Write-Host "Page size: $PageSize"
    Write-Host "Proxy: $(if ($UseSystemProxy) { 'system proxy' } else { 'direct' })"
    Write-Host "Engineers: $($selectedNames -join ' ')"

    if ($DryRun) {
        foreach ($name in $selectedNames) {
            $id = [int]$idMap[$name]
            $url = Build-RdmUrl -EngineerId $id -DateQueryValue $dateRange.QueryValue -Size $PageSize -Offset 0
            "$name,$id,$url" | Add-Content -LiteralPath $urlLogPath -Encoding UTF8
        }
        Write-Host "Dry run done. URL log: $urlLogPath"
        return
    }

    $summary = New-Object System.Collections.Generic.List[object]
    $grandDownloaded = 0
    $hadFailure = $false

    foreach ($name in $selectedNames) {
        $id = [int]$idMap[$name]
        $offset = 0
        $downloaded = 0
        $total = $null
        $pages = 0
        $status = "OK"
        $errorText = ""

        Write-Host "Start: $name ($id)"
        try {
            while ($true) {
                $url = Build-RdmUrl -EngineerId $id -DateQueryValue $dateRange.QueryValue -Size $PageSize -Offset $offset
                "$name,$id,$offset,$url" | Add-Content -LiteralPath $urlLogPath -Encoding UTF8

                $page = Get-RdmPage -Url $url -EngineerName $name -Offset $offset
                $rows = @($page.Rows)
                $total = [int]$page.Total
                $pages++

                if ($rows.Count -gt 0) {
                    $csvRows = @($rows | ForEach-Object { Convert-RdmRowToCsvObject $_ })
                    Append-CsvRows -Path $csvPath -Rows $csvRows
                    if (-not $NoRawJson) {
                        foreach ($row in $rows) {
                            $rawLine = [pscustomobject]@{
                                engineer_name = $name
                                engineer_id = $id
                                row = $row
                            } | ConvertTo-Json -Depth 30 -Compress
                            Add-Content -LiteralPath $rawJsonlPath -Value $rawLine -Encoding UTF8
                        }
                    }
                }

                $downloaded += $rows.Count
                $grandDownloaded += $rows.Count
                Write-Host ("  page {0}: +{1}, {2}/{3}" -f $pages, $rows.Count, $downloaded, $total)

                if ($rows.Count -eq 0 -or $downloaded -ge $total) {
                    break
                }
                if ($MaxPages -gt 0 -and $pages -ge $MaxPages) {
                    $status = "PARTIAL_MAX_PAGES"
                    break
                }
                $offset += $rows.Count
                Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 900)
            }
        } catch {
            $status = "FAILED"
            $errorText = $_.Exception.Message
            $hadFailure = $true
            Write-Warning $errorText
        }

        $summary.Add([pscustomobject]@{
            engineer = $name
            engineer_id = $id
            total_reported = $total
            downloaded_rows = $downloaded
            pages = $pages
            status = $status
            error = $errorText
        })

        if ($status -eq "FAILED") {
            break
        }
    }

    Export-CsvUtf8Bom -Path $summaryPath -Rows $summary.ToArray()
    Write-Host "Done."
    Write-Host "Downloaded rows: $grandDownloaded"
    Write-Host "CSV: $csvPath"
    Write-Host "Summary: $summaryPath"
    if (-not $NoRawJson) {
        Write-Host "Raw JSONL: $rawJsonlPath"
    }
    if ($hadFailure) {
        throw "One or more engineers failed. See summary: $summaryPath"
    }
} finally {
    try {
        Stop-Transcript | Out-Null
    } catch {}
}
