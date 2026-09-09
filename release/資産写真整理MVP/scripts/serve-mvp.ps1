param(
    [Parameter(Mandatory = $true)]
    [string]$AppRoot,
    [int]$Port = 8765,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
$resolvedRoot = (Resolve-Path -LiteralPath $AppRoot).Path
$prefix = "http://127.0.0.1:$Port/"
$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
$utf8 = [Text.UTF8Encoding]::new($false)
$sessionBase = Join-Path ([IO.Path]::GetTempPath()) 'sm-photo-organizer'
New-Item -ItemType Directory -Path $sessionBase -Force | Out-Null

function Send-Response {
    param(
        [Net.Sockets.NetworkStream]$Stream,
        [int]$StatusCode,
        [string]$StatusText,
        [string]$ContentType,
        [byte[]]$Body,
        [bool]$HeadOnly = $false
    )
    $header = "HTTP/1.1 $StatusCode $StatusText`r`nContent-Type: $ContentType`r`nContent-Length: $($Body.Length)`r`nConnection: close`r`nCache-Control: no-store`r`n`r`n"
    $headerBytes = $utf8.GetBytes($header)
    $Stream.Write($headerBytes, 0, $headerBytes.Length)
    if (-not $HeadOnly) { $Stream.Write($Body, 0, $Body.Length) }
}

function Send-JsonError {
    param([Net.Sockets.NetworkStream]$Stream, [int]$StatusCode, [string]$Message)
    $json = @{ message = $Message } | ConvertTo-Json -Compress
    Send-Response $Stream $StatusCode 'Error' 'application/json; charset=utf-8' $utf8.GetBytes($json)
}

function Read-Request {
    param([Net.Sockets.NetworkStream]$Stream)
    $headerStream = [IO.MemoryStream]::new()
    $state = 0
    while ($state -lt 4) {
        $value = $Stream.ReadByte()
        if ($value -lt 0) { throw 'Connection closed before headers were complete.' }
        $headerStream.WriteByte([byte]$value)
        if ($headerStream.Length -gt 32768) { throw 'Request headers are too large.' }
        if (($state -eq 0 -or $state -eq 2) -and $value -eq 13) { $state++ }
        elseif (($state -eq 1 -or $state -eq 3) -and $value -eq 10) { $state++ }
        elseif ($value -eq 13) { $state = 1 }
        else { $state = 0 }
    }

    $headerBytes = $headerStream.ToArray()
    $headerText = [Text.Encoding]::ASCII.GetString($headerBytes, 0, $headerBytes.Length - 4)
    $lines = $headerText -split "`r`n"
    if ($lines[0] -notmatch '^(GET|HEAD|POST)\s+([^\s]+)\s+HTTP/') { throw 'Unsupported request.' }
    $method = $Matches[1]
    $requestPath = $Matches[2].Split('?')[0]
    $headers = @{}
    if ($lines.Length -gt 1) {
        foreach ($line in $lines[1..($lines.Length - 1)]) {
            $separator = $line.IndexOf(':')
            if ($separator -gt 0) { $headers[$line.Substring(0, $separator).Trim().ToLowerInvariant()] = $line.Substring($separator + 1).Trim() }
        }
    }

    $contentLength = if ($headers.ContainsKey('content-length')) { [int64]$headers['content-length'] } else { 0 }
    if ($contentLength -lt 0 -or $contentLength -gt 157286400) { throw 'Request body is too large.' }
    $body = [byte[]]::new($contentLength)
    $offset = 0
    while ($offset -lt $contentLength) {
        $read = $Stream.Read($body, $offset, $contentLength - $offset)
        if ($read -le 0) { throw 'Connection closed before the request body was complete.' }
        $offset += $read
    }
    return [PSCustomObject]@{ Method = $method; Path = [Uri]::UnescapeDataString($requestPath); Headers = $headers; Body = $body }
}

function Get-SessionRoot {
    param([string]$SessionId)
    if ($SessionId -notmatch '^[a-f0-9-]{20,64}$') { throw 'Invalid session identifier.' }
    return Join-Path $sessionBase $SessionId
}

try {
    $listener.Start()
    Write-Host "Photo organizer MVP started: $prefix"
    Write-Host 'Close this window to stop the app.'
    if (-not $NoOpen) { Start-Process $prefix }

    while ($true) {
        $client = $listener.AcceptTcpClient()
        try {
            $stream = $client.GetStream()
            $request = Read-Request $stream
            $method = $request.Method
            $path = $request.Path

            if ($method -eq 'POST' -and $path -match '^/api/session/([a-f0-9-]{20,64})/(health|album)$') {
                $sessionRoot = Get-SessionRoot $Matches[1]
                New-Item -ItemType Directory -Path $sessionRoot -Force | Out-Null
                $name = if ($Matches[2] -eq 'health') { 'health.xlsx' } else { 'album.xlsx' }
                if ($request.Body.Length -gt 104857600) { throw 'Workbook is too large.' }
                [IO.File]::WriteAllBytes((Join-Path $sessionRoot $name), $request.Body)
                Send-Response $stream 200 'OK' 'text/plain; charset=utf-8' $utf8.GetBytes('OK')
                continue
            }

            if ($method -eq 'POST' -and $path -match '^/api/session/([a-f0-9-]{20,64})/photo/(\d{4})$') {
                $sessionRoot = Get-SessionRoot $Matches[1]
                $photoRoot = Join-Path $sessionRoot 'photos'
                New-Item -ItemType Directory -Path $photoRoot -Force | Out-Null
                if ($request.Body.Length -gt 1048576) { throw 'Photo is too large.' }
                [IO.File]::WriteAllBytes((Join-Path $photoRoot ($Matches[2] + '.jpg')), $request.Body)
                Send-Response $stream 200 'OK' 'text/plain; charset=utf-8' $utf8.GetBytes('OK')
                continue
            }

            if ($method -eq 'POST' -and $path -match '^/api/session/([a-f0-9-]{20,64})/build$') {
                $sessionRoot = Get-SessionRoot $Matches[1]
                if (-not (Test-Path -LiteralPath (Join-Path $sessionRoot 'health.xlsx')) -or -not (Test-Path -LiteralPath (Join-Path $sessionRoot 'album.xlsx'))) { throw 'Required workbooks were not uploaded.' }
                if ($request.Body.Length -gt 5242880) { throw 'Manifest is too large.' }
                [IO.File]::WriteAllBytes((Join-Path $sessionRoot 'manifest.json'), $request.Body)
                & (Join-Path $PSScriptRoot 'build-photo-album.ps1') -SessionRoot $sessionRoot
                $outputPath = Join-Path $sessionRoot 'output.xlsx'
                if (-not (Test-Path -LiteralPath $outputPath)) { throw 'Excel output was not created.' }
                $outputBytes = [IO.File]::ReadAllBytes($outputPath)
                Send-Response $stream 200 'OK' 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' $outputBytes
                Remove-Item -LiteralPath $sessionRoot -Recurse -Force -ErrorAction SilentlyContinue
                continue
            }

            if ($method -notin @('GET', 'HEAD')) {
                Send-JsonError $stream 404 'Endpoint not found.'
                continue
            }

            $relativePath = $path.TrimStart('/')
            if ([string]::IsNullOrWhiteSpace($relativePath)) { $relativePath = 'index.html' }
            $candidate = [IO.Path]::GetFullPath((Join-Path $resolvedRoot $relativePath))
            $insideRoot = $candidate.Equals($resolvedRoot, [StringComparison]::OrdinalIgnoreCase) -or
                $candidate.StartsWith($resolvedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
            if (-not $insideRoot -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
                Send-Response $stream 404 'Not Found' 'text/plain; charset=utf-8' $utf8.GetBytes('Not found') ($method -eq 'HEAD')
                continue
            }
            $body = [IO.File]::ReadAllBytes($candidate)
            $contentType = if ([IO.Path]::GetExtension($candidate) -eq '.html') { 'text/html; charset=utf-8' } else { 'application/octet-stream' }
            Send-Response $stream 200 'OK' $contentType $body ($method -eq 'HEAD')
        }
        catch {
            Write-Warning $_.Exception.Message
            try { Send-JsonError $stream 500 $_.Exception.Message } catch { }
        }
        finally {
            $client.Close()
        }
    }
}
finally {
    $listener.Stop()
}
