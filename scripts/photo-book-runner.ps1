param(
    [Parameter(Mandatory = $true)]
    [string]$InstructionPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Select-InputFile {
    param([string]$Title)
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Title = $Title
    $dialog.Filter = 'Excel ファイル (*.xlsx)|*.xlsx'
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'ファイル選択を取り消しました。' }
    return $dialog.FileName
}

function Select-InputFolder {
    param([string]$Description)
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'フォルダ選択を取り消しました。' }
    return $dialog.SelectedPath
}

function Convert-ToJpeg {
    param([string]$SourcePath, [string]$DestinationPath)
    $image = $null
    try {
        $image = [System.Drawing.Image]::FromFile($SourcePath)
        $image.Save($DestinationPath, [System.Drawing.Imaging.ImageFormat]::Jpeg)
    }
    finally {
        if ($image -ne $null) { $image.Dispose() }
    }
}

try {
    if (-not (Test-Path -LiteralPath $InstructionPath -PathType Leaf)) { throw '写真帳作成指示.json が見つかりません。展開したフォルダ内で実行してください。' }
    $healthSourcePath = Join-Path $PSScriptRoot 'health.xlsx'
    if (-not (Test-Path -LiteralPath $healthSourcePath -PathType Leaf)) { throw '健全度判定表が見つかりません。写真帳作成セットを保存し直してください。' }
    $instruction = Get-Content -LiteralPath $InstructionPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($instruction.kind -ne 'asset-photo-album-instruction' -or $instruction.version -ne 1) { throw 'この写真整理MVPで作成した写真帳作成セットではありません。' }
    if (-not @($instruction.photos).Count) { throw '写真帳へ貼り付ける写真がありません。' }

    $templatePath = Select-InputFile '写真を貼り付ける写真帳Excelを選択してください'
    $photoFolder = Select-InputFolder '元写真が入ったフォルダを選択してください（子フォルダも検索します）'
    $outputFolder = Select-InputFolder '写真帳の出力先フォルダを選択してください'

    $filesByName = @{}
    Get-ChildItem -LiteralPath $photoFolder -Recurse -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' } | ForEach-Object {
        $key = $_.Name.ToLowerInvariant()
        if (-not $filesByName.ContainsKey($key)) { $filesByName[$key] = @() }
        $filesByName[$key] += $_
    }

    $sessionRoot = Join-Path ([IO.Path]::GetTempPath()) ("sm-photo-book-" + [guid]::NewGuid().ToString())
    $photoRoot = Join-Path $sessionRoot 'photos'
    New-Item -ItemType Directory -Path $photoRoot -Force | Out-Null
    Copy-Item -LiteralPath $templatePath -Destination (Join-Path $sessionRoot 'album.xlsx') -Force
    Copy-Item -LiteralPath $healthSourcePath -Destination (Join-Path $sessionRoot 'health.xlsx') -Force

    $manifestPhotos = @()
    $index = 0
    foreach ($photo in @($instruction.photos)) {
        $photoName = [string]$photo.fileName
        $matches = @($filesByName[$photoName.ToLowerInvariant()])
        if ($matches.Count -eq 0) { throw "元写真が見つかりません：$($photo.fileName)" }
        if ($matches.Count -gt 1) { throw "同名の元写真が複数あります：$($photo.fileName)。写真フォルダを整理してから再実行してください。" }
        $index += 1
        $fileKey = $index.ToString('0000')
        try { Convert-ToJpeg -SourcePath $matches[0].FullName -DestinationPath (Join-Path $photoRoot ($fileKey + '.jpg')) }
        catch { throw "写真をJPEGへ変換できません：$($photo.fileName)。JPG、JPEG、PNG形式の写真を選択してください。" }
        $manifestPhotos += [PSCustomObject]@{
            fileKey = $fileKey
            sourceName = [string]$photo.fileName
            assetNumber = [string]$photo.assetNumber
            destinationType = [string]$photo.destinationType
            itemNumber = if ($null -eq $photo.itemNumber) { $null } else { [string]$photo.itemNumber }
            slotIndex = [int]$photo.slotIndex
        }
    }
    $manifest = [PSCustomObject]@{
        outputName = [IO.Path]::GetFileNameWithoutExtension($templatePath) + '_写真貼付済.xlsx'
        assets = @($instruction.assets)
        photos = $manifestPhotos
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $sessionRoot 'manifest.json') -Encoding UTF8

    & (Join-Path $PSScriptRoot 'build-photo-album.ps1') -SessionRoot $sessionRoot
    # 選択した写真帳の名前を残し、同じ分に複数回実行しても重複しないよう秒まで付ける。
    $outputName = [IO.Path]::GetFileNameWithoutExtension($templatePath) + '_写真貼付済_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx'
    $outputPath = Join-Path $outputFolder $outputName
    Copy-Item -LiteralPath (Join-Path $sessionRoot 'output.xlsx') -Destination $outputPath -Force
    [System.Windows.Forms.MessageBox]::Show("写真帳を作成しました。`n$outputPath", '資産写真整理MVP') | Out-Null
}
catch {
    $message = $_.Exception.Message
    $errorLogPath = Join-Path $PSScriptRoot 'photo-book-error.log'
    Set-Content -LiteralPath $errorLogPath -Value $message -Encoding UTF8
    [System.Windows.Forms.MessageBox]::Show("$message`n`n詳細を $errorLogPath に保存しました。", '写真帳作成エラー') | Out-Null
    Write-Error $message
}
finally {
    if ($sessionRoot -and (Test-Path -LiteralPath $sessionRoot)) { Remove-Item -LiteralPath $sessionRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
