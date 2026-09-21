param(
    [Parameter(Mandatory = $true)]
    [string]$SessionRoot
)

$ErrorActionPreference = 'Stop'
$healthPath = Join-Path $SessionRoot 'health.xlsx'
$templatePath = Join-Path $SessionRoot 'album.xlsx'
$manifestPath = Join-Path $SessionRoot 'manifest.json'
$outputPath = Join-Path $SessionRoot 'output.xlsx'
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json

Copy-Item -LiteralPath $templatePath -Destination $outputPath -Force

$excel = $null
$healthBook = $null
$albumBook = $null
$sourceNo11 = $null
$targetNo11 = $null
try {
    try { $excel = New-Object -ComObject Excel.Application }
    catch { throw 'Microsoft Excel is not installed or could not be started.' }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    $healthBook = $excel.Workbooks.Open($healthPath, 0, $true)
    $albumBook = $excel.Workbooks.Open($outputPath, 0, $false)
    $sourceNo11 = $healthBook.Worksheets.Item('No11')
    $targetNo11 = $albumBook.Worksheets.Item('No11')
    if ($targetNo11.ProtectContents) { throw 'The No11 sheet in the photo album is protected.' }

    # No11全体は置き換えず、点検結果1・2の値と入力規則だけを対応セルへ反映する。
    # 写真帳側の書式、数式、行構成、既存のセレクトボックスは維持する。
    $sourceUsed = $sourceNo11.UsedRange
    $targetUsed = $targetNo11.UsedRange
    $sourceValues = $sourceUsed.Value2
    $targetValues = $targetUsed.Value2
    # Excel COMのプロパティは環境によりObject[]として返ることがあるため、
    # 算術演算に使う値は最初に数値へ固定する。
    $sourceFirstRow = [int]$sourceUsed.Row
    $sourceFirstColumn = [int]$sourceUsed.Column
    $sourceRowCount = [int]$sourceUsed.Rows.Count
    $sourceColumnCount = [int]$sourceUsed.Columns.Count
    $targetFirstRow = [int]$targetUsed.Row
    $targetFirstColumn = [int]$targetUsed.Column
    $targetRowCount = [int]$targetUsed.Rows.Count
    $targetColumnCount = [int]$targetUsed.Columns.Count
    $normalizeHeader = {
        param($value)
        ([string]$value).Trim().Replace(' ', '').Replace('　', '').Replace('１', '1').Replace('２', '2')
    }
    $sourceHeaders = @{}
    $targetHeaders = @{}
    $sourceHeaderRow = 0
    $targetHeaderRow = 0
    foreach ($row in 1..([Math]::Min(30, $sourceRowCount))) {
        $candidate = @{}
        for ($column = 1; $column -le $sourceColumnCount; $column++) {
            $sourceHeader = & $normalizeHeader $sourceValues[$row, $column]
            if ($sourceHeader -in @('資産番号', '項目番号', '点検結果1', '点検結果2')) { $candidate[$sourceHeader] = [int]($sourceFirstColumn + $column - 1) }
        }
        if (@('資産番号', '項目番号', '点検結果1', '点検結果2') | Where-Object { -not $candidate.ContainsKey($_) }) { continue }
        $sourceHeaders = $candidate
        $sourceHeaderRow = $row
        break
    }
    foreach ($row in 1..([Math]::Min(30, $targetRowCount))) {
        $candidate = @{}
        for ($column = 1; $column -le $targetColumnCount; $column++) {
            $targetHeader = & $normalizeHeader $targetValues[$row, $column]
            if ($targetHeader -in @('資産番号', '項目番号', '点検結果1', '点検結果2')) { $candidate[$targetHeader] = [int]($targetFirstColumn + $column - 1) }
        }
        if (@('資産番号', '項目番号', '点検結果1', '点検結果2') | Where-Object { -not $candidate.ContainsKey($_) }) { continue }
        $targetHeaders = $candidate
        $targetHeaderRow = $row
        break
    }
    foreach ($requiredHeader in @('資産番号', '項目番号', '点検結果1', '点検結果2')) {
        if (-not $sourceHeaders.ContainsKey($requiredHeader) -or -not $targetHeaders.ContainsKey($requiredHeader)) {
            throw "No11シートに「$requiredHeader」の列が見つかりません。"
        }
    }

    $targetRows = @{}
    $targetAssetIndex = [int]$targetHeaders['資産番号'] - $targetFirstColumn + 1
    $targetItemIndex = [int]$targetHeaders['項目番号'] - $targetFirstColumn + 1
    for ($row = $targetHeaderRow + 1; $row -le $targetRowCount; $row++) {
        $assetNumber = ([string]$targetValues[$row, $targetAssetIndex]).Trim()
        $itemNumber = ([string]$targetValues[$row, $targetItemIndex]).Trim()
        if ($assetNumber -and $itemNumber) { $targetRows["$assetNumber|$itemNumber"] = [int]($targetFirstRow + $row - 1) }
    }
    $sourceAssetIndex = [int]$sourceHeaders['資産番号'] - $sourceFirstColumn + 1
    $sourceItemIndex = [int]$sourceHeaders['項目番号'] - $sourceFirstColumn + 1
    for ($row = $sourceHeaderRow + 1; $row -le $sourceRowCount; $row++) {
        $assetNumber = ([string]$sourceValues[$row, $sourceAssetIndex]).Trim()
        $itemNumber = ([string]$sourceValues[$row, $sourceItemIndex]).Trim()
        $targetRow = $targetRows["$assetNumber|$itemNumber"]
        if (-not $targetRow) { continue }
        foreach ($resultHeader in @('点検結果1', '点検結果2')) {
            $sourceColumn = $sourceHeaders[$resultHeader]
            $targetColumn = $targetHeaders[$resultHeader]
            $sourceCell = $sourceNo11.Cells.Item([int]($sourceFirstRow + $row - 1), $sourceColumn)
            $targetCell = $targetNo11.Cells.Item($targetRow, $targetColumn)
            # xlPasteValidation=6。入力規則だけをコピーし、書式・数式は触らない。
            $sourceCell.Copy()
            $targetCell.PasteSpecial(6)
            $targetCell.Value2 = $sourceCell.Value2
        }
    }
    $excel.CutCopyMode = 0
    $healthBook.Close($false)
    $healthBook = $null

    $albumBook.Save()
    $excel.CalculateFullRebuild()

    $sheetByAsset = @{}
    foreach ($sheet in $albumBook.Worksheets) {
        if ($sheet.Name -match '^\d+_\d+$') {
            $sheet.Calculate()
            $assetNumber = [string]$sheet.Range('AB4').Value2
            if ([string]::IsNullOrWhiteSpace($assetNumber)) { $assetNumber = [string]$sheet.Range('AO2').Value2 }
            if (-not [string]::IsNullOrWhiteSpace($assetNumber)) { $sheetByAsset[$assetNumber.Trim()] = $sheet }
        }
    }

    $photosByItem = @{}
    foreach ($photo in $manifest.photos) {
        $key = if ($photo.destinationType -eq 'full') { "$($photo.assetNumber)/full" } else { "$($photo.assetNumber)/$($photo.itemNumber)" }
        if (-not $photosByItem.ContainsKey($key)) { $photosByItem[$key] = @() }
        $photosByItem[$key] += $photo
    }

    foreach ($asset in $manifest.assets) {
        if (-not $sheetByAsset.ContainsKey([string]$asset.assetNumber)) { throw "No photo sheet found for asset $($asset.assetNumber)." }
        $sheet = $sheetByAsset[[string]$asset.assetNumber]
        $validItems = @{}
        foreach ($itemNumber in $asset.itemNumbers) { $validItems[[string]$itemNumber] = $true }

        $fullKey = "$($asset.assetNumber)/full"
        if ($photosByItem.ContainsKey($fullKey)) {
            $photo = $photosByItem[$fullKey][0]
            $photoPath = Join-Path (Join-Path $SessionRoot 'photos') ($photo.fileKey + '.jpg')
            $target = $sheet.Range('E5:S18')
            $shape = $sheet.Shapes.AddPicture($photoPath, $false, $true, $target.Left, $target.Top, -1, -1)
            $shape.LockAspectRatio = -1
            $shape.Width = [Math]::Min($target.Width, 260.7874)
            if ($shape.Height -gt $target.Height) { $shape.Height = $target.Height }
            $shape.Left = $target.Left
            $shape.Top = $target.Top
            $shape.Placement = 1
            $shape.Name = "SM_$($asset.assetNumber)_full"
        }

        $itemRows = @()
        for ($headerRow = 21; $headerRow -le $sheet.UsedRange.Rows.Count; $headerRow += 14) {
            $itemNumber = [string]$sheet.Range("AP$headerRow").Value2
            if (-not [string]::IsNullOrWhiteSpace($itemNumber)) { $itemRows += [PSCustomObject]@{ Row = $headerRow; ItemNumber = $itemNumber.Trim() } }
        }

        foreach ($item in $itemRows) {
            if (-not $validItems.ContainsKey($item.ItemNumber)) { continue }
            $key = "$($asset.assetNumber)/$($item.ItemNumber)"
            if (-not $photosByItem.ContainsKey($key)) { continue }
            foreach ($photo in $photosByItem[$key]) {
                $slot = [int]$photo.slotIndex
                $topRow = $item.Row + 4
                $bottomRow = $item.Row + 12
                $address = switch ($slot) {
                    0 { "C${topRow}:K${bottomRow}" }
                    1 { "N${topRow}:T${bottomRow}" }
                    2 { "W${topRow}:AC${bottomRow}" }
                    3 { "AF${topRow}:AL${bottomRow}" }
                    default { throw "Invalid photo slot for item $($item.ItemNumber)." }
                }
                $photoPath = Join-Path (Join-Path $SessionRoot 'photos') ($photo.fileKey + '.jpg')
                $target = $sheet.Range($address)
                $shape = $sheet.Shapes.AddPicture($photoPath, $false, $true, $target.Left, $target.Top, -1, -1)
                $shape.LockAspectRatio = -1
                $shape.Width = [Math]::Min($target.Width, 170.0787)
                if ($shape.Height -gt $target.Height) { $shape.Height = $target.Height }
                $shape.Left = $target.Left
                $shape.Top = $target.Top
                $shape.Placement = 1
                $shape.Name = "SM_$($asset.assetNumber)_$($item.ItemNumber)_$slot"
            }
        }

        # 写真がない項目も、写真帳様式の行構成・数式・表示を維持する。
    }

    $albumBook.Save()
    $albumBook.Close($true)
    $albumBook = $null
}
finally {
    if ($healthBook -ne $null) { $healthBook.Close($false) }
    if ($albumBook -ne $null) { $albumBook.Close($false) }
    if ($excel -ne $null) { $excel.Quit() }
    foreach ($item in @($targetNo11, $sourceNo11, $healthBook, $albumBook, $excel)) {
        if ($item -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
