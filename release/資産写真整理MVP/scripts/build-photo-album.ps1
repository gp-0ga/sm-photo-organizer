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

    $targetNo11.Cells.UnMerge()
    $targetNo11.Cells.Clear()
    $sourceNo11.UsedRange.Copy($targetNo11.Range('A1'))
    $sourceNo11.UsedRange.Copy()
    $targetNo11.Range('A1').PasteSpecial(8)
    for ($row = 1; $row -le $sourceNo11.UsedRange.Rows.Count; $row++) {
        $targetNo11.Rows.Item($row).RowHeight = $sourceNo11.Rows.Item($row).RowHeight
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

        foreach ($item in ($itemRows | Sort-Object Row -Descending)) {
            if (-not $validItems.ContainsKey($item.ItemNumber)) {
                $sheet.Rows("$($item.Row):$($item.Row + 2)").Delete()
                continue
            }
            $key = "$($asset.assetNumber)/$($item.ItemNumber)"
            if (-not $photosByItem.ContainsKey($key)) {
                $rows = $sheet.Rows("$($item.Row + 3):$($item.Row + 13)")
                $rows.Group()
                $rows.Hidden = $true
            }
        }
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
