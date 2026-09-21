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
$stage = 'Excelを起動しています'
$rowTopCache = @{}

function Get-ExactRowTop {
    param(
        $Sheet,
        [int]$Row
    )

    $key = "$($Sheet.Name)|$Row"
    if ($script:rowTopCache.ContainsKey($key)) { return [single]$script:rowTopCache[$key] }

    # Range.Topは行が下がるほど丸め誤差が累積する様式があるため、
    # 実際の各行高を合計して結合セル上端を求める。
    $top = 0.0
    for ($rowIndex = 1; $rowIndex -lt $Row; $rowIndex++) {
        $height = @($Sheet.Rows.Item($rowIndex).RowHeight)[0]
        $top += [double]$height
    }
    $script:rowTopCache[$key] = [single]$top
    return [single]$top
}

function Add-EmbeddedPicture {
    param(
        $Sheet,
        [string]$PhotoPath,
        $TargetRange
    )
    $anchorCell = $TargetRange.Cells.Item(1, 1)
    $left = [single](@($anchorCell.Left)[0])
    $top = Get-ExactRowTop -Sheet $Sheet -Row ([int]$anchorCell.Row)
    # ExcelのAddPictureはファイル名をString、座標とサイズをSingleで受け取る。
    return $Sheet.Shapes.AddPicture(
        [string]$PhotoPath,
        [int]0,
        [int]-1,
        $left,
        $top,
        [single]-1,
        [single]-1
    )
}

function Get-MergedFrame {
    param(
        $Sheet,
        [string]$AnchorAddress,
        $FallbackRange
    )

    $anchorCell = $Sheet.Range($AnchorAddress)
    if ([bool]$anchorCell.MergeCells) {
        Write-Output -NoEnumerate $anchorCell.MergeArea
        return
    }
    Write-Output -NoEnumerate $FallbackRange
}

try {
    $stage = 'Excelを起動しています'
    try { $excel = New-Object -ComObject Excel.Application }
    catch { throw 'Microsoft Excel is not installed or could not be started.' }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    $stage = '写真帳コピーを開いています'
    $albumBook = $excel.Workbooks.Open($outputPath, 0, $false)

    # No11の構成を先に全件検査し、点検結果1・2の値だけを同じセル位置へ転記する。
    # 書式・入力規則・数式・ほかのセルには触れない。
    if (Test-Path -LiteralPath $healthPath -PathType Leaf) {
    $stage = '健全度判定表を開いています'
    $healthBook = $excel.Workbooks.Open($healthPath, 0, $true)
    $sourceNo11 = $healthBook.Worksheets.Item('No11')
    $targetNo11 = $albumBook.Worksheets.Item('No11')
    if ($targetNo11.ProtectContents) { throw 'The No11 sheet in the photo album is protected.' }

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

    $sourceHeaderAbsoluteRow = [int]($sourceFirstRow + $sourceHeaderRow - 1)
    $targetHeaderAbsoluteRow = [int]($targetFirstRow + $targetHeaderRow - 1)
    if ($sourceHeaderAbsoluteRow -ne $targetHeaderAbsoluteRow) {
        throw '健全度判定表と写真帳で、No11の見出し行が一致しません。写真帳は変更していません。'
    }
    foreach ($requiredHeader in @('資産番号', '項目番号', '点検結果1', '点検結果2')) {
        if ([int]$sourceHeaders[$requiredHeader] -ne [int]$targetHeaders[$requiredHeader]) {
            throw "健全度判定表と写真帳で、No11の「$requiredHeader」の列位置が一致しません。写真帳は変更していません。"
        }
    }

    $sourceAssetIndex = [int]$sourceHeaders['資産番号'] - $sourceFirstColumn + 1
    $sourceItemIndex = [int]$sourceHeaders['項目番号'] - $sourceFirstColumn + 1
    $targetAssetIndex = [int]$targetHeaders['資産番号'] - $targetFirstColumn + 1
    $targetItemIndex = [int]$targetHeaders['項目番号'] - $targetFirstColumn + 1
    $sourceLastRow = [int]($sourceFirstRow + $sourceRowCount - 1)
    $targetLastRow = [int]($targetFirstRow + $targetRowCount - 1)
    if ($sourceLastRow -ne $targetLastRow) {
        throw '健全度判定表と写真帳で、No11の最終行が一致しません。写真帳は変更していません。'
    }

    $stage = 'No11の構成と転記先を確認しています'
    for ($row = $sourceHeaderRow + 1; $row -le $sourceRowCount; $row++) {
        $absoluteRow = [int]($sourceFirstRow + $row - 1)
        $sourceAssetNumber = ([string]$sourceValues[$row, $sourceAssetIndex]).Trim()
        $sourceItemNumber = ([string]$sourceValues[$row, $sourceItemIndex]).Trim()
        $targetRowIndex = [int]($absoluteRow - $targetFirstRow + 1)
        $targetAssetNumber = ([string]$targetValues[$targetRowIndex, $targetAssetIndex]).Trim()
        $targetItemNumber = ([string]$targetValues[$targetRowIndex, $targetItemIndex]).Trim()
        if ($sourceAssetNumber -ne $targetAssetNumber -or $sourceItemNumber -ne $targetItemNumber) {
            throw "健全度判定表と写真帳で、No11の$($absoluteRow)行目の資産番号または項目番号が一致しません。写真帳は変更していません。"
        }
        foreach ($resultHeader in @('点検結果1', '点検結果2')) {
            $targetCell = $targetNo11.Cells.Item($absoluteRow, [int]$targetHeaders[$resultHeader])
            if ([bool]$targetCell.HasFormula) {
                throw "写真帳No11の$($targetCell.Address($false, $false))に数式があります。写真帳は変更していません。"
            }
        }
    }

    $stage = 'No11の点検結果1・2を転記しています'
    for ($row = $sourceHeaderRow + 1; $row -le $sourceRowCount; $row++) {
        $absoluteRow = [int]($sourceFirstRow + $row - 1)
        foreach ($resultHeader in @('点検結果1', '点検結果2')) {
            $column = [int]$sourceHeaders[$resultHeader]
            $sourceCell = $sourceNo11.Cells.Item($absoluteRow, $column)
            $targetCell = $targetNo11.Cells.Item($absoluteRow, $column)
            $targetCell.Value2 = $sourceCell.Value2
        }
    }
    $healthBook.Close($false)
    $healthBook = $null
    }

    $sheetByAsset = @{}
    $stage = '写真帳の資産シートを確認しています'
    foreach ($sheet in $albumBook.Worksheets) {
        if ($sheet.Name -match '^\d+_\d+$') {
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
        $stage = "資産 $($asset.assetNumber) の写真貼り付け先を確認しています"
        if (-not $sheetByAsset.ContainsKey([string]$asset.assetNumber)) { throw "No photo sheet found for asset $($asset.assetNumber)." }
        $sheet = $sheetByAsset[[string]$asset.assetNumber]
        $validItems = @{}
        foreach ($itemNumber in $asset.itemNumbers) { $validItems[[string]$itemNumber] = $true }

        $fullKey = "$($asset.assetNumber)/full"
        if ($photosByItem.ContainsKey($fullKey)) {
            $photo = $photosByItem[$fullKey][0]
            $stage = "資産 $($asset.assetNumber) の全景写真「$($photo.sourceName)」を貼り付けています"
            $photoPath = Join-Path (Join-Path $SessionRoot 'photos') ($photo.fileKey + '.jpg')
            $fallback = $sheet.Range('E5:S18')
            $target = Get-MergedFrame -Sheet $sheet -AnchorAddress 'E5' -FallbackRange $fallback
            $shape = Add-EmbeddedPicture -Sheet $sheet -PhotoPath $photoPath -TargetRange $target
            $shape.LockAspectRatio = -1
            $shape.Width = [single](@($target.Width)[0])
            $targetHeight = [single](@($target.Height)[0])
            if ($shape.Height -gt $targetHeight) { $shape.Height = $targetHeight }
            $anchorCell = $target.Cells.Item(1, 1)
            $shape.Left = [single](@($anchorCell.Left)[0])
            $shape.Top = Get-ExactRowTop -Sheet $sheet -Row ([int]$anchorCell.Row)
            $shape.Placement = 1
            $shape.Name = "SM_$($asset.assetNumber)_full"
        }

        # 項目は14行間隔と決め打ちせず、写真帳のAP列にある実際の項目番号の行を探す。
        # 項目ごとの高さが異なっても、次の項目の直前までを写真枠として扱う。
        $itemRows = @()
        $sheetFirstRow = [int]$sheet.UsedRange.Row
        $sheetLastRow = $sheetFirstRow + [int]$sheet.UsedRange.Rows.Count - 1
        for ($row = $sheetFirstRow; $row -le $sheetLastRow; $row++) {
            $itemNumber = ([string]$sheet.Range("AP$row").Value2).Trim()
            if ($itemNumber -and $validItems.ContainsKey($itemNumber)) {
                $itemRows += [PSCustomObject]@{ Row = $row; ItemNumber = $itemNumber }
            }
        }

        for ($itemIndex = 0; $itemIndex -lt $itemRows.Count; $itemIndex++) {
            $item = $itemRows[$itemIndex]
            $key = "$($asset.assetNumber)/$($item.ItemNumber)"
            if (-not $photosByItem.ContainsKey($key)) { continue }
            foreach ($photo in $photosByItem[$key]) {
                $stage = "資産 $($asset.assetNumber)・項目 $($item.ItemNumber) の写真「$($photo.sourceName)」を貼り付けています"
                $slot = [int]$photo.slotIndex
                $topRow = $item.Row + 4
                $nextHeaderRow = if ($itemIndex -lt $itemRows.Count - 1) { $itemRows[$itemIndex + 1].Row } else { $item.Row + 14 }
                $bottomRow = [Math]::Max($topRow, $nextHeaderRow - 2)
                $frame = switch ($slot) {
                    0 { @{ Column = 'C'; Address = "C${topRow}:K${bottomRow}" } }
                    1 { @{ Column = 'N'; Address = "N${topRow}:T${bottomRow}" } }
                    2 { @{ Column = 'W'; Address = "W${topRow}:AC${bottomRow}" } }
                    3 { @{ Column = 'AF'; Address = "AF${topRow}:AL${bottomRow}" } }
                    default { throw "Invalid photo slot for item $($item.ItemNumber)." }
                }
                $photoPath = Join-Path (Join-Path $SessionRoot 'photos') ($photo.fileKey + '.jpg')
                $fallback = $sheet.Range($frame.Address)
                $target = Get-MergedFrame -Sheet $sheet -AnchorAddress "$($frame.Column)$topRow" -FallbackRange $fallback
                $shape = Add-EmbeddedPicture -Sheet $sheet -PhotoPath $photoPath -TargetRange $target
                $shape.LockAspectRatio = -1
                $shape.Width = [single](@($target.Width)[0])
                $targetHeight = [single](@($target.Height)[0])
                if ($shape.Height -gt $targetHeight) { $shape.Height = $targetHeight }
                $anchorCell = $target.Cells.Item(1, 1)
                $shape.Left = [single](@($anchorCell.Left)[0])
                $shape.Top = Get-ExactRowTop -Sheet $sheet -Row ([int]$anchorCell.Row)
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
catch {
    throw "処理段階：$stage`n$($_.Exception.Message)"
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
