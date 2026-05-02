<#
.SYNOPSIS
    Intraday Trading Agent (Simulation Mode)
.DESCRIPTION
    Manages a corpus, tracks 5 stocks, updates Excel every minute.
    WARNING: Uses simulated prices. Integrate Broker API for live trading.
#>

# --- CONFIGURATION ---
$InitialCorpus = 500000
$StockList = @("RELIANCE", "TATASTEEL", "INFY", "HDFCBANK", "ICICIBANK")
$BasePrices = @{
    "RELIANCE" = 2400; "TATASTEEL" = 140; "INFY" = 1500; 
    "HDFCBANK" = 1600; "ICICIBANK" = 950
}
$ExcelPath = "$PSScriptRoot\Trading_Log.xlsx"
$UpdateInterval = 60 # Seconds

# --- STATE VARIABLES ---
$CurrentCash = $InitialCorpus
$Portfolio = @{}
foreach ($stock in $StockList) {
    $Portfolio[$stock] = @{ Qty = 0; AvgPrice = 0.0 }
}

# --- EXCEL SETUP ---
function Initialize-Excel {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        
        # Check if file exists, else create
        if (!(Test-Path $ExcelPath)) {
            $workbook = $excel.Workbooks.Add()
            $workbook.SaveAs($ExcelPath)
        } else {
            $workbook = $excel.Workbooks.Open($ExcelPath)
        }

        # Setup Dashboard Sheet
        $sheetDash = $workbook.Sheets.Item(1)
        $sheetDash.Name = "Dashboard"
        $sheetDash.Cells.Item(1,1) = "Stock"
        $sheetDash.Cells.Item(1,2) = "Qty"
        $sheetDash.Cells.Item(1,3) = "Avg Price"
        $sheetDash.Cells.Item(1,4) = "Current Price"
        $sheetDash.Cells.Item(1,5) = "Invested Val"
        $sheetDash.Cells.Item(1,6) = "Current Val"
        $sheetDash.Cells.Item(1,7) = "PnL"
        $sheetDash.Cells.Item(1,8) = "Corpus Balance"
        
        # Format Headers
        $sheetDash.Range("A1:H1").Font.Bold = $true
        $sheetDash.Range("A1:H1").Interior.Color = 65535 # Yellow

        # Setup Log Sheet
        if ($workbook.Sheets.Count -eq 1) {
            $sheetLog = $workbook.Sheets.Add()
            $sheetLog.Name = "Transaction_Log"
            $sheetLog.Cells.Item(1,1) = "Time"
            $sheetLog.Cells.Item(1,2) = "Stock"
            $sheetLog.Cells.Item(1,3) = "Action"
            $sheetLog.Cells.Item(1,4) = "Price"
            $sheetLog.Cells.Item(1,5) = "Corpus"
            $sheetLog.Range("A1:E1").Font.Bold = $true
        }
        $sheetLog = $workbook.Sheets.Item("Transaction_Log")

        return @{ Excel = $excel; Workbook = $workbook; Dash = $sheetDash; Log = $sheetLog }
    }
    catch {
        Write-Host "ERROR: Failed to initialize Excel. Ensure Excel is installed." -ForegroundColor Red
        exit
    }
}

# --- PRICE SIMULATION (Replace for Live API) ---
function Get-StockPrice {
    param($Symbol)
    # SIMULATION: Fluctuates base price by +/- 1% randomly
    $base = $BasePrices[$Symbol]
    $change = (Get-Random -Minimum -100 -Maximum 100) / 100
    $newPrice = $base + $change
    # Update base for next tick to simulate trend
    $BasePrices[$Symbol] = $newPrice 
    return [math]::Round($newPrice, 2)
}

# --- TRADING LOGIC ---
function Execute-TradingCycle {
    param($ExcelObj)
    $row = 2
    $totalPnL = 0
    $totalInvestedValue = 0

    foreach ($stock in $StockList) {
        $price = Get-StockPrice -Symbol $stock
        $qty = $Portfolio[$stock].Qty
        $avg = $Portfolio[$stock].AvgPrice
        
        $action = "HOLD"
        $tradeQty = 0

        # Simple Strategy: Buy if no position, Sell if 1% Profit/Loss
        if ($qty -eq 0 -and $CurrentCash -gt ($price * 100)) {
            $action = "BUY"
            $tradeQty = 100 # Fixed lot size for simulation
            $cost = $price * $tradeQty
            $CurrentCash -= $cost
            $Portfolio[$stock].Qty = $tradeQty
            $Portfolio[$stock].AvgPrice = $price
        }
        elseif ($qty -gt 0) {
            $pnlPercent = (($price - $avg) / $avg) * 100
            if ($pnlPercent -gt 1.0 -or $pnlPercent -lt -1.0) {
                $action = "SELL"
                $tradeQty = $qty
                $CurrentCash += ($price * $qty)
                $Portfolio[$stock].Qty = 0
                $Portfolio[$stock].AvgPrice = 0.0
            }
        }

        # Calculations
        $investedVal = $avg * $qty
        $currentVal = $price * $qty
        $pnl = $currentVal - $investedVal
        $totalPnL += $pnl
        $totalInvestedValue += $currentVal

        # Update Dashboard Excel
        $ExcelObj.Dash.Cells.Item($row, 1) = $stock
        $ExcelObj.Dash.Cells.Item($row, 2) = $qty
        $ExcelObj.Dash.Cells.Item($row, 3) = $avg
        $ExcelObj.Dash.Cells.Item($row, 4) = $price
        $ExcelObj.Dash.Cells.Item($row, 5) = [math]::Round($investedVal, 2)
        $ExcelObj.Dash.Cells.Item($row, 6) = [math]::Round($currentVal, 2)
        $ExcelObj.Dash.Cells.Item($row, 7) = [math]::Round($pnl, 2)
        $ExcelObj.Dash.Cells.Item($row, 8) = [math]::Round($CurrentCash + $totalInvestedValue, 2)
        
        # Color PnL
        if ($pnl -gt 0) { $ExcelObj.Dash.Cells.Item($row, 7).Font.Color = 65280 } # Green
        elseif ($pnl -lt 0) { $ExcelObj.Dash.Cells.Item($row, 7).Font.Color = 255 } # Red

        # Log Transaction if Action taken
        if ($action -ne "HOLD") {
            $logRow = $ExcelObj.Log.UsedRange.Rows.Count + 1
            $ExcelObj.Log.Cells.Item($logRow, 1) = (Get-Date -Format "HH:mm:ss")
            $ExcelObj.Log.Cells.Item($logRow, 2) = $stock
            $ExcelObj.Log.Cells.Item($logRow, 3) = $action
            $ExcelObj.Log.Cells.Item($logRow, 4) = $price
            $ExcelObj.Log.Cells.Item($logRow, 5) = [math]::Round($CurrentCash + $totalInvestedValue, 2)
        }

        $row++
    }
    
    # Update Global Corpus Cell (I1)
    $ExcelObj.Dash.Cells.Item(1, 9) = "Total Corpus"
    $ExcelObj.Dash.Cells.Item(2, 9) = [math]::Round($CurrentCash + $totalInvestedValue, 2)
    
    # Save Workbook
    try { $ExcelObj.Workbook.Save() } catch {}
}

# --- MAIN LOOP ---
Write-Host "----------------------------------------" -ForegroundColor Cyan
Write-Host "  TRADING AGENT STARTED (POWERShell)  " -ForegroundColor Cyan
Write-Host "  Corpus: ₹$InitialCorpus               " -ForegroundColor Cyan
Write-Host "  Mode: SIMULATION (Do not use live)  " -ForegroundColor Red
Write-Host "----------------------------------------" -ForegroundColor Cyan

$ExcelObj = Initialize-Excel
if ($null -eq $ExcelObj) { exit }

try {
    while ($true) {
        $now = Get-Date
        # Market Hours Check (9:15 AM to 3:30 PM IST)
        if ($now.Hour -ge 9 -and ($now.Hour -lt 15 -or ($now.Hour -eq 15 -and $now.Minute -le 30))) {
            Write-Host "[$($now.ToString('HH:mm:ss'))] Trading Cycle..." -ForegroundColor Green
            Execute-TradingCycle -ExcelObj $ExcelObj
        } else {
            Write-Host "[$($now.ToString('HH:mm:ss'))] Market Closed. Waiting..." -ForegroundColor Gray
        }
        Start-Sleep -Seconds $UpdateInterval
    }
}
catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
}
finally {
    Write-Host "Closing Excel..." -ForegroundColor Yellow
    $ExcelObj.Workbook.Close($true)
    $ExcelObj.Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelObj.Excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}