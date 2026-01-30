# Script to monitor gold prices continuously
# Runs goldprices.ps1 in quiet mode and displays price every 1 minute

param(
    [int]$IntervalSeconds = 60,  # Update interval in seconds (default: 60 = 1 minute)
    [switch]$ShowTimestamp  # Show timestamp with each price update
)

# Get the script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$goldPriceScript = Join-Path $scriptDir "goldprices.ps1"

# Check if goldprices.ps1 exists
if (-not (Test-Path $goldPriceScript)) {
    Write-Host "Error: goldprices.ps1 not found at: $goldPriceScript" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Gold Price Monitor" -ForegroundColor Yellow
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Monitoring gold prices every $IntervalSeconds seconds..." -ForegroundColor Green
Write-Host "Press Ctrl+C to stop" -ForegroundColor Yellow
Write-Host ""

# Function to get and display price
function Get-GoldPrice {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    
    try {
        $price = & $goldPriceScript -Quiet
        
        if ($price) {
            if ($ShowTimestamp) {
                Write-Host "[$timestamp] " -NoNewline -ForegroundColor Gray
            }
            
            Write-Host "24K Gold Price: Rs. $price per gram" -ForegroundColor Green
        } else {
            Write-Host "[$timestamp] Error: Could not fetch price" -ForegroundColor Red
        }
    } catch {
        Write-Host "[$timestamp] Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Run continuously
try {
    while ($true) {
        Get-GoldPrice
        Start-Sleep -Seconds $IntervalSeconds
    }
} catch {
    # Handle Ctrl+C gracefully
    Write-Host ""
    Write-Host "Monitoring stopped." -ForegroundColor Yellow
    exit 0
}
