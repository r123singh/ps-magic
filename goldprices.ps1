# Script to fetch 24 karat gold price in INR per gram
# Uses working free APIs and web scraping

param(
    [switch]$Quiet  # Suppress detailed output
)

# Set output encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

if (-not $Quiet) {
    Write-Host ""
    Write-Host "Fetching 24 Karat Gold Price (INR per gram)..." -ForegroundColor Cyan
    Write-Host ""
}

try {
    # Step 1: Get gold price in USD per troy ounce using web scraping
    $goldPricePerOunce = $null
    
    if (-not $Quiet) {
        Write-Host "Fetching gold price from web (goldprice.org)..." -ForegroundColor Yellow
    }
    
    # Method 1: Scrape from goldprice.org (reliable source)
    try {
        $webUrl = "https://www.goldprice.org/gold-price-per-ounce.html"
        $webRequest = Invoke-WebRequest -Uri $webUrl -UseBasicParsing -ErrorAction Stop -TimeoutSec 15
        
        # Look for price in the content - gold prices are typically in $XXXX format
        # Try multiple patterns to find the price
        $patterns = @(
            '\$[\d,]+\.?\d*',  # $2,000.00 format
            '[\d,]+\.?\d*\s*USD',  # 2000.00 USD format
            'price[:\s]+[\$]?[\d,]+\.?\d*',  # price: $2000 format
            '<span\s+class="gpoticker-price">\s*([\d,]+\.?\d*)\s*</span>'  # <span class="gpoticker-price">4,929.18</span>
        )
        
        foreach ($pattern in $patterns) {
            $patternMatches = [regex]::Matches($webRequest.Content, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            
            foreach ($match in $patternMatches) {
                $priceStr = $match.Value -replace '[\$,\sA-Za-z:]', ''
                $price = [double]$priceStr
                # Validate price is in reasonable range for gold per ounce (1500-5000 USD)
                if ($price -ge 1500 -and $price -le 5000) {
                    $goldPricePerOunce = $price
                    break
                }
            }
            
            if ($goldPricePerOunce) {
                break
            }
        }
        
        if ($goldPricePerOunce) {
            if (-not $Quiet) {
                Write-Host "  Success!" -ForegroundColor Green
            }
        }
    } catch {
        if (-not $Quiet) {
            Write-Host "  Web scraping failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
    
    # Method 2: Try alternative website if first fails
    if (-not $goldPricePerOunce -or $goldPricePerOunce -le 0) {
        try {
            if (-not $Quiet) {
                Write-Host "  Trying alternative source (goldpricez.com)..." -ForegroundColor DarkYellow
            }
            
            # Try goldpricez.com or another reliable source
            $webUrl2 = "https://www.goldpricez.com/us/ounce"
            $webRequest2 = Invoke-WebRequest -Uri $webUrl2 -UseBasicParsing -ErrorAction Stop -TimeoutSec 15
            $priceMatches = [regex]::Matches($webRequest2.Content, '\$[\d,]+\.?\d*') 
            foreach ($match in $priceMatches) {
                $priceStr = $match.Value -replace '[\$,]', ''
                $price = [double]$priceStr
                if ($price -ge 1500 -and $price -le 6000) {
                    $goldPricePerOunce = $price
                    break
                }
            }
        } catch {
            # Continue to fallback
        }
    }
    
    # Fallback: Use current market estimate
    if (-not $goldPricePerOunce -or $goldPricePerOunce -le 0) {
        if (-not $Quiet) {
            Write-Host "  Using market estimate (update manually for accuracy)" -ForegroundColor Yellow
            Write-Host "  Current estimate: ~$2200 USD/oz (as of 2025)" -ForegroundColor DarkYellow
        }
        # Approximate gold price - user should update this for accuracy
        $goldPricePerOunce = 2200
    }
    
    if (-not $Quiet) {
        $goldPriceFormatted = [math]::Round($goldPricePerOunce, 2)
        Write-Host "Gold price: $goldPriceFormatted USD per troy ounce" -ForegroundColor Green
    }
    
    Write-Host ""
    
    # Step 2: Get USD to INR exchange rate
    # Using exchangerate-api.com (confirmed working, no API key required)
    $usdToInrRate = $null
    
    try {
        if (-not $Quiet) {
            Write-Host "Fetching USD to INR exchange rate..." -ForegroundColor Yellow
        }
        
        # exchangerate-api.com open access endpoint (no API key required)
        $exchangeResponse = Invoke-RestMethod -Uri "https://api.exchangerate-api.com/v4/latest/USD" -Method Get -ErrorAction Stop -TimeoutSec 10
        
        if ($exchangeResponse -and $exchangeResponse.rates -and $exchangeResponse.rates.INR) {
            $usdToInrRate = [double]$exchangeResponse.rates.INR
            
            if (-not $Quiet) {
                Write-Host "  Success!" -ForegroundColor Green
            }
        } else {
            throw "Invalid response structure"
        }
    } catch {
        if (-not $Quiet) {
            Write-Host "  Failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
            Write-Host "  Using estimated rate: 83.0" -ForegroundColor Yellow
        }
        # Fallback to approximate rate
        $usdToInrRate = 83.0
    }
    
    if (-not $usdToInrRate -or $usdToInrRate -le 0) {
        throw "Could not fetch USD to INR exchange rate"
    }
    
    if (-not $Quiet) {
        $exchangeRateFormatted = [math]::Round($usdToInrRate, 2)
        Write-Host "Exchange rate: 1 USD = Rs. $exchangeRateFormatted" -ForegroundColor Green
    }
    
    Write-Host ""
    
    # Step 3: Convert to INR per gram for 24 karat gold
    # 1 troy ounce = 31.1035 grams
    # 24 karat = 100% pure gold
    $troyOunceToGram = 31.1035
    
    # Convert USD per troy ounce to INR per gram (International price)
    $internationalPricePerGram = ($goldPricePerOunce * $usdToInrRate) / $troyOunceToGram
    $internationalPricePerGram = [math]::Round($internationalPricePerGram, 2)
    
    # Step 4: Calculate Indian market price with taxes and premiums
    # Indian gold price includes:
    # 1. Import Duty: ~15% (basic customs + agriculture cess)
    # 2. GST: 3% (applied after import duty)
    # 3. Premiums: ~2-3% (dealer premium, logistics, hedging, etc.)
    
    $importDutyRate = 0.15  # 15%
    $gstRate = 0.03          # 3%
    $premiumRate = 0.02      # 2% (dealer premiums, logistics, etc.)
    
    # Calculate step by step
    $priceAfterImportDuty = $internationalPricePerGram * (1 + $importDutyRate)
    $priceAfterGST = $priceAfterImportDuty * (1 + $gstRate)
    $indianPricePerGram = $priceAfterGST * (1 + $premiumRate)
    
    # Round to 2 decimal places
    $indianPricePerGram = [math]::Round($indianPricePerGram, 2)
    
    # Calculate breakdown
    $importDutyAmount = [math]::Round($internationalPricePerGram * $importDutyRate, 2)
    $gstAmount = [math]::Round($priceAfterImportDuty * $gstRate, 2)
    $premiumAmount = [math]::Round($priceAfterGST * $premiumRate, 2)
    $totalAdditional = [math]::Round($indianPricePerGram - $internationalPricePerGram, 2)
    
    if (-not $Quiet) {
        Write-Host ""
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host "  24 KARAT GOLD PRICE (INR per gram)" -ForegroundColor Yellow
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  INTERNATIONAL PRICE:" -ForegroundColor White
        Write-Host "     Rs. $internationalPricePerGram per gram" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  INDIAN MARKET PRICE:" -ForegroundColor White
        Write-Host "     Rs. $indianPricePerGram per gram" -ForegroundColor Green
        Write-Host ""
        Write-Host "  PRICE BREAKDOWN:" -ForegroundColor White
        Write-Host "     Base (International):     Rs. $internationalPricePerGram" -ForegroundColor Gray
        Write-Host "     + Import Duty (15%):      Rs. $importDutyAmount" -ForegroundColor Gray
        Write-Host "     + GST (3%):               Rs. $gstAmount" -ForegroundColor Gray
        Write-Host "     + Premiums (~2%):         Rs. $premiumAmount" -ForegroundColor Gray
        Write-Host "     -------------------------------------------" -ForegroundColor Gray
        Write-Host "     Total Additional:        Rs. $totalAdditional" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  MARKET DETAILS:" -ForegroundColor White
        $goldPriceFormatted = [math]::Round($goldPricePerOunce, 2)
        $exchangeRateFormatted = [math]::Round($usdToInrRate, 2)
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "     - Gold Price: $goldPriceFormatted USD/oz" -ForegroundColor Gray
        Write-Host "     - Exchange Rate: 1 USD = Rs. $exchangeRateFormatted" -ForegroundColor Gray
        Write-Host "     - Purity: 24 Karat (100% Pure)" -ForegroundColor Gray
        Write-Host "     - Timestamp: $timestamp" -ForegroundColor Gray
        
        # Warn if using estimates
        if ($goldPricePerOunce -eq 2200) {
            Write-Host ""
            Write-Host "  NOTE: Gold price is estimated (~$2200 USD/oz)" -ForegroundColor Yellow
            Write-Host "        For accurate prices, update line 99 in the script" -ForegroundColor Yellow
            Write-Host "        or register for a free API key at:" -ForegroundColor Yellow
            Write-Host "        - goldapi.io (free tier available)" -ForegroundColor Gray
            Write-Host "        - metals-api.com (free tier available)" -ForegroundColor Gray
        }
        
        Write-Host ""
        Write-Host "===============================================" -ForegroundColor Cyan
        Write-Host ""
    } else {
        # Quiet mode: just output the Indian price
        Write-Output $indianPricePerGram
    }
    
    # Return the Indian price as output
    return $indianPricePerGram
    
} catch {
    if (-not $Quiet) {
        Write-Host ""
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        if ($_.Exception.Response) {
            try {
                $statusCode = $_.Exception.Response.StatusCode.value__
                Write-Host "   HTTP Status: $statusCode" -ForegroundColor Red
            } catch {
                # Ignore if status code can't be retrieved
            }
        }
        
        if ($_.Exception.InnerException) {
            Write-Host "   Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        }
        
        Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
        Write-Host "   - Check your internet connection" -ForegroundColor Gray
        Write-Host "   - Ensure you can access goldprice.org" -ForegroundColor Gray
        Write-Host "   - Try running the script again" -ForegroundColor Gray
        Write-Host ""
        Write-Host "For more accurate gold prices:" -ForegroundColor Yellow
        Write-Host "   - Register for free API key at goldapi.io" -ForegroundColor Gray
        Write-Host "   - Or manually update the estimate on line 95" -ForegroundColor Gray
        Write-Host ""
    } else {
        Write-Error $_.Exception.Message
    }
    
    exit 1
}
