# Test script for API endpoint: /api/projects/16/details
# Make sure the backend server is running on port 3000 before running this script

Write-Host "🧪 Testing API Endpoint: /api/projects/16/details" -ForegroundColor Cyan
Write-Host ""

# Step 1: Login to get JWT token
Write-Host "Step 1: Logging in..." -ForegroundColor Yellow
try {
    $loginBody = @{
        email = "sales1@flytxt.com"
    } | ConvertTo-Json

    $loginResponse = Invoke-RestMethod -Uri "http://localhost:3000/api/auth/login" `
        -Method Post `
        -ContentType "application/json" `
        -Body $loginBody `
        -ErrorAction Stop

    $TOKEN = $loginResponse.token
    Write-Host "✅ Login successful!" -ForegroundColor Green
    Write-Host "   User: $($loginResponse.user.email) ($($loginResponse.user.role))" -ForegroundColor Gray
    Write-Host "   Token: $($TOKEN.Substring(0, 30))..." -ForegroundColor Gray
    Write-Host ""
} catch {
    Write-Host "❌ Login failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Response) {
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $responseBody = $reader.ReadToEnd()
        Write-Host "   Response: $responseBody" -ForegroundColor Red
    }
    exit 1
}

# Step 2: Test the project details endpoint
Write-Host "Step 2: Testing /api/projects/16/details endpoint..." -ForegroundColor Yellow
try {
    $headers = @{
        Authorization = "Bearer $TOKEN"
    }

    $projectDetails = Invoke-RestMethod -Uri "http://localhost:3000/api/projects/16/details" `
        -Method Get `
        -Headers $headers `
        -ErrorAction Stop

    Write-Host "✅ Request successful!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Response:" -ForegroundColor Cyan
    $projectDetails | ConvertTo-Json -Depth 10 | Write-Host
} catch {
    Write-Host "❌ Request failed: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "   Status Code: $statusCode" -ForegroundColor Red
        
        $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
        $responseBody = $reader.ReadToEnd()
        Write-Host "   Response: $responseBody" -ForegroundColor Red
    }
    exit 1
}

Write-Host ""
Write-Host "✨ Test completed successfully!" -ForegroundColor Green
