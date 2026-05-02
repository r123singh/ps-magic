# Script to generate speech using OpenAI TTS API
# Converts text to speech and saves as MP3 file

param(
    [Parameter(Mandatory=$false)]
    [string]$InputText = "Today is a wonderful day to build something people love!",
    
    [Parameter(Mandatory=$false)]
    [string]$Voice = "coral",
    
    [Parameter(Mandatory=$false)]
    [string]$Model = "gpt-4o-mini-tts",
    
    [Parameter(Mandatory=$false)]
    [string]$Instructions = "Speak in a cheerful and positive tone.",
    
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "speech.mp3",
    
    [Parameter(Mandatory=$false)]
    [switch]$PlayAudio
)

# OpenAI API Key - Replace with your actual API key
$ApiKey = "YOUR_OPENAI_API_KEY"

# Function to clear file lock if it exists
function Clear-FileLock {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    if (Test-Path $FilePath) {
        $maxRetries = 10
        $retryCount = 0
        
        while ($retryCount -lt $maxRetries) {
            try {
                # Try to open file for write access to check if it's locked
                $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
                $fileStream.Close()
                return $true
            } catch {
                $retryCount++
                if ($retryCount -lt $maxRetries) {
                    Start-Sleep -Milliseconds 200
                }
            }
        }
        return $false
    }
    return $true
}

# Function to play audio file inline
function Start-AudioPlayback {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    $absolutePath = (Resolve-Path $FilePath).Path
    
    try {
        # Try using Windows Media Player COM object (works inline)
        $mediaPlayer = New-Object -ComObject WMPlayer.OCX
        $mediaPlayer.URL = $absolutePath
        $mediaPlayer.controls.play()
        
        # Wait for playback to complete
        while ($mediaPlayer.playState -eq 3) {  # 3 = playing
            Start-Sleep -Milliseconds 100
        }
        
        # Cleanup
        $mediaPlayer.close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($mediaPlayer) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Small delay to ensure file is released
        Start-Sleep -Milliseconds 300
        
    } catch {
        # Fallback to default audio player
        Write-Host "Note: Using default audio player..." -ForegroundColor Gray
        Start-Process -FilePath $absolutePath -WindowStyle Hidden
        Start-Sleep -Seconds 2  # Give it a moment to start
    }
}

# API endpoint
$uri = "https://api.openai.com/v1/audio/speech"

# Prepare request body
$body = @{
    model = $Model
    input = $InputText
    voice = $Voice
    instructions = $Instructions
} | ConvertTo-Json

# Prepare headers
$headers = @{
    "Authorization" = "Bearer $ApiKey"
    "Content-Type" = "application/json"
}

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  OpenAI Text-to-Speech" -ForegroundColor Yellow
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Model: $Model" -ForegroundColor Green
Write-Host "Voice: $Voice" -ForegroundColor Green
Write-Host "Input: $InputText" -ForegroundColor Gray
Write-Host "Output: $OutputFile" -ForegroundColor Green
Write-Host ""

try {
    Write-Host "Generating speech..." -ForegroundColor Yellow
    
    # Clear file lock if file exists and is locked
    if (-not (Clear-FileLock -FilePath $OutputFile)) {
        Write-Host "Warning: File is locked. Attempting to wait for release..." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    
    # Make the API request
    $response = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -Body $body -ContentType "application/json" -UseBasicParsing
    
    # Save the audio file (use temp file first to avoid locking issues)
    $tempFile = "$OutputFile.tmp"
    [System.IO.File]::WriteAllBytes($tempFile, $response.Content)
    
    # Replace original file atomically
    if (Test-Path $OutputFile) {
        Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue
    }
    Move-Item $tempFile $OutputFile -Force
    
    Write-Host "Success! Speech saved to: $OutputFile" -ForegroundColor Green
    
    # Get file size
    $fileInfo = Get-Item $OutputFile
    $fileSizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
    Write-Host "File size: $fileSizeKB KB" -ForegroundColor Gray
    
    # Play audio if requested
    if ($PlayAudio) {
        Write-Host ""
        Write-Host "Playing audio..." -ForegroundColor Yellow
        Start-AudioPlayback -FilePath $OutputFile
    }
    
} catch {
    Write-Host "Error: Failed to generate speech" -ForegroundColor Red
    
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "HTTP Status Code: $statusCode" -ForegroundColor Red
        
        # Try to get error details
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            $errorDetails = $responseBody | ConvertFrom-Json
            Write-Host "Error Message: $($errorDetails.error.message)" -ForegroundColor Red
        } catch {
            Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    exit 1
}

# Voice options alloy
# ash
# ballad
# coral
# echo
# fable
# nova
# onyx
# sage
# shimmer
# verse
# marin
# cedar