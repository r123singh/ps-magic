# Read plain-text API key (or use: Read-Host -AsSecureString | ConvertFrom-SecureString | Out-File "elevenlabs.key" for encrypted storage)
$apiKey = (Get-Content "elevenlabs.key" -Raw).Trim()

$headers = @{
    "xi-api-key" = $apiKey
    "Content-Type" = "application/json"
}
$response = Invoke-RestMethod -Uri "https://api.elevenlabs.io/v1/voices" -Headers $headers
$response | ConvertTo-Json | Out-File "elevenlabs-voices.json"