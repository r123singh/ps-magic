<#
.SYNOPSIS
    ElevenLabs API - Full-featured PowerShell script for voice generation, transcription, and account management.

.DESCRIPTION
    Explores full ElevenLabs capabilities: TTS, speech-to-text, voices, models, voice settings, subscription, and history.

.PARAMETER Action
    Action to perform: voices, models, voice, settings, subscription, convert, transcribe, history

.PARAMETER VoiceId
    Voice ID (e.g., EXAVITQu4vr4xnSDxMaL). Use 'voices' action to list.

.PARAMETER VoiceName
    Voice name (partial match, e.g., "Sarah", "Roger"). Resolves to voice_id from the voices list.

.PARAMETER VoiceList
    Comma-separated voice names or IDs for convert-from-list. E.g., "Sarah,Roger,Laura"

.PARAMETER Text
    Text to convert to speech (for convert action).

.PARAMETER ModelId
    TTS model: eleven_multilingual_v2 (default), eleven_turbo_v2_5, eleven_flash_v2_5, eleven_v2

.PARAMETER OutputFormat
    Audio format: mp3_44100_128 (default), mp3_22050_32, pcm_16000, pcm_22050, pcm_24000, pcm_44100

.EXAMPLE
    .\elevenlabs-tts.ps1 -Action voices -SaveVoicesToJson
    .\elevenlabs-tts.ps1 -Action models
    .\elevenlabs-tts.ps1 -Action subscription
    .\elevenlabs-tts.ps1 -Action convert -VoiceId EXAVITQu4vr4xnSDxMaL -Text "Hello world" -OutputFile out.mp3 -PlayAudio
    .\elevenlabs-tts.ps1 -Action convert -Text "Custom speech" -Stability 0.7 -SimilarityBoost 0.8
    .\elevenlabs-tts.ps1 -Action transcribe -InputFile audio.mp3
    .\elevenlabs-tts.ps1 -Action voice -VoiceId EXAVITQu4vr4xnSDxMaL
    .\elevenlabs-tts.ps1 -Action settings -VoiceId EXAVITQu4vr4xnSDxMaL
    .\elevenlabs-tts.ps1 -Action history
#>

param(
    [Parameter(Mandatory = $false)]
    [ValidateSet("voices", "models", "voice", "settings", "subscription", "convert", "convert-from-list", "transcribe", "history")]
    [string]$Action = "voices",

    [Parameter(Mandatory = $false)]
    [string]$VoiceId = "EXAVITQu4vr4xnSDxMaL",

    [Parameter(Mandatory = $false)]
    [string]$VoiceName = "",

    [Parameter(Mandatory = $false)]
    [string[]]$VoiceList = @(),

    [Parameter(Mandatory = $false)]
    [string]$Text = "Today is a wonderful day to build something people love!",

    [Parameter(Mandatory = $false)]
    [string]$ModelId = "eleven_multilingual_v2",

    [Parameter(Mandatory = $false)]
    [ValidateSet("mp3_44100_128", "mp3_22050_32", "mp3_44100_192", "pcm_16000", "pcm_22050", "pcm_24000", "pcm_44100", "ulaw_8000")]
    [string]$OutputFormat = "mp3_44100_128",

    [Parameter(Mandatory = $false)]
    [double]$Stability = 0.5,

    [Parameter(Mandatory = $false)]
    [double]$SimilarityBoost = 0.75,

    [Parameter(Mandatory = $false)]
    [double]$Style = 0,

    [Parameter(Mandatory = $false)]
    [bool]$UseSpeakerBoost = $true,

    [Parameter(Mandatory = $false)]
    [double]$Speed = 1.0,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile = "elevenlabs-output.mp3",

    [Parameter(Mandatory = $false)]
    [string]$InputFile = "",

    [Parameter(Mandatory = $false)]
    [string]$TranscribeModel = "scribe_v2",

    [Parameter(Mandatory = $false)]
    [string]$LanguageCode = "",

    [Parameter(Mandatory = $false)]
    [string]$VoicesJson = "elevenlabs-voices.json",

    [Parameter(Mandatory = $false)]
    [switch]$PlayAudio,

    [Parameter(Mandatory = $false)]
    [switch]$SaveVoicesToJson,

    [Parameter(Mandatory = $false)]
    [int]$HistoryLimit = 100
)

# --- Config ---
$script:BaseUrl = "https://api.elevenlabs.io/v1"
$apiKey = (Get-Content "elevenlabs.key" -Raw -ErrorAction Stop).Trim()

$script:Headers = @{
    "xi-api-key"   = $apiKey
    "Content-Type" = "application/json"
}

# --- Helpers ---
function Invoke-ElevenLabsApi {
    param([string]$Method = "GET", [string]$Uri, [object]$Body = $null, [string]$ContentType = "application/json")
    $params = @{
        Uri     = $Uri
        Method  = $Method
        Headers = $script:Headers
    }
    if ($Body) {
        $params["Body"] = if ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 10 }
        $params["ContentType"] = $ContentType
    }
    Invoke-RestMethod @params
}

function Resolve-VoiceId {
    param([string]$NameOrId, [object[]]$Voices)
    if (-not $NameOrId) { return $null }
    # If it looks like a voice ID (alphanumeric, ~20 chars), use as-is
    if ($NameOrId -match '^[A-Za-z0-9]{18,}$') { return $NameOrId }
    # Look up by voice_id or name (partial, case-insensitive)
    $v = $Voices | Where-Object { $_.voice_id -eq $NameOrId } | Select-Object -First 1
    if ($v) { return $v.voice_id }
    $v = $Voices | Where-Object { $_.name -like "*$NameOrId*" } | Select-Object -First 1
    if ($v) { return $v.voice_id }
    Write-Host "No voice found matching '$NameOrId'" -ForegroundColor Red
    return $null
}

function Start-AudioPlayback {
    param([string]$FilePath)
    $abs = (Resolve-Path $FilePath).Path
    try {
        $wm = New-Object -ComObject WMPlayer.OCX
        $wm.URL = $abs
        $wm.controls.play()
        while ($wm.playState -eq 3) { Start-Sleep -Milliseconds 100 }
        $wm.close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wm) | Out-Null
    } catch {
        Start-Process -FilePath $abs
    }
}

# --- Actions ---
function Get-Voices {
    $r = Invoke-ElevenLabsApi -Uri "$script:BaseUrl/voices"
    $voices = if ($r.voices) { $r.voices } else { $r }
    Write-Host "`nVoices ($(($voices | Measure-Object).Count)):" -ForegroundColor Cyan
    $voices | ForEach-Object {
        $vid = $_.voice_id
        $name = $_.name
        $desc = if ($_.description) { $_.description.Substring(0, [Math]::Min(60, $_.description.Length)) + "..." } else { "" }
        Write-Host "  $vid | $name" -ForegroundColor White
        if ($desc) { Write-Host "    $desc" -ForegroundColor Gray }
    }
    if ($SaveVoicesToJson) {
        $voices | ConvertTo-Json -Depth 6 | Out-File $VoicesJson -Encoding utf8
        Write-Host "`nSaved to $VoicesJson" -ForegroundColor Green
    }
    return $voices
}

function Get-Models {
    $r = Invoke-ElevenLabsApi -Uri "$script:BaseUrl/models"
    $models = if ($r -is [array]) { $r } else { $r.models }
    $tts = $models | Where-Object { $_.can_do_text_to_speech -eq $true }
    Write-Host "`nTTS Models:" -ForegroundColor Cyan
    $tts | ForEach-Object {
        Write-Host "  $($_.model_id) | $($_.name)" -ForegroundColor White
        Write-Host "    Languages: $($_.languages -join ', ')" -ForegroundColor Gray
    }
    return $tts
}

function Get-Voice {
    param([string]$Id)
    if (-not $Id) { Write-Host "VoiceId required." -ForegroundColor Red; return }
    $r = Invoke-ElevenLabsApi -Uri "$script:BaseUrl/voices/$Id"
    $r | ConvertTo-Json -Depth 6
}

function Get-VoiceSettings {
    param([string]$Id)
    if (-not $Id) { Write-Host "VoiceId required." -ForegroundColor Red; return }
    $r = Invoke-ElevenLabsApi -Uri "$script:BaseUrl/voices/$Id/settings"
    Write-Host "`nVoice Settings for $Id" -ForegroundColor Cyan
    $r | Format-List
    return $r
}

function Get-Subscription {
    $r = Invoke-ElevenLabsApi -Uri "$script:BaseUrl/user/subscription"
    Write-Host "`nSubscription:" -ForegroundColor Cyan
    Write-Host "  Tier: $($r.tier)" -ForegroundColor White
    Write-Host "  Status: $($r.status)" -ForegroundColor White
    if ($r.character_count) {
        Write-Host "  Characters used: $($r.character_count) / $($r.character_limit)" -ForegroundColor Yellow
    }
    $r | Format-List
    return $r
}

function Convert-ToSpeech {
    param([string]$VId, [string]$Txt, [string]$MId, [string]$Fmt, [double]$Stab, [double]$Sim, [double]$Sty, [bool]$Spk, [double]$Spd, [string]$OutFile, [bool]$Play)
    if (-not $VId) { Write-Host "VoiceId required for convert. Run with -Action voices to list." -ForegroundColor Red; return }
    $uri = "$script:BaseUrl/text-to-speech/$VId"
    $uri += "?output_format=$Fmt"
    $body = @{
        text  = $Txt
        model_id = $MId
        voice_settings = @{
            stability         = $Stab
            similarity_boost  = $Sim
            style             = $Sty
            use_speaker_boost = $Spk
        }
    }
    if ($Spd -ne 1.0) { $body.voice_settings["speed"] = $Spd }
    Write-Host "`nConverting to speech..." -ForegroundColor Yellow
    Write-Host "  Voice: $VId | Model: $MId | Format: $Fmt" -ForegroundColor Gray
    try {
        $response = Invoke-WebRequest -Uri $uri -Method Post -Headers $script:Headers -Body ($body | ConvertTo-Json) -ContentType "application/json" -UseBasicParsing
        $outPath = if ([System.IO.Path]::IsPathRooted($OutFile)) { $OutFile } else { Join-Path (Get-Location) $OutFile }
        [System.IO.File]::WriteAllBytes($outPath, $response.Content)
        Write-Host "Saved: $OutFile" -ForegroundColor Green
        if ($Play) { Start-AudioPlayback -FilePath $OutFile }
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.Response) {
            $sr = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            Write-Host $sr.ReadToEnd() -ForegroundColor Red
        }
    }
}

function Convert-Transcribe {
    param([string]$InFile, [string]$Model, [string]$Lang)
    if (-not (Test-Path $InFile)) { Write-Host "Input file not found: $InFile" -ForegroundColor Red; return }
    $uri = "$script:BaseUrl/speech-to-text"
    $absPath = (Resolve-Path $InFile).Path
    try {
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $form = @{ file = Get-Item $absPath; model_id = $Model }
            if ($Lang) { $form["language_code"] = $Lang }
            $resp = Invoke-RestMethod -Uri $uri -Method Post -Headers @{ "xi-api-key" = $apiKey } -Form $form
        } else {
            $json = curl.exe -sS -X POST $uri -H "xi-api-key: $apiKey" -F "file=@$absPath" -F "model_id=$Model"
            if ($LASTEXITCODE -ne 0) { throw "curl failed" }
            $resp = $json | ConvertFrom-Json
        }
        Write-Host "`nTranscript:" -ForegroundColor Cyan
        Write-Host $resp.text -ForegroundColor White
        if ($resp.words) { Write-Host "`nWord timestamps available." -ForegroundColor Gray }
        return $resp
    } catch {
        if ($_.ErrorDetails.Message) {
            try { $err = $_.ErrorDetails.Message | ConvertFrom-Json; Write-Host "Error: $($err.detail.message)" -ForegroundColor Red }
            catch { Write-Host "Error: $($_.ErrorDetails.Message)" -ForegroundColor Red }
        } else { Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red }
    }
}

function Get-History {
    param([int]$Limit)
    $uri = "$script:BaseUrl/history?page_size=$Limit"
    $r = Invoke-ElevenLabsApi -Uri $uri
    $items = if ($r.history) { $r.history } else { @($r) }
    Write-Host "`nHistory (last $Limit):" -ForegroundColor Cyan
    $items | ForEach-Object {
        $txt = if ($_.text) { $_.text.Substring(0, [Math]::Min(50, $_.text.Length)) + "..." } else { "-" }
        $line = "  $($_.history_item_id) | $($_.source) | $txt"
        Write-Host $line -ForegroundColor White
    }
    return $items
}

# --- Main ---
try {
    switch ($Action) {
        "voices"      { Get-Voices | Out-Null }
        "models"      { Get-Models | Out-Null }
        "voice"       { Get-Voice -Id $VoiceId }
        "settings"    { Get-VoiceSettings -Id $VoiceId }
        "subscription" { Get-Subscription | Out-Null }
        "convert"     {
            $voices = (Invoke-ElevenLabsApi -Uri "$script:BaseUrl/voices").voices
            $vId = if ($VoiceName) { Resolve-VoiceId -NameOrId $VoiceName -Voices $voices } else { $VoiceId }
            if (-not $vId -and $VoiceName) { Write-Host "Voice not found: $VoiceName. Run -Action voices to list." -ForegroundColor Red; exit 1 }
            Convert-ToSpeech -VId $vId -Txt $Text -MId $ModelId -Fmt $OutputFormat -Stab $Stability -Sim $SimilarityBoost -Sty $Style -Spk $UseSpeakerBoost -Spd $Speed -OutFile $OutputFile -Play $PlayAudio.IsPresent
        }
        "convert-from-list" {
            $r = Invoke-ElevenLabsApi -Uri "$script:BaseUrl/voices"
            $voices = if ($r.voices) { $r.voices } else { $r }
            $list = @()
            if ($VoiceList.Count -gt 0) { $list = $VoiceList | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } }
            if ($VoiceName) { $list = @($VoiceName) + $list }
            if ($list.Count -eq 0) { Write-Host "Provide -VoiceList Sarah,Roger,Laura or -VoiceName Sarah" -ForegroundColor Red; exit 1 }
            foreach ($item in $list) {
                $vId = Resolve-VoiceId -NameOrId $item.Trim() -Voices $voices
                if (-not $vId) { Write-Host "Skipping unknown voice: $item" -ForegroundColor Yellow; continue }
                $v = $voices | Where-Object { $_.voice_id -eq $vId } | Select-Object -First 1
                $safeName = ($v.name -replace '[^\w\-]', '_').Substring(0, [Math]::Min(40, ($v.name -replace '[^\w\-]', '_').Length))
                $out = "speech-$safeName.mp3"
                Convert-ToSpeech -VId $vId -Txt $Text -MId $ModelId -Fmt $OutputFormat -Stab $Stability -Sim $SimilarityBoost -Sty $Style -Spk $UseSpeakerBoost -Spd $Speed -OutFile $out -Play $false
            }
        }
        "transcribe"   {
            $f = if ($InputFile) { $InputFile } else { $OutputFile }
            Convert-Transcribe -InFile $f -Model $TranscribeModel -Lang $LanguageCode
        }
        "history"     { Get-History -Limit $HistoryLimit | Out-Null }
    }
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
