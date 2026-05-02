# PowerShell Magic for different use cases

A collection of PowerShell scripts for various development and testing scenarios. These scripts showcase how terminal scripts in Windows PowerShell can simplify common tasks without needing to open browsers, prepare Postman collections, or set up complex testing environments.

## Use case: Test API Endpoint

The `test-endpoint.ps1` script demonstrates how to test API endpoints directly from PowerShell. This is particularly useful for backend development where you need to quickly test endpoints without opening a browser or configuring Postman collections.

## Use case: Gold Price Monitoring in PowerShell

The `goldprices.ps1` script demonstrates how to fetch gold prices in India from the international market. This is particularly useful for gold dealers and investors who need to know the price of gold in India.

## Use case: Trading Agent in PowerShell

The `tradingagent.ps1` script demonstrates how to create a trading agent in PowerShell. This is particularly useful for traders who need to automate their trading strategies.

## Use case: ElevenLabs TTS and ElevenLabs API for text to speech and speech to text.

The `elevenlabs-tts.ps1` script demonstrates how to use the ElevenLabs API for text to speech and speech to text. This is particularly useful for developers who need to convert text to speech and speech to text.

## Use case: OpenAI TTS and OpenAI API for text to speech and speech to text.

The `openai-tts.ps1` script demonstrates how to use the OpenAI API for text to speech and speech to text. This is particularly useful for developers who need to convert text to speech and speech to text.


### Features:

- **Error Handling**: Comprehensive error handling with detailed error messages
- **Color-coded Output**: Visual feedback with colored console output for better readability
- **No Dependencies**: Pure PowerShell - no additional tools or setup required
- **Quick Testing**: Test endpoints instantly without frontend or external tools

### Prerequisites:

- Windows PowerShell (5.1 or later) or PowerShell Core
- Backend server running on `http://localhost:3000`
- Valid API credentials (email: `sales1@company.com`)

### Usage:

1. **Start your backend server** on port 3000:
   ```powershell
   # Make sure your backend is running on http://localhost:3000
   ```

2. **Run the test script**:
   ```powershell
   .\test-endpoint.ps1
   ```

### What the Script Does:

The script performs the following steps:

1. **Login**: Authenticates with the API using the configured email address and retrieves a JWT token
2. **Test Endpoint**: Makes an authenticated GET request to your endpoint
3. **Display Results**: Shows the response in formatted JSON

### Example Output

```
🧪 Testing API Endpoint: <your-endpoint>

Step 1: Logging in...
✅ Login successful!
   User: sales1@company.com (role)
   Token: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...

Step 2: Testing <your-endpoint> endpoint...
✅ Request successful!

Response:
{
  "id": 16,
  "name": "Project Name",
  ...
}

✨ Test completed successfully!
```

### Customization

To test different endpoints, modify the script:
- Change the endpoint URL on line 42
- Update the email address on line 11
- Adjust the server URL if your backend runs on a different port

### Benefits

- **No Browser Required**: Test APIs directly from the terminal
- **No Postman Setup**: Skip creating and maintaining Postman collections
- **No Frontend Needed**: Test backend endpoints independently
- **Fast Iteration**: Quick feedback loop for API development
- **CI/CD Friendly**: Can be easily integrated into automated testing pipelines

## Use case: Gold Prices in India
The `goldprices.ps1` script demonstrates how to fetch gold prices in India from the international market. This is particularly useful for gold dealers and investors who need to know the price of gold in India.

### Features:

- **Real-time Pricing**: Fetch gold prices in real-time from the international market
- **No Dependencies**: Pure PowerShell - no additional tools or setup required
- **Quick Testing**: Test endpoints instantly without frontend or external tools
- **International price**: Base price in INR per gram (from USD/oz)
- **Breakdown**: Shows the cost breakdown
- **Indian market price**: Applies import duty, GST, and premiums 
  - **Import duty**: 15%
  - **GST**: 3% (on price after import duty)
  - **Premiums**: ~2% (dealer premiums, logistics, hedging)
  - **Dual display**: Shows both international and Indian prices
- **Excel Output**: Saves the results to an Excel file (goldprices.xlsx)
- **MP3 Output**: Saves the results to an MP3 file (Output.mp3)

### Prerequisites:

- Windows PowerShell (5.1 or later) or PowerShell Core
- Internet connection

### Usage:

1. **Run the gold prices script**:
   ```powershell
   .\goldprices.ps1
   ```
2. **Modes of running the script**:
   - **Interactive mode**: Displays the results in a formatted output (default)
   - **Quiet mode**: Displays only the Indian market price (useful for automation)
   
   ```powershell
   .\goldprices.ps1 -Quiet
   ```

### What the Script Does:

The script performs the following steps:

1. **Fetch gold prices**: Fetches gold prices in real-time from the international market and converts it to INR per gram
2. **Calculate Indian market price**: Calculates the Indian market price of gold including import duty, GST, and premiums
3. **Display results**: Displays the results in a formatted output

### Example Output

```
Fetching 24 Karat Gold Price (INR per gram)...
Fetching gold price from web (goldprice.org)...
Success!
Exchange rate: 1 USD = Rs. 83.0
```

## Use case: Gold Prices Monitor

The `goldprices-monitor.ps1` script demonstrates how to monitor gold prices continuously in real-time. This is particularly useful for gold dealers and investors who need to know the price of gold in India continuously.

### Features:

- **Real-time Pricing**: Fetch gold prices in real-time from the international market
- **No Dependencies**: Pure PowerShell - no additional tools or setup required
- **Quick Testing**: Test endpoints instantly without frontend or external tools
- **International price**: Base price in INR per gram (from USD/oz)
- **Continuous Monitoring**: Monitors gold prices continuously in real-time
- **Timestamp**: Shows the timestamp of the price
- **Interval**: Sets the interval for the price update
- **Both options**: Shows both the timestamp and the price

### Basic usage (updates every 1 minute)
```powershell
.\goldprices-monitor.ps1
```
### With timestamps
```powershell
.\goldprices-monitor.ps1 -ShowTimestamp
```
### Custom interval (e.g., every 30 seconds)
```powershell
.\goldprices-monitor.ps1 -IntervalSeconds 30
```
### Both options (every 30 seconds with timestamps)
```powershell
.\goldprices-monitor.ps1 -IntervalSeconds 30 -ShowTimestamp
```

### What the Script Does:

The script performs the following steps:

1. **Fetch gold prices**: Fetches gold prices in real-time from the international market and converts it to INR per gram
2. **Calculate Indian market price**: Calculates the Indian market price of gold including import duty, GST, and premiums
3. **Display results**: Displays the results in a formatted output

ElevenLabs TTS full features basic usage:
# List voices and save to JSON
.\elevenlabs-tts.ps1 -Action voices -SaveVoicesToJson

# List TTS models
.\elevenlabs-tts.ps1 -Action models

# Check subscription and usage
.\elevenlabs-tts.ps1 -Action subscription

# Generate speech (default voice: Sarah)
.\elevenlabs-tts.ps1 -Action convert -Text "Hello from ElevenLabs" -OutputFile out.mp3 -PlayAudio

# Tune voice behavior
.\elevenlabs-tts.ps1 -Action convert -Text "Custom tone" -Stability 0.7 -SimilarityBoost 0.8 -ModelId eleven_flash_v2_5

# Transcribe audio
.\elevenlabs-tts.ps1 -Action transcribe -InputFile audio.mp3

# View a single voice and its settings
.\elevenlabs-tts.ps1 -Action voice -VoiceId EXAVITQu4vr4xnSDxMaL
.\elevenlabs-tts.ps1 -Action settings -VoiceId EXAVITQu4vr4xnSDxMaL

# List history
.\elevenlabs-tts.ps1 -Action history -HistoryLimit 20


# Basic usage (uses defaults)
.\openai-tts.ps1

# Custom text and voice
.\openai-tts.ps1 -InputText "Hello, world!" -Voice "alloy"

# Custom output file
.\openai-tts.ps1 -OutputFile "my-speech.mp3"

# Full customization
.\openai-tts.ps1 -InputText "Custom text" -Voice "nova" -OutputFile "output.mp3"

# Play audio
.\openai-tts.ps1 -PlayAudio

# Full customization with play
.\openai-tts.ps1 -InputText "Custom text" -Voice "nova" -OutputFile "output.mp3" -PlayAudio

