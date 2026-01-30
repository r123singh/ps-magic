# PowerShell Magic for different use cases

A collection of PowerShell scripts for various development and testing scenarios. These scripts showcase how terminal scripts in Windows PowerShell can simplify common tasks without needing to open browsers, prepare Postman collections, or set up complex testing environments.

## Use case: Test API Endpoint

The `test-endpoint.ps1` script demonstrates how to test API endpoints directly from PowerShell. This is particularly useful for backend development where you need to quickly test endpoints without opening a browser or configuring Postman collections.

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