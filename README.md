# MST Data Processor Web App

A fully client-side web application for processing magnetic sensor test (MST) data. This application runs entirely in your browser using Python via WebAssembly - no server required, and your data never leaves your computer.

## Features

- **100% Client-Side Processing**: All Excel processing happens in your browser using Pyodide (Python compiled to WebAssembly)
- **Complete Privacy**: Your confidential Excel files are never uploaded to any server
- **No Installation Required**: No need to install Python or any dependencies
- **Easy to Use**: Simple drag-and-drop interface
- **Configurable Mappings**: Customize sample and axis mappings directly in the UI

## How to Use

1. **Open the Application**
   - Simply open `index.html` in a modern web browser (Chrome, Firefox, Edge, Safari)
   - Wait for the Python environment to initialize (first load may take 1-2 minutes)

2. **Upload Your Excel File**
   - Click on the upload area or drag and drop your `Book1.xlsx` file
   - The file is loaded into your browser's memory only

3. **Configure Mappings**
   - Edit the Sample Mapping (e.g., `1: L1C_31`)
   - Edit the Axis Mapping (e.g., `1: XY`)

4. **Process Data**
   - Click the "Process Data" button
   - Wait for processing to complete (usually takes 10-30 seconds)

5. **Download Results**
   - Click the download button to save your processed Excel file

## Technical Details

### What This App Does

The application processes magnetic sensor test data with 2 sweeps (down-up):
- Skips sheets containing "precon" in the name
- Processes data in groups of 5 rows (2 sweeps with overlapping points)
- Calculates offset values (Off_DU)
- Calculates sensitivity values (Sens_DU)
- Calculates drift values for both offset and sensitivity
- Generates summary sheets with out-of-limit violations

### Browser Compatibility

- ✅ Chrome/Edge (recommended)
- ✅ Firefox
- ✅ Safari
- ⚠️ Requires a modern browser with WebAssembly support

### File Size Limitations

- Input files up to ~50MB should work fine
- Larger files may take longer to process
- Processing time depends on your computer's performance

## Privacy & Security

This application is designed with privacy in mind:
- **No Network Requests**: After the initial page load, no data is sent over the network
- **Local Processing**: All Python code runs in your browser using WebAssembly
- **No Tracking**: No analytics, cookies, or tracking of any kind
- **Open Source**: All code is visible and can be audited

## Running Locally

Simply open `index.html` in your browser. No web server is required, though you can use one if you prefer:

```bash
# Using Python's built-in server
python -m http.server 8000

# Or using Node.js
npx serve
```

Then navigate to `http://localhost:8000`

## Files

- `index.html` - Main HTML page with UI
- `app.js` - JavaScript code handling file processing and Pyodide interaction
- `README.md` - This file

## Dependencies

The app uses:
- [Pyodide](https://pyodide.org/) - Python runtime for WebAssembly
- pandas, numpy, openpyxl - Loaded automatically via Pyodide

All dependencies are loaded from CDN on first use and cached by your browser.

## Troubleshooting

**Python environment fails to load**
- Check your internet connection (needed for first load only)
- Try refreshing the page
- Clear browser cache and try again

**Processing takes too long**
- Large files naturally take longer
- Close other browser tabs to free up memory
- Try using a desktop browser instead of mobile

**Download doesn't work**
- Check if your browser is blocking downloads
- Try using a different browser
- Check browser console for errors (F12)

## License

This application is provided as-is for processing MST data. Feel free to modify and use as needed.
