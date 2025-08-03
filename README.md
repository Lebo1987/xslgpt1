# XSLGPT1 - AI-Powered Excel Formula Generator

An Excel Add-in that helps users generate Excel formulas using AI. Users can describe what they want to calculate in natural language and get the corresponding Excel formula with an explanation.

## Features

- **Natural Language Input**: Describe what you want to calculate in plain English
- **AI-Powered Generation**: Get accurate Excel formulas based on your description
- **Formula Explanation**: Understand what each formula does
- **One-Click Insert**: Insert formulas directly into Excel cells
- **Copy to Clipboard**: Copy formulas for use elsewhere
- **Modern UI**: Clean, professional interface with smooth animations

## Prerequisites

- Node.js (version 14 or higher)
- Excel (desktop or online)
- A modern web browser

## Installation & Setup

### 1. Install Dependencies

```bash
npm install
```

### 2. Build the Project

```bash
npm run build
```

### 3. Start the Development Server

```bash
npm run dev
```

The server will start on `http://localhost:8080`

### 4. Sideload the Add-in in Excel

1. Open Excel
2. Go to **Insert** > **My Add-ins** > **Upload My Add-in**
3. Browse to the `manifest-local.xml` file in this project
4. Click **Upload**

The add-in will appear in the **HOME** tab with a "Show Taskpane" button.

## Usage

1. **Open the Taskpane**: Click the "Show Taskpane" button in the HOME tab
2. **Describe Your Calculation**: Type what you want to calculate (e.g., "Sum from A1 to A10")
3. **Generate Formula**: Click "Generate Formula" or press Enter
4. **Review Results**: See the generated formula and explanation
5. **Insert or Copy**: Click "Insert into Excel" to add it to your selected cell, or "Copy" to copy to clipboard

## Example Prompts

- "Sum from A1 to A10"
- "Average of B1:B20"
- "Count non-empty cells in C1:C50"
- "Find maximum value in D1:D100"
- "Find minimum value in E1:E25"

## Project Structure

```
XSLGPT1/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Main taskpane UI
│   │   ├── taskpane.js        # Taskpane functionality
│   │   └── taskpane.css       # Styling
│   └── commands/
│       └── commands.html      # Office commands
├── assets/                    # Icons and static assets
├── dist/                      # Built files (generated)
├── server.js                  # Backend API server
├── webpack.config.js          # Build configuration
├── package.json               # Dependencies and scripts
├── manifest-local.xml         # Office Add-in manifest
└── README.md                  # This file
```

## Development

### Available Scripts

- `npm start` - Start the production server
- `npm run dev` - Start the development server with auto-reload
- `npm run build` - Build the project for production
- `npm run build:dev` - Build the project for development
- `npm run watch` - Watch for changes and rebuild automatically

### API Endpoints

- `POST /api/generate` - Generate formula from prompt
  - Body: `{ "prompt": "your description here" }`
  - Response: `{ "success": true, "formula": "=SUM(A1:A10)", "explanation": "..." }`

## Configuration

### Environment Variables

Create a `.env` file in the root directory:

```env
PORT=8080
OPENAI_API_KEY=your_openai_api_key_here
```

### OpenAI Integration

To use real AI instead of mock data:

1. Get an OpenAI API key from [OpenAI Platform](https://platform.openai.com/)
2. Add your API key to the `.env` file
3. Uncomment the OpenAI integration code in `server.js`

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Make sure the server is running on `localhost:8080`
2. **CORS errors**: The server includes CORS headers, but check your browser console
3. **Manifest errors**: Ensure the manifest file points to the correct URLs
4. **Build errors**: Make sure all dependencies are installed

### Debug Mode

To enable debug logging, set the `DEBUG` environment variable:

```bash
DEBUG=* npm run dev
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For issues and questions:
- Check the troubleshooting section above
- Review the Office Add-ins documentation
- Open an issue in this repository 