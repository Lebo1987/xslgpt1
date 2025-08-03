const express = require('express');
const cors = require('cors');
const path = require('path');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 8080;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'dist')));

// Mock data for formula generation (can be replaced with OpenAI integration)
const mockFormulas = {
  'sum from a1 to a10': {
    formula: '=SUM(A1:A10)',
    explanation: 'This formula sums all values in cells A1 through A10'
  },
  'average b1 to b20': {
    formula: '=AVERAGE(B1:B20)',
    explanation: 'This formula calculates the average of all values in cells B1 through B20'
  },
  'count non-empty cells in c1:c50': {
    formula: '=COUNTA(C1:C50)',
    explanation: 'This formula counts all non-empty cells in the range C1:C50'
  },
  'find maximum value in d1:d100': {
    formula: '=MAX(D1:D100)',
    explanation: 'This formula finds the highest value in cells D1 through D100'
  },
  'find minimum value in e1:e25': {
    formula: '=MIN(E1:E25)',
    explanation: 'This formula finds the lowest value in cells E1 through E25'
  }
};

// API endpoint for formula generation
app.post('/api/generate', (req, res) => {
  try {
    const { prompt } = req.body;
    
    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    const lowerPrompt = prompt.toLowerCase();
    let result = null;

    // Check for exact matches first
    for (const [key, value] of Object.entries(mockFormulas)) {
      if (lowerPrompt.includes(key)) {
        result = value;
        break;
      }
    }

    // If no exact match, try to generate based on keywords
    if (!result) {
      if (lowerPrompt.includes('sum') || lowerPrompt.includes('add')) {
        result = {
          formula: '=SUM(A1:A10)',
          explanation: 'Generated sum formula - adjust the range as needed'
        };
      } else if (lowerPrompt.includes('average') || lowerPrompt.includes('mean')) {
        result = {
          formula: '=AVERAGE(A1:A10)',
          explanation: 'Generated average formula - adjust the range as needed'
        };
      } else if (lowerPrompt.includes('count')) {
        result = {
          formula: '=COUNTA(A1:A10)',
          explanation: 'Generated count formula - adjust the range as needed'
        };
      } else if (lowerPrompt.includes('max') || lowerPrompt.includes('maximum')) {
        result = {
          formula: '=MAX(A1:A10)',
          explanation: 'Generated maximum formula - adjust the range as needed'
        };
      } else if (lowerPrompt.includes('min') || lowerPrompt.includes('minimum')) {
        result = {
          formula: '=MIN(A1:A10)',
          explanation: 'Generated minimum formula - adjust the range as needed'
        };
      } else {
        result = {
          formula: '=SUM(A1:A10)',
          explanation: 'Default formula - please specify your requirements more clearly'
        };
      }
    }

    res.json({
      success: true,
      formula: result.formula,
      explanation: result.explanation
    });

  } catch (error) {
    console.error('Error generating formula:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Serve the taskpane HTML
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'taskpane.html'));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`API endpoint: ${BASE_URL}/api/generate`);
}); 