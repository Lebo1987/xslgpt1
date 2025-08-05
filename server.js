const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const OpenAI = require('openai');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 8080;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;

// Initialize OpenAI client
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

// Usage tracking file
const USAGE_FILE = 'usage_data.json';
const MONTHLY_LIMIT = 10;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'dist')));

// Load usage data from file
function loadUsageData() {
  try {
    if (fs.existsSync(USAGE_FILE)) {
      const data = fs.readFileSync(USAGE_FILE, 'utf8');
      return JSON.parse(data);
    }
  } catch (error) {
    console.error('Error loading usage data:', error);
  }
  return {};
}

// Save usage data to file
function saveUsageData(usageData) {
  try {
    fs.writeFileSync(USAGE_FILE, JSON.stringify(usageData, null, 2));
  } catch (error) {
    console.error('Error saving usage data:', error);
  }
}

// Get current month key (YYYY-MM format)
function getCurrentMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

// Check if user has exceeded monthly limit
function checkUsageLimit(userEmail) {
  const usageData = loadUsageData();
  const currentMonth = getCurrentMonth();
  
  // Initialize user data if not exists
  if (!usageData[userEmail]) {
    usageData[userEmail] = {};
  }
  
  if (!usageData[userEmail][currentMonth]) {
    usageData[userEmail][currentMonth] = 0;
  }
  
  const currentUsage = usageData[userEmail][currentMonth];
  
  if (currentUsage >= MONTHLY_LIMIT) {
    return { exceeded: true, currentUsage };
  }
  
  return { exceeded: false, currentUsage };
}

// Increment user usage
function incrementUsage(userEmail) {
  const usageData = loadUsageData();
  const currentMonth = getCurrentMonth();
  
  // Initialize user data if not exists
  if (!usageData[userEmail]) {
    usageData[userEmail] = {};
  }
  
  if (!usageData[userEmail][currentMonth]) {
    usageData[userEmail][currentMonth] = 0;
  }
  
  usageData[userEmail][currentMonth]++;
  saveUsageData(usageData);
  
  return usageData[userEmail][currentMonth];
}

// Generate Excel formula using OpenAI
async function generateFormulaWithAI(prompt) {
  try {
    const systemPrompt = `You are an Excel formula expert. Generate Excel formulas based on user descriptions. 
    
    Rules:
    1. Always return ONLY the Excel formula (e.g., =SUM(A1:A10))
    2. Provide a brief explanation of what the formula does
    3. Use standard Excel functions like SUM, AVERAGE, COUNT, MAX, MIN, VLOOKUP, IF, etc.
    4. Make formulas practical and commonly used
    5. If the request is unclear, ask for clarification or provide a basic SUM formula
    
    Format your response as JSON:
    {
      "formula": "=SUM(A1:A10)",
      "explanation": "This formula sums all values in cells A1 through A10"
    }`;

    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: prompt }
      ],
      max_tokens: 200,
      temperature: 0.3
    });

    const response = completion.choices[0].message.content;
    
    // Try to parse as JSON, if it fails, create a structured response
    try {
      return JSON.parse(response);
    } catch (parseError) {
      // If response isn't valid JSON, create a structured response
      return {
        formula: response.trim(),
        explanation: "Generated Excel formula based on your request"
      };
    }
  } catch (error) {
    console.error('OpenAI API Error:', error);
    throw new Error('Failed to generate formula with AI');
  }
}

// API endpoint for formula generation
app.post('/api/generate', async (req, res) => {
  try {
    const { prompt, userEmail = 'default@example.com' } = req.body;
    
    if (!prompt) {
      return res.status(400).json({ error: 'Prompt is required' });
    }

    // Check usage limit
    const usageCheck = checkUsageLimit(userEmail);
    if (usageCheck.exceeded) {
      return res.status(429).json({
        error: "You've reached your 10 free messages. Upgrade to continue →",
        currentUsage: usageCheck.currentUsage,
        limit: MONTHLY_LIMIT
      });
    }

    // Generate formula using OpenAI
    const result = await generateFormulaWithAI(prompt);

    // Increment usage after successful generation
    const newUsage = incrementUsage(userEmail);

    res.json({
      success: true,
      formula: result.formula,
      explanation: result.explanation,
      usage: {
        current: newUsage,
        limit: MONTHLY_LIMIT,
        remaining: MONTHLY_LIMIT - newUsage
      }
    });

  } catch (error) {
    console.error('Error generating formula:', error);
    res.status(500).json({ 
      error: 'Failed to generate formula. Please try again.',
      details: error.message 
    });
  }
});

// API endpoint to check user usage
app.get('/api/usage/:userEmail', (req, res) => {
  try {
    const { userEmail } = req.params;
    const usageCheck = checkUsageLimit(userEmail);
    const currentMonth = getCurrentMonth();
    
    res.json({
      userEmail,
      currentMonth,
      currentUsage: usageCheck.currentUsage,
      limit: MONTHLY_LIMIT,
      remaining: MONTHLY_LIMIT - usageCheck.currentUsage,
      exceeded: usageCheck.exceeded
    });
  } catch (error) {
    console.error('Error checking usage:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// Serve the manifest.xml file
app.get('/manifest.xml', (req, res) => {
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Serve the taskpane HTML
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'taskpane.html'));
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`API endpoint: ${BASE_URL}/api/generate`);
  console.log(`Usage tracking enabled - limit: ${MONTHLY_LIMIT} messages per month`);
  console.log(`OpenAI integration enabled`);
}); 