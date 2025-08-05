import './taskpane.css';

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('generate-btn').onclick = generateFormula;
        document.getElementById('copy-btn').onclick = copyFormula;
        document.getElementById('insert-btn').onclick = insertFormula;
        
        // Enable Enter key to generate formula
        document.getElementById('prompt-input').addEventListener('keydown', (e) => {
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                generateFormula();
            }
        });

        // Load user email from localStorage or use default
        loadUserEmail();
    }
});

function loadUserEmail() {
    const userEmail = localStorage.getItem('userEmail') || 'default@example.com';
    document.getElementById('user-email').value = userEmail;
}

function saveUserEmail(email) {
    localStorage.setItem('userEmail', email);
}

async function generateFormula() {
    const promptInput = document.getElementById('prompt-input');
    const userEmailInput = document.getElementById('user-email');
    const generateBtn = document.getElementById('generate-btn');
    const resultSection = document.getElementById('result-section');
    const errorSection = document.getElementById('error-section');
    
    const prompt = promptInput.value.trim();
    const userEmail = userEmailInput.value.trim() || 'default@example.com';
    
    if (!prompt) {
        showError('Please enter a description of what you want to calculate.');
        return;
    }
    
    // Save user email
    saveUserEmail(userEmail);
    
    // Show loading state
    generateBtn.classList.add('loading');
    generateBtn.disabled = true;
    hideError();
    hideResult();
    
    try {
        const response = await fetch('https://xslgpt1-excel-addin.onrender.com/api/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ prompt, userEmail })
        });
        
        const data = await response.json();
        
        if (response.status === 429) {
            // Usage limit exceeded
            showUsageLimitError(data.error, data.currentUsage, data.limit);
        } else if (data.success) {
            showResult(data.formula, data.explanation, data.usage);
        } else {
            showError(data.error || 'Failed to generate formula');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Failed to connect to the server. Please try again later.');
    } finally {
        // Reset loading state
        generateBtn.classList.remove('loading');
        generateBtn.disabled = false;
    }
}

function showResult(formula, explanation, usage) {
    const resultSection = document.getElementById('result-section');
    const formulaDisplay = document.getElementById('formula-display');
    const explanationDisplay = document.getElementById('explanation-display');
    const usageDisplay = document.getElementById('usage-display');
    
    formulaDisplay.textContent = formula;
    explanationDisplay.textContent = explanation;
    
    // Show usage information
    if (usage) {
        usageDisplay.innerHTML = `
            <div class="usage-info">
                <span class="usage-text">Usage: ${usage.current}/${usage.limit} messages</span>
                <span class="usage-remaining">(${usage.remaining} remaining)</span>
            </div>
        `;
        usageDisplay.style.display = 'block';
    } else {
        usageDisplay.style.display = 'none';
    }
    
    resultSection.style.display = 'block';
    hideError();
}

function showUsageLimitError(message, currentUsage, limit) {
    const errorSection = document.getElementById('error-section');
    const errorMessage = document.getElementById('error-message');
    
    errorMessage.innerHTML = `
        <div class="usage-limit-error">
            <div class="error-text">${message}</div>
            <div class="usage-details">
                You've used ${currentUsage} of ${limit} free messages this month.
            </div>
            <div class="upgrade-suggestion">
                Contact support to upgrade your plan for unlimited access.
            </div>
        </div>
    `;
    errorSection.style.display = 'block';
    hideResult();
}

function showError(message) {
    const errorSection = document.getElementById('error-section');
    const errorMessage = document.getElementById('error-message');
    
    errorMessage.textContent = message;
    errorSection.style.display = 'block';
    hideResult();
}

function hideResult() {
    document.getElementById('result-section').style.display = 'none';
}

function hideError() {
    document.getElementById('error-section').style.display = 'none';
}

function copyFormula() {
    const formulaDisplay = document.getElementById('formula-display');
    const formula = formulaDisplay.textContent;
    
    navigator.clipboard.writeText(formula).then(() => {
        const copyBtn = document.getElementById('copy-btn');
        const originalText = copyBtn.textContent;
        copyBtn.textContent = 'Copied!';
        copyBtn.style.background = '#218838';
        
        setTimeout(() => {
            copyBtn.textContent = originalText;
            copyBtn.style.background = '#28a745';
        }, 2000);
    }).catch(err => {
        console.error('Failed to copy formula:', err);
        showError('Failed to copy formula to clipboard');
    });
}

function insertFormula() {
    const formulaDisplay = document.getElementById('formula-display');
    const formula = formulaDisplay.textContent;
    
    Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.formulas = [[formula]];
        
        await context.sync();
        
        // Show success feedback
        const insertBtn = document.getElementById('insert-btn');
        const originalText = insertBtn.textContent;
        insertBtn.textContent = 'Inserted!';
        insertBtn.style.background = '#28a745';
        
        setTimeout(() => {
            insertBtn.textContent = originalText;
            insertBtn.style.background = '#6c757d';
        }, 2000);
    }).catch(error => {
        console.error('Error inserting formula:', error);
        showError('Failed to insert formula into Excel');
    });
} 