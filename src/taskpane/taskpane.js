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
    }
});

async function generateFormula() {
    const promptInput = document.getElementById('prompt-input');
    const generateBtn = document.getElementById('generate-btn');
    const resultSection = document.getElementById('result-section');
    const errorSection = document.getElementById('error-section');
    
    const prompt = promptInput.value.trim();
    
    if (!prompt) {
        showError('Please enter a description of what you want to calculate.');
        return;
    }
    
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
            body: JSON.stringify({ prompt })
        });
        
        const data = await response.json();
        
        if (data.success) {
            showResult(data.formula, data.explanation);
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

function showResult(formula, explanation) {
    const resultSection = document.getElementById('result-section');
    const formulaDisplay = document.getElementById('formula-display');
    const explanationDisplay = document.getElementById('explanation-display');
    
    formulaDisplay.textContent = formula;
    explanationDisplay.textContent = explanation;
    
    resultSection.style.display = 'block';
    hideError();
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