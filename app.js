/* ========================================
   DataBot – Core Application Logic
   Excel parsing, Gemini AI chatbot
   ======================================== */

// ===================== GLOBAL STATE =====================
let excelData = [];
let columns = [];
const GEMINI_API_KEY = 'AIzaSyBm-_MptEa5CU9HggioZis9XY4MnReJqao';
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

// ===================== INITIALIZATION =====================
document.addEventListener('DOMContentLoaded', () => {
    initParticles();
    initTabs();
    initChatInput();
    loadExcelData();
});

// ===================== BACKGROUND PARTICLES =====================
function initParticles() {
    const container = document.getElementById('bgParticles');
    for (let i = 0; i < 30; i++) {
        const p = document.createElement('div');
        p.className = 'particle';
        const size = Math.random() * 4 + 2;
        p.style.width = size + 'px';
        p.style.height = size + 'px';
        p.style.left = Math.random() * 100 + '%';
        p.style.animationDuration = (Math.random() * 20 + 15) + 's';
        p.style.animationDelay = (Math.random() * 10) + 's';
        container.appendChild(p);
    }
}

// ===================== TAB NAVIGATION =====================
function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tab = btn.dataset.tab;
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(tab + 'Tab').classList.add('active');
        });
    });
}

// ===================== CHAT INPUT =====================
function initChatInput() {
    const input = document.getElementById('chatInput');
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && input.value.trim()) {
            sendMessage();
        }
    });
}

// ===================== EXCEL DATA LOADING =====================
async function loadExcelData() {
    try {
        const response = await fetch('sample_data.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(sheet);
        columns = excelData.length > 0 ? Object.keys(excelData[0]) : [];
        renderTable(excelData);
        renderStats();
        initTableSearch();
        document.getElementById('rowCount').textContent = `${excelData.length} records loaded • ${columns.length} columns`;
    } catch (err) {
        console.error('Error loading Excel:', err);
        document.getElementById('rowCount').textContent = '⚠ Could not load data file. Make sure sample_data.xlsx exists.';
    }
}

// ===================== TABLE RENDERING =====================
function renderTable(data) {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');

    // Header
    thead.innerHTML = '<tr>' + columns.map(col => `<th>${escapeHtml(col)}</th>`).join('') + '</tr>';

    // Body
    tbody.innerHTML = data.map(row =>
        '<tr>' + columns.map(col => `<td>${escapeHtml(String(row[col] ?? ''))}</td>`).join('') + '</tr>'
    ).join('');
}

function initTableSearch() {
    document.getElementById('tableSearch').addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (!query) {
            renderTable(excelData);
            return;
        }
        const filtered = excelData.filter(row =>
            columns.some(col => String(row[col] ?? '').toLowerCase().includes(query))
        );
        renderTable(filtered);
    });
}

// ===================== STATS RENDERING =====================
function renderStats() {
    const grid = document.getElementById('statsGrid');
    const stats = [];

    stats.push({ label: 'Total Employees', value: excelData.length, detail: 'Records in dataset' });

    // Unique departments
    const depts = [...new Set(excelData.map(r => r['Department']).filter(Boolean))];
    if (depts.length) stats.push({ label: 'Departments', value: depts.length, detail: depts.join(', ') });

    // Average salary
    const salaries = excelData.map(r => Number(r['Salary'])).filter(s => !isNaN(s));
    if (salaries.length) {
        const avg = salaries.reduce((a, b) => a + b, 0) / salaries.length;
        stats.push({ label: 'Avg Salary', value: '$' + Math.round(avg).toLocaleString(), detail: `Range: $${Math.min(...salaries).toLocaleString()} – $${Math.max(...salaries).toLocaleString()}` });
    }

    // Unique cities
    const cities = [...new Set(excelData.map(r => r['City']).filter(Boolean))];
    if (cities.length) stats.push({ label: 'Locations', value: cities.length, detail: cities.join(', ') });

    grid.innerHTML = stats.map(s => `
        <div class="stat-card">
            <div class="stat-label">${s.label}</div>
            <div class="stat-value">${s.value}</div>
            <div class="stat-detail">${s.detail}</div>
        </div>
    `).join('');
}

// ===================== CHATBOT ENGINE (GEMINI AI) =====================
async function sendMessage() {
    const input = document.getElementById('chatInput');
    const text = input.value.trim();
    if (!text) return;

    addMessage(text, 'user');
    input.value = '';
    input.disabled = true;
    document.getElementById('sendBtn').disabled = true;

    // Show typing indicator
    const typingEl = showTyping();

    try {
        const response = await askGemini(text);
        typingEl.remove();
        addMessage(response, 'bot');
    } catch (err) {
        console.error('Gemini API error:', err);
        typingEl.remove();
        addMessage('<p>⚠ AI is temporarily unavailable. Using local search instead...</p>' + processQueryLocal(text), 'bot');
    } finally {
        input.disabled = false;
        document.getElementById('sendBtn').disabled = false;
        input.focus();
    }
}

// ===================== GEMINI API CALL =====================
async function askGemini(question) {
    // Build data context for Gemini
    const dataContext = JSON.stringify(excelData, null, 2);

    const systemPrompt = `You are DataBot, a helpful assistant for Mahabyte Employee Data. You answer questions based ONLY on the employee data provided below. 

RULES:
- Answer based on the data below. Do not make up information.
- Be concise and friendly. Use emojis sparingly.
- When listing employees, format as an HTML table with <table class="chat-result-table">, <thead>, <tbody>, <th>, <td> tags.
- For salary values, format with $ sign and commas (e.g., $85,000).
- When giving counts or statistics, bold the numbers using <strong> tags.
- Wrap text responses in <p> tags.
- If the question is unrelated to the data, politely say you can only help with employee data questions.

EMPLOYEE DATA:
${dataContext}`;

    const response = await fetch(GEMINI_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            contents: [{
                parts: [{ text: systemPrompt + '\n\nUser question: ' + question }]
            }],
            generationConfig: {
                temperature: 0.3,
                maxOutputTokens: 1024
            }
        })
    });

    if (!response.ok) {
        throw new Error(`API error: ${response.status}`);
    }

    const data = await response.json();
    const aiText = data.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!aiText) throw new Error('Empty response from Gemini');

    // Clean up markdown-style formatting to HTML
    return formatGeminiResponse(aiText);
}

function formatGeminiResponse(text) {
    // If Gemini already returned HTML tags, use as-is
    if (text.includes('<table') || text.includes('<p>') || text.includes('<strong>')) {
        // Clean any markdown code fences
        return text.replace(/```html\n?/g, '').replace(/```\n?/g, '');
    }
    // Convert markdown-style response to HTML
    let html = text
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')  // bold
        .replace(/\*(.+?)\*/g, '<em>$1</em>')               // italic
        .replace(/\n\n/g, '</p><p>')                         // paragraphs
        .replace(/\n/g, '<br>');                             // line breaks
    return '<p>' + html + '</p>';
}

function askSuggestion(btn) {
    document.getElementById('chatInput').value = btn.textContent;
    sendMessage();
}

function addMessage(content, type) {
    const container = document.getElementById('chatMessages');
    const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
    const avatar = type === 'bot' ? '🤖' : '👤';

    const div = document.createElement('div');
    div.className = `message ${type}-message animate-in`;
    div.innerHTML = `
        <div class="message-avatar">${avatar}</div>
        <div class="message-content">
            <div class="message-bubble">${content}</div>
            <span class="message-time">${time}</span>
        </div>
    `;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
}

function showTyping() {
    const container = document.getElementById('chatMessages');
    const div = document.createElement('div');
    div.className = 'message bot-message animate-in';
    div.innerHTML = `
        <div class="message-avatar">🤖</div>
        <div class="message-content">
            <div class="message-bubble">
                <div class="typing-indicator">
                    <div class="typing-dot"></div>
                    <div class="typing-dot"></div>
                    <div class="typing-dot"></div>
                </div>
            </div>
        </div>
    `;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
    return div;
}

// ===================== LOCAL FALLBACK SEARCH =====================
function processQueryLocal(query) {
    if (excelData.length === 0) {
        return '<p>No data loaded.</p>';
    }
    const q = query.toLowerCase().trim();

    // Quick keyword search
    const keywords = q.replace(/[?.,!]/g, '').split(/\s+/).filter(w => w.length > 2);
    const matches = excelData.filter(row =>
        keywords.some(kw =>
            columns.some(col => String(row[col] || '').toLowerCase().includes(kw))
        )
    );

    if (matches.length > 0) {
        const displayCols = matches.length <= 5 ? columns.slice(0, 7) : ['Name', 'Department', 'Role', 'City'];
        return `<p>🔍 Found <strong>${matches.length} result${matches.length > 1 ? 's' : ''}</strong>:</p>` + buildMiniTable(matches.slice(0, 10), displayCols);
    }
    return '<p>No results found. Try different keywords.</p>';
}

// ===================== HELPERS =====================
function buildMiniTable(data, cols) {
    const availCols = cols.filter(c => columns.includes(c));
    if (availCols.length === 0) return '<p>No matching columns found.</p>';

    let html = '<table class="chat-result-table"><thead><tr>';
    html += availCols.map(c => `<th>${escapeHtml(c)}</th>`).join('');
    html += '</tr></thead><tbody>';
    data.forEach(row => {
        html += '<tr>';
        html += availCols.map(c => {
            let val = row[c] ?? '';
            if (c === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
            return `<td>${escapeHtml(String(val))}</td>`;
        }).join('');
        html += '</tr>';
    });
    html += '</tbody></table>';
    return html;
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}
