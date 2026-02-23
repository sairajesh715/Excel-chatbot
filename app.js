/* ========================================
   DataBot – Core Application Logic
   Excel parsing, smart search, chatbot
   ======================================== */

// ===================== GLOBAL STATE =====================
let excelData = [];
let columns = [];

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

// ===================== CHATBOT ENGINE =====================
function sendMessage() {
    const input = document.getElementById('chatInput');
    const text = input.value.trim();
    if (!text) return;

    addMessage(text, 'user');
    input.value = '';

    // Show typing indicator
    const typingEl = showTyping();

    // Simulate thinking delay
    setTimeout(() => {
        typingEl.remove();
        const response = processQuery(text);
        addMessage(response, 'bot');
    }, 600 + Math.random() * 600);
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

// ===================== QUERY PROCESSOR =====================
function processQuery(query) {
    if (excelData.length === 0) {
        return '<p>⚠ No data loaded. Please make sure the Excel file is available.</p>';
    }

    const q = query.toLowerCase().trim();

    // --- Greeting ---
    if (/^(hi|hello|hey|good morning|good afternoon|good evening)\b/.test(q)) {
        return `<p>Hello! 👋 I'm ready to help you explore the employee data.</p>
                <p>We have <strong>${excelData.length} employees</strong> across <strong>${[...new Set(excelData.map(r => r['Department']))].length} departments</strong>. Ask me anything!</p>`;
    }

    // --- Total count ---
    if (/how many (employees|people|records|rows|staff|workers)/.test(q) || /total (employees|count|records)/.test(q) || q === 'count') {
        return `<p>There are <strong>${excelData.length} employees</strong> in the dataset.</p>`;
    }

    // --- List all departments ---
    if (/list.*(department|dept)s?|all.*(department|dept)s?|what.*(department|dept)s?|show.*(department|dept)s?/.test(q)) {
        const depts = [...new Set(excelData.map(r => r['Department']).filter(Boolean))];
        const deptCounts = depts.map(d => {
            const count = excelData.filter(r => r['Department'] === d).length;
            return `<strong>${d}</strong>: ${count} employee${count > 1 ? 's' : ''}`;
        });
        return `<p>📋 Departments (${depts.length}):</p><ul>${deptCounts.map(d => '<li>' + d + '</li>').join('')}</ul>`;
    }

    // --- List all cities/locations ---
    if (/list.*(city|cities|location)s?|all.*(city|cities|location)s?|what.*(city|cities|location)s?|show.*(city|cities|location)s?/.test(q)) {
        const cities = [...new Set(excelData.map(r => r['City']).filter(Boolean))];
        const cityCounts = cities.map(c => {
            const count = excelData.filter(r => r['City'] === c).length;
            return `<strong>${c}</strong>: ${count} employee${count > 1 ? 's' : ''}`;
        });
        return `<p>📍 Locations (${cities.length}):</p><ul>${cityCounts.map(c => '<li>' + c + '</li>').join('')}</ul>`;
    }

    // --- Average salary ---
    if (/average salary|avg salary|mean salary/.test(q)) {
        const salaries = excelData.map(r => Number(r['Salary'])).filter(s => !isNaN(s));
        if (salaries.length) {
            const avg = Math.round(salaries.reduce((a, b) => a + b, 0) / salaries.length);
            return `<p>💰 The average salary is <strong>$${avg.toLocaleString()}</strong></p>
                    <p>Range: $${Math.min(...salaries).toLocaleString()} – $${Math.max(...salaries).toLocaleString()}</p>`;
        }
        return '<p>Could not calculate the average salary from the data.</p>';
    }

    // --- Highest salary ---
    if (/highest salary|max salary|top salary|most paid|highest paid|top earner/.test(q)) {
        const sorted = [...excelData].sort((a, b) => Number(b['Salary'] || 0) - Number(a['Salary'] || 0));
        const top = sorted.slice(0, 3);
        return `<p>🏆 Top earners:</p>` + buildMiniTable(top, ['Name', 'Role', 'Department', 'Salary']);
    }

    // --- Lowest salary ---
    if (/lowest salary|min salary|least paid|lowest paid/.test(q)) {
        const sorted = [...excelData].sort((a, b) => Number(a['Salary'] || 0) - Number(b['Salary'] || 0));
        const bottom = sorted.slice(0, 3);
        return `<p>📉 Lowest salaries:</p>` + buildMiniTable(bottom, ['Name', 'Role', 'Department', 'Salary']);
    }

    // --- Most experience ---
    if (/most experience|highest experience|senior|most experienced/.test(q)) {
        const expCol = columns.find(c => c.toLowerCase().includes('experience'));
        if (expCol) {
            const sorted = [...excelData].sort((a, b) => Number(b[expCol] || 0) - Number(a[expCol] || 0));
            const top = sorted.slice(0, 3);
            return `<p>🎖 Most experienced employees:</p>` + buildMiniTable(top, ['Name', 'Role', 'Department', expCol]);
        }
    }

    // --- Specific person lookup ---
    const nameMatch = q.match(/(?:who is|about|info|details|show me|tell me about|find)\s+(.+)/);
    if (nameMatch) {
        const searchName = nameMatch[1].replace(/[?.,!]/g, '').trim();
        const matches = excelData.filter(r =>
            String(r['Name'] || '').toLowerCase().includes(searchName)
        );
        if (matches.length > 0) {
            return `<p>👤 Found ${matches.length} match${matches.length > 1 ? 'es' : ''}:</p>` + buildMiniTable(matches, columns.slice(0, 7));
        }
    }

    // --- Filter by department ---
    const deptMatch = findColumnMatch(q, 'Department');
    if (deptMatch) {
        const filtered = excelData.filter(r => String(r['Department'] || '').toLowerCase() === deptMatch.toLowerCase());
        if (filtered.length > 0) {
            return `<p>🏢 <strong>${deptMatch}</strong> department (${filtered.length} employee${filtered.length > 1 ? 's' : ''}):</p>` + buildMiniTable(filtered, ['Name', 'Role', 'Salary', 'City']);
        }
    }

    // --- Filter by city ---
    const cityMatch = findColumnMatch(q, 'City');
    if (cityMatch) {
        const filtered = excelData.filter(r => String(r['City'] || '').toLowerCase() === cityMatch.toLowerCase());
        if (filtered.length > 0) {
            return `<p>📍 Employees in <strong>${cityMatch}</strong> (${filtered.length}):</p>` + buildMiniTable(filtered, ['Name', 'Department', 'Role', 'Salary']);
        }
    }

    // --- Filter by role ---
    const roleMatch = findColumnMatch(q, 'Role');
    if (roleMatch) {
        const filtered = excelData.filter(r => String(r['Role'] || '').toLowerCase().includes(roleMatch.toLowerCase()));
        if (filtered.length > 0) {
            return `<p>💼 Employees with role matching "<strong>${roleMatch}</strong>" (${filtered.length}):</p>` + buildMiniTable(filtered, ['Name', 'Department', 'Salary', 'City']);
        }
    }

    // --- Salary of a specific person ---
    const salaryOf = q.match(/(?:salary of|salary for|how much does|how much is)\s+(.+?)(?:\s+(?:earn|make|get))?(?:\?|$)/);
    if (salaryOf) {
        const searchName = salaryOf[1].replace(/[?.,!'s]/g, '').trim();
        const matches = excelData.filter(r =>
            String(r['Name'] || '').toLowerCase().includes(searchName)
        );
        if (matches.length > 0) {
            return matches.map(m => `<p>💰 <strong>${m['Name']}</strong>'s salary is <strong>$${Number(m['Salary']).toLocaleString()}</strong> (${m['Role']}, ${m['Department']})</p>`).join('');
        }
    }

    // --- Show all / list all ---
    if (/show all|list all|all employees|everyone|all data|all records/.test(q)) {
        if (excelData.length > 10) {
            return `<p>📋 Showing all <strong>${excelData.length} employees</strong>:</p>` + buildMiniTable(excelData, ['Name', 'Department', 'Role', 'City']);
        }
        return `<p>📋 All employees:</p>` + buildMiniTable(excelData, columns.slice(0, 6));
    }

    // --- Generic keyword search (fallback) ---
    const keywords = q.replace(/[?.,!]/g, '').split(/\s+/).filter(w => w.length > 2);
    const matches = excelData.filter(row =>
        keywords.some(kw =>
            columns.some(col => String(row[col] || '').toLowerCase().includes(kw))
        )
    );

    if (matches.length > 0) {
        const displayCols = matches.length <= 5 ? columns.slice(0, 7) : ['Name', 'Department', 'Role', 'City'];
        return `<p>🔍 Found <strong>${matches.length} result${matches.length > 1 ? 's' : ''}</strong> matching your query:</p>` + buildMiniTable(matches.slice(0, 10), displayCols);
    }

    // --- No matches ---
    return `<p>🤔 I couldn't find results for "<em>${escapeHtml(query)}</em>".</p>
            <p>Try asking about:</p>
            <ul>
                <li>A department (e.g. "Who works in Engineering?")</li>
                <li>A city (e.g. "Show employees in Mumbai")</li>
                <li>A person (e.g. "Tell me about Rahul")</li>
                <li>Salaries (e.g. "What is the average salary?")</li>
                <li>Statistics (e.g. "How many employees are there?")</li>
            </ul>`;
}

// ===================== HELPERS =====================
function findColumnMatch(query, colName) {
    const values = [...new Set(excelData.map(r => r[colName]).filter(Boolean))];
    const q = query.toLowerCase();
    for (const val of values) {
        if (q.includes(val.toLowerCase())) {
            return val;
        }
    }
    return null;
}

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
