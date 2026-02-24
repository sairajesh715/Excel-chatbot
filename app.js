/* ========================================
   DataBot – Smart AI-like Chatbot Engine
   Excel parsing + intelligent local NLP
   No API key needed!
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
        document.getElementById('rowCount').textContent = '⚠ Could not load data file.';
    }
}

// ===================== TABLE RENDERING =====================
function renderTable(data) {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    thead.innerHTML = '<tr>' + columns.map(col => `<th>${esc(col)}</th>`).join('') + '</tr>';
    tbody.innerHTML = data.map(row =>
        '<tr>' + columns.map(col => `<td>${esc(String(row[col] ?? ''))}</td>`).join('') + '</tr>'
    ).join('');
}

function initTableSearch() {
    document.getElementById('tableSearch').addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase().trim();
        if (!query) { renderTable(excelData); return; }
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
    const depts = unique('Department');
    if (depts.length) stats.push({ label: 'Departments', value: depts.length, detail: depts.join(', ') });
    const salaries = numericCol('Salary');
    if (salaries.length) {
        const avg = salaries.reduce((a, b) => a + b, 0) / salaries.length;
        stats.push({ label: 'Avg Salary', value: '$' + Math.round(avg).toLocaleString(), detail: `Range: $${Math.min(...salaries).toLocaleString()} – $${Math.max(...salaries).toLocaleString()}` });
    }
    const cities = unique('City');
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
    const typingEl = showTyping();
    setTimeout(() => {
        typingEl.remove();
        const response = processQuery(text);
        addMessage(response, 'bot');
    }, 400 + Math.random() * 400);
}

function askSuggestion(btn) {
    document.getElementById('chatInput').value = btn.textContent.trim();
    sendMessage();
}

// ===================== SMART QUERY PROCESSOR =====================
function processQuery(query) {
    if (excelData.length === 0) return p('⚠ No data loaded.');

    const q = query.toLowerCase().trim();
    const words = q.replace(/[?.,!'"]/g, '').split(/\s+/);

    // === GREETINGS ===
    if (/^(hi|hello|hey|good\s*(morning|afternoon|evening)|greetings|howdy)\b/.test(q)) {
        return p(`Hello! 👋 I'm your AI assistant for employee data.`) +
            p(`We have <strong>${excelData.length} employees</strong> across <strong>${unique('Department').length} departments</strong> in <strong>${unique('City').length} cities</strong>. Ask me anything!`);
    }

    // === HELP ===
    if (/^(help|what can you|how to|commands|features)/.test(q)) {
        return p('🤖 Here are things you can ask me:') +
            '<ul>' +
            '<li>How many employees are there?</li>' +
            '<li>What is the total/average salary?</li>' +
            '<li>Who has the highest/lowest salary?</li>' +
            '<li>Show employees in [city/department]</li>' +
            '<li>Tell me about [person name]</li>' +
            '<li>Compare departments by salary</li>' +
            '<li>Who joined in [year]?</li>' +
            '<li>List all departments/cities/roles</li>' +
            '<li>Who earns more than $80,000?</li>' +
            '<li>Show active/inactive employees</li>' +
            '</ul>';
    }

    // === TOTAL SALARY ===
    if (match(q, ['total salary', 'sum salary', 'sum of salary', 'total of salary', 'combined salary', 'salary total', 'total salaries', 'sum of all salaries', 'total pay', 'payroll'])) {
        const salaries = numericCol('Salary');
        if (salaries.length) {
            const total = salaries.reduce((a, b) => a + b, 0);
            return p(`💰 The <strong>total salary</strong> of all ${salaries.length} employees is <strong>$${total.toLocaleString()}</strong>`);
        }
        return p('Could not calculate total salary.');
    }

    // === AVERAGE SALARY (with optional filter) ===
    if (match(q, ['average salary', 'avg salary', 'mean salary'])) {
        // Check for department/city filter
        const dept = findValue(q, 'Department');
        const city = findValue(q, 'City');
        let subset = excelData;
        let label = '';
        if (dept) { subset = filterBy('Department', dept); label = ` in <strong>${dept}</strong>`; }
        else if (city) { subset = filterBy('City', city); label = ` in <strong>${city}</strong>`; }

        const salaries = subset.map(r => Number(r['Salary'])).filter(s => !isNaN(s));
        if (salaries.length) {
            const avg = Math.round(salaries.reduce((a, b) => a + b, 0) / salaries.length);
            return p(`💰 The average salary${label} is <strong>$${avg.toLocaleString()}</strong>`) +
                p(`Range: $${Math.min(...salaries).toLocaleString()} – $${Math.max(...salaries).toLocaleString()} (${salaries.length} employees)`);
        }
        return p('Could not calculate average salary.');
    }

    // === COUNT / HOW MANY ===
    if (match(q, ['how many', 'total employees', 'total count', 'total records', 'total number', 'count of', 'number of employees', 'employee count', 'headcount'])) {
        const dept = findValue(q, 'Department');
        const city = findValue(q, 'City');
        const role = findValue(q, 'Role');
        if (dept) {
            const count = filterBy('Department', dept).length;
            return p(`🏢 There are <strong>${count} employees</strong> in the <strong>${dept}</strong> department.`);
        }
        if (city) {
            const count = filterBy('City', city).length;
            return p(`📍 There are <strong>${count} employees</strong> in <strong>${city}</strong>.`);
        }
        if (role) {
            const count = excelData.filter(r => String(r['Role'] || '').toLowerCase().includes(role.toLowerCase())).length;
            return p(`💼 There are <strong>${count} employees</strong> with role matching "<strong>${role}</strong>".`);
        }
        return p(`👥 There are <strong>${excelData.length} employees</strong> in the dataset.`);
    }

    // === HIGHEST / TOP SALARY ===
    if (match(q, ['highest salary', 'max salary', 'top salary', 'most paid', 'highest paid', 'top earner', 'maximum salary', 'best paid', 'highest earning', 'who earns the most', 'who makes the most'])) {
        const sorted = [...excelData].sort((a, b) => num(b, 'Salary') - num(a, 'Salary'));
        const top = sorted.slice(0, 5);
        return p('🏆 Top earners:') + miniTable(top, ['Name', 'Department', 'Role', 'Salary', 'City']);
    }

    // === LOWEST SALARY ===
    if (match(q, ['lowest salary', 'min salary', 'least paid', 'lowest paid', 'minimum salary', 'least earning', 'who earns the least', 'who makes the least', 'lowest earning'])) {
        const sorted = [...excelData].sort((a, b) => num(a, 'Salary') - num(b, 'Salary'));
        const bottom = sorted.slice(0, 5);
        return p('📉 Lowest salaries:') + miniTable(bottom, ['Name', 'Department', 'Role', 'Salary', 'City']);
    }

    // === SALARY RANGE / ABOVE / BELOW ===
    const aboveMatch = q.match(/(?:earn|salary|make|paid).*(?:more than|above|over|greater than|exceeds?|>\s*)\$?([\d,]+)/);
    const belowMatch = q.match(/(?:earn|salary|make|paid).*(?:less than|below|under|<\s*)\$?([\d,]+)/);
    if (aboveMatch) {
        const threshold = parseInt(aboveMatch[1].replace(/,/g, ''));
        const filtered = excelData.filter(r => num(r, 'Salary') > threshold);
        if (filtered.length) return p(`💰 <strong>${filtered.length} employees</strong> earn more than <strong>$${threshold.toLocaleString()}</strong>:`) + miniTable(filtered, ['Name', 'Department', 'Role', 'Salary']);
        return p(`No employees earn more than $${threshold.toLocaleString()}.`);
    }
    if (belowMatch) {
        const threshold = parseInt(belowMatch[1].replace(/,/g, ''));
        const filtered = excelData.filter(r => num(r, 'Salary') < threshold);
        if (filtered.length) return p(`💰 <strong>${filtered.length} employees</strong> earn less than <strong>$${threshold.toLocaleString()}</strong>:`) + miniTable(filtered, ['Name', 'Department', 'Role', 'Salary']);
        return p(`No employees earn less than $${threshold.toLocaleString()}.`);
    }

    // === SALARY BETWEEN ===
    const betweenMatch = q.match(/(?:salary|earn|paid).*(?:between)\s*\$?([\d,]+)\s*(?:and|to|-)\s*\$?([\d,]+)/);
    if (betweenMatch) {
        const low = parseInt(betweenMatch[1].replace(/,/g, ''));
        const high = parseInt(betweenMatch[2].replace(/,/g, ''));
        const filtered = excelData.filter(r => { const s = num(r, 'Salary'); return s >= low && s <= high; });
        if (filtered.length) return p(`💰 <strong>${filtered.length} employees</strong> earn between <strong>$${low.toLocaleString()}</strong> and <strong>$${high.toLocaleString()}</strong>:`) + miniTable(filtered, ['Name', 'Department', 'Role', 'Salary']);
        return p(`No employees earn between $${low.toLocaleString()} and $${high.toLocaleString()}.`);
    }

    // === MOST / LEAST EXPERIENCE ===
    if (match(q, ['most experience', 'highest experience', 'most senior', 'most experienced', 'longest serving', 'longest tenure'])) {
        const expCol = findExpCol();
        if (expCol) {
            const sorted = [...excelData].sort((a, b) => num(b, expCol) - num(a, expCol));
            return p('🎖 Most experienced employees:') + miniTable(sorted.slice(0, 5), ['Name', 'Department', 'Role', expCol, 'City']);
        }
    }
    if (match(q, ['least experience', 'lowest experience', 'newest', 'most junior', 'least experienced', 'shortest tenure'])) {
        const expCol = findExpCol();
        if (expCol) {
            const sorted = [...excelData].sort((a, b) => num(a, expCol) - num(b, expCol));
            return p('🆕 Least experienced employees:') + miniTable(sorted.slice(0, 5), ['Name', 'Department', 'Role', expCol, 'City']);
        }
    }

    // === LIST DEPARTMENTS ===
    if (match(q, ['list department', 'all department', 'what department', 'show department', 'which department', 'departments'])) {
        const depts = unique('Department');
        const rows = depts.map(d => {
            const emps = filterBy('Department', d);
            const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
            const avg = sals.length ? Math.round(sals.reduce((a, b) => a + b, 0) / sals.length) : 0;
            return `<li><strong>${d}</strong>: ${emps.length} employee${emps.length > 1 ? 's' : ''} (avg salary: $${avg.toLocaleString()})</li>`;
        });
        return p(`📋 Departments (${depts.length}):`) + '<ul>' + rows.join('') + '</ul>';
    }

    // === LIST CITIES ===
    if (match(q, ['list city', 'list cities', 'all city', 'all cities', 'what city', 'what cities', 'show city', 'show cities', 'which city', 'which cities', 'locations', 'list location', 'all location'])) {
        const cities = unique('City');
        const rows = cities.map(c => {
            const count = filterBy('City', c).length;
            return `<li><strong>${c}</strong>: ${count} employee${count > 1 ? 's' : ''}</li>`;
        });
        return p(`📍 Locations (${cities.length}):`) + '<ul>' + rows.join('') + '</ul>';
    }

    // === LIST ROLES ===
    if (match(q, ['list role', 'all role', 'what role', 'show role', 'which role', 'roles', 'list position', 'all position', 'job title'])) {
        const roles = unique('Role');
        const rows = roles.map(r => {
            const count = excelData.filter(e => e['Role'] === r).length;
            return `<li><strong>${r}</strong>: ${count} employee${count > 1 ? 's' : ''}</li>`;
        });
        return p(`💼 Roles (${roles.length}):`) + '<ul>' + rows.join('') + '</ul>';
    }

    // === COMPARE DEPARTMENTS ===
    if (match(q, ['compare department', 'department comparison', 'department wise', 'department breakdown', 'by department', 'per department', 'each department', 'department summary', 'department stats'])) {
        const depts = unique('Department');
        const data = depts.map(d => {
            const emps = filterBy('Department', d);
            const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
            const avg = sals.length ? Math.round(sals.reduce((a, b) => a + b, 0) / sals.length) : 0;
            const total = sals.reduce((a, b) => a + b, 0);
            return { Department: d, Employees: emps.length, 'Avg Salary': '$' + avg.toLocaleString(), 'Total Salary': '$' + total.toLocaleString() };
        });
        return p('📊 Department Comparison:') + objTable(data);
    }

    // === COMPARE CITIES ===
    if (match(q, ['compare city', 'compare cities', 'city comparison', 'city wise', 'city breakdown', 'by city', 'per city', 'each city', 'city summary', 'location comparison'])) {
        const cities = unique('City');
        const data = cities.map(c => {
            const emps = filterBy('City', c);
            const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
            const avg = sals.length ? Math.round(sals.reduce((a, b) => a + b, 0) / sals.length) : 0;
            return { City: c, Employees: emps.length, 'Avg Salary': '$' + avg.toLocaleString() };
        });
        return p('📊 City Comparison:') + objTable(data);
    }

    // === JOIN DATE / YEAR ===
    const yearMatch = q.match(/(?:join|joined|hired|started|recruited).*?(20\d{2})/);
    if (yearMatch) {
        const year = yearMatch[1];
        const dateCol = columns.find(c => c.toLowerCase().includes('join') || c.toLowerCase().includes('date'));
        if (dateCol) {
            const filtered = excelData.filter(r => String(r[dateCol] || '').includes(year));
            if (filtered.length) return p(`📅 Employees who joined in <strong>${year}</strong> (${filtered.length}):`) + miniTable(filtered, ['Name', 'Department', 'Role', dateCol]);
            return p(`No employees found who joined in ${year}.`);
        }
    }
    if (match(q, ['recent hire', 'recently joined', 'newest employee', 'latest hire', 'last joined', 'most recent'])) {
        const dateCol = columns.find(c => c.toLowerCase().includes('join') || c.toLowerCase().includes('date'));
        if (dateCol) {
            const sorted = [...excelData].sort((a, b) => String(b[dateCol] || '').localeCompare(String(a[dateCol] || '')));
            return p('📅 Most recent hires:') + miniTable(sorted.slice(0, 5), ['Name', 'Department', 'Role', dateCol]);
        }
    }

    // === STATUS FILTER ===
    if (match(q, ['active employee', 'active staff', 'who is active', 'show active', 'currently active'])) {
        const filtered = excelData.filter(r => String(r['Status'] || '').toLowerCase() === 'active');
        return p(`✅ <strong>${filtered.length} active</strong> employees:`) + miniTable(filtered, ['Name', 'Department', 'Role', 'City']);
    }
    if (match(q, ['inactive', 'not active', 'left', 'resigned', 'terminated'])) {
        const filtered = excelData.filter(r => String(r['Status'] || '').toLowerCase() !== 'active');
        if (filtered.length) return p(`⛔ <strong>${filtered.length} inactive</strong> employees:`) + miniTable(filtered, ['Name', 'Department', 'Role', 'Status']);
        return p('All employees are currently active! ✅');
    }

    // === SPECIFIC PERSON LOOKUP ===
    const namePatterns = [
        /(?:who is|about|info|details|tell me about|find|show me|profile|lookup)\s+(.+)/,
        /(.+?)(?:'s|'s)\s+(?:salary|detail|info|department|role|email|city|experience)/,
        /(?:salary of|salary for|how much does|how much is)\s+(.+?)(?:\s+(?:earn|make|get))?$/,
    ];
    for (const pattern of namePatterns) {
        const m = q.match(pattern);
        if (m) {
            const searchName = m[1].replace(/[?.,!']/g, '').trim();
            if (searchName.length >= 2) {
                const matches = excelData.filter(r =>
                    String(r['Name'] || '').toLowerCase().includes(searchName)
                );
                if (matches.length > 0) {
                    // Show full details for single person
                    if (matches.length === 1) {
                        const person = matches[0];
                        let details = '<div style="margin:8px 0">';
                        columns.forEach(col => {
                            let val = person[col] ?? '';
                            if (col === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
                            details += `<p style="margin:2px 0"><strong>${esc(col)}:</strong> ${esc(String(val))}</p>`;
                        });
                        details += '</div>';
                        return p(`👤 <strong>${esc(person['Name'])}</strong>`) + details;
                    }
                    return p(`👤 Found ${matches.length} matches:`) + miniTable(matches, columns.slice(0, 7));
                }
            }
        }
    }

    // === FILTER BY DEPARTMENT ===
    const deptVal = findValue(q, 'Department');
    if (deptVal) {
        const filtered = filterBy('Department', deptVal);
        if (filtered.length) return p(`🏢 <strong>${deptVal}</strong> department (${filtered.length} employee${filtered.length > 1 ? 's' : ''}):`) + miniTable(filtered, ['Name', 'Role', 'Salary', 'City', 'Experience (Years)']);
    }

    // === FILTER BY CITY ===
    const cityVal = findValue(q, 'City');
    if (cityVal) {
        const filtered = filterBy('City', cityVal);
        if (filtered.length) return p(`📍 Employees in <strong>${cityVal}</strong> (${filtered.length}):`) + miniTable(filtered, ['Name', 'Department', 'Role', 'Salary']);
    }

    // === FILTER BY ROLE ===
    const roleVal = findValue(q, 'Role');
    if (roleVal) {
        const filtered = excelData.filter(r => String(r['Role'] || '').toLowerCase().includes(roleVal.toLowerCase()));
        if (filtered.length) return p(`💼 Employees with role "<strong>${roleVal}</strong>" (${filtered.length}):`) + miniTable(filtered, ['Name', 'Department', 'Salary', 'City']);
    }

    // === SHOW ALL ===
    if (match(q, ['show all', 'list all', 'all employees', 'everyone', 'all data', 'all records', 'show everything', 'entire list', 'full list', 'complete list'])) {
        return p(`📋 All <strong>${excelData.length} employees</strong>:`) + miniTable(excelData, ['Name', 'Department', 'Role', 'Salary', 'City']);
    }

    // === SUMMARY / OVERVIEW ===
    if (match(q, ['summary', 'overview', 'statistics', 'stats', 'report', 'dashboard', 'insights', 'tell me about the data', 'data summary', 'data overview'])) {
        const sals = numericCol('Salary');
        const total = sals.reduce((a, b) => a + b, 0);
        const avg = Math.round(total / sals.length);
        const depts = unique('Department');
        const cities = unique('City');
        const expCol = findExpCol();
        let avgExp = '';
        if (expCol) {
            const exps = excelData.map(r => Number(r[expCol])).filter(e => !isNaN(e));
            if (exps.length) avgExp = `<li>📈 Average experience: <strong>${(exps.reduce((a, b) => a + b, 0) / exps.length).toFixed(1)} years</strong></li>`;
        }
        return p('📊 <strong>Data Summary</strong>') +
            '<ul>' +
            `<li>👥 Total employees: <strong>${excelData.length}</strong></li>` +
            `<li>🏢 Departments: <strong>${depts.join(', ')}</strong> (${depts.length})</li>` +
            `<li>📍 Cities: <strong>${cities.join(', ')}</strong> (${cities.length})</li>` +
            `<li>💰 Total payroll: <strong>$${total.toLocaleString()}</strong></li>` +
            `<li>💰 Average salary: <strong>$${avg.toLocaleString()}</strong></li>` +
            `<li>💰 Salary range: <strong>$${Math.min(...sals).toLocaleString()} – $${Math.max(...sals).toLocaleString()}</strong></li>` +
            avgExp +
            '</ul>';
    }

    // === THANK YOU / BYE ===
    if (match(q, ['thank', 'thanks', 'bye', 'goodbye', 'see you', 'that\'s all', 'thats all'])) {
        return p('You\'re welcome! 😊 Feel free to ask me anything else about the data anytime.');
    }

    // === GENERIC KEYWORD SEARCH (SMART FALLBACK) ===
    const keywords = q.replace(/[?.,!'"]/g, '').split(/\s+/).filter(w => w.length > 2 && !['the', 'and', 'for', 'are', 'was', 'what', 'who', 'how', 'can', 'you', 'show', 'tell', 'give', 'get', 'from', 'with', 'this', 'that', 'have', 'has', 'does', 'will'].includes(w));

    if (keywords.length > 0) {
        const matches = excelData.filter(row =>
            keywords.some(kw => columns.some(col => String(row[col] || '').toLowerCase().includes(kw)))
        );
        if (matches.length > 0) {
            const displayCols = matches.length <= 5 ? columns.slice(0, 7) : ['Name', 'Department', 'Role', 'Salary', 'City'];
            return p(`🔍 Found <strong>${matches.length} result${matches.length > 1 ? 's' : ''}</strong> matching your query:`) +
                miniTable(matches.slice(0, 15), displayCols);
        }
    }

    // === NO MATCH ===
    return p(`🤔 I couldn't find results for "<em>${esc(query)}</em>".`) +
        p('Try asking about:') +
        '<ul>' +
        '<li>"What is the total salary?"</li>' +
        '<li>"Show employees in Engineering"</li>' +
        '<li>"Who has the highest salary?"</li>' +
        '<li>"Tell me about Guru Charan"</li>' +
        '<li>"Compare departments"</li>' +
        '<li>"Summary" for a data overview</li>' +
        '</ul>';
}

// ===================== UI HELPERS =====================
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

// ===================== DATA HELPERS =====================
function unique(col) {
    return [...new Set(excelData.map(r => r[col]).filter(Boolean))];
}

function numericCol(col) {
    return excelData.map(r => Number(r[col])).filter(s => !isNaN(s) && s > 0);
}

function num(row, col) {
    return Number(row[col] || 0);
}

function filterBy(col, val) {
    return excelData.filter(r => String(r[col] || '').toLowerCase() === val.toLowerCase());
}

function findValue(query, colName) {
    const vals = unique(colName);
    const q = query.toLowerCase();
    for (const val of vals) {
        if (q.includes(val.toLowerCase())) return val;
    }
    return null;
}

function findExpCol() {
    return columns.find(c => c.toLowerCase().includes('experience'));
}

function match(q, patterns) {
    return patterns.some(p => q.includes(p));
}

function p(text) {
    return `<p>${text}</p>`;
}

function esc(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function miniTable(data, cols) {
    const availCols = cols.filter(c => columns.includes(c));
    if (availCols.length === 0) return p('No matching columns found.');
    let html = '<table class="chat-result-table"><thead><tr>';
    html += availCols.map(c => `<th>${esc(c)}</th>`).join('');
    html += '</tr></thead><tbody>';
    data.forEach(row => {
        html += '<tr>';
        html += availCols.map(c => {
            let val = row[c] ?? '';
            if (c === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
            return `<td>${esc(String(val))}</td>`;
        }).join('');
        html += '</tr>';
    });
    html += '</tbody></table>';
    return html;
}

function objTable(data) {
    if (!data.length) return '';
    const keys = Object.keys(data[0]);
    let html = '<table class="chat-result-table"><thead><tr>';
    html += keys.map(k => `<th>${esc(k)}</th>`).join('');
    html += '</tr></thead><tbody>';
    data.forEach(row => {
        html += '<tr>' + keys.map(k => `<td>${esc(String(row[k]))}</td>`).join('') + '</tr>';
    });
    html += '</tbody></table>';
    return html;
}
