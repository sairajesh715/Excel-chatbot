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
const APP_VERSION = 'v5.0';
console.log('DataBot ' + APP_VERSION + ' loaded');

function processQuery(query) {
    if (excelData.length === 0) return p('⚠ No data loaded.');

    const q = query.toLowerCase().trim();
    console.log('[DataBot] Query:', q, '| Data rows:', excelData.length);
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
            '<li><strong>Math:</strong> "Sum of salary of Guru Charan and Sai Rajesh"</li>' +
            '<li><strong>Compare:</strong> "Compare Guru Charan vs Guru Dayal"</li>' +
            '<li><strong>Total:</strong> "What is the total salary?"</li>' +
            '<li><strong>Average:</strong> "Average salary in Engineering"</li>' +
            '<li><strong>Salary range:</strong> "Who earns more than $80,000?"</li>' +
            '<li><strong>Person info:</strong> "Tell me about Guru Charan"</li>' +
            '<li><strong>Filter:</strong> "Show employees in Hyderabad"</li>' +
            '<li><strong>Compare teams:</strong> "Compare departments"</li>' +
            '<li><strong>Rankings:</strong> "Who has the highest salary?"</li>' +
            '<li><strong>Overview:</strong> "Summary"</li>' +
            '</ul>';
    }

    // === PERSON-SPECIFIC MATH (sum of salary of X and Y) ===
    // Detect names mentioned in the query
    const mentionedPeople = findMentionedPeople(q);

    if (mentionedPeople.length >= 2 && match(q, ['sum', 'total', 'combined', 'add', 'together', 'plus', 'both', 'salary', 'earn', 'pay', 'income'])) {
        const salaries = mentionedPeople.map(p => ({ name: p['Name'], salary: num(p, 'Salary') }));
        const total = salaries.reduce((a, b) => a + b.salary, 0);
        let html = p(`💰 Combined salary of <strong>${mentionedPeople.length} employees</strong>:`);
        html += '<ul>' + salaries.map(s => `<li><strong>${esc(s.name)}</strong>: $${s.salary.toLocaleString()}</li>`).join('') + '</ul>';
        html += p(`<strong>Total: $${total.toLocaleString()}</strong>`);
        return html;
    }

    // Compare two people
    if (mentionedPeople.length >= 2 && match(q, ['compare', 'difference', 'diff', 'vs', 'versus', 'more than', 'less than', 'between', 'who earns more', 'who makes more', 'higher', 'lower'])) {
        const p1 = mentionedPeople[0], p2 = mentionedPeople[1];
        const s1 = num(p1, 'Salary'), s2 = num(p2, 'Salary');
        const diff = Math.abs(s1 - s2);
        const higher = s1 >= s2 ? p1 : p2;
        const lower = s1 < s2 ? p1 : p2;
        let html = p(`📊 Salary comparison:`);
        html += '<ul>';
        html += `<li><strong>${esc(p1['Name'])}</strong>: $${s1.toLocaleString()} (${p1['Role']})</li>`;
        html += `<li><strong>${esc(p2['Name'])}</strong>: $${s2.toLocaleString()} (${p2['Role']})</li>`;
        html += '</ul>';
        if (s1 === s2) {
            html += p(`They both earn the same salary! 🤝`);
        } else {
            html += p(`💡 <strong>${esc(higher['Name'])}</strong> earns <strong>$${diff.toLocaleString()}</strong> more than <strong>${esc(lower['Name'])}</strong>`);
            html += p(`That's a <strong>${((diff / Math.min(s1, s2)) * 100).toFixed(1)}%</strong> difference.`);
        }
        return html;
    }

    // Average salary of mentioned people
    if (mentionedPeople.length >= 2 && match(q, ['average', 'avg', 'mean'])) {
        const salaries = mentionedPeople.map(p => ({ name: p['Name'], salary: num(p, 'Salary') }));
        const avg = Math.round(salaries.reduce((a, b) => a + b.salary, 0) / salaries.length);
        let html = p(`💰 Average salary of <strong>${mentionedPeople.length} employees</strong>:`);
        html += '<ul>' + salaries.map(s => `<li><strong>${esc(s.name)}</strong>: $${s.salary.toLocaleString()}</li>`).join('') + '</ul>';
        html += p(`<strong>Average: $${avg.toLocaleString()}</strong>`);
        return html;
    }

    // Min/Max of mentioned people
    if (mentionedPeople.length >= 2 && match(q, ['highest', 'max', 'maximum', 'most', 'top', 'best'])) {
        const sorted = [...mentionedPeople].sort((a, b) => num(b, 'Salary') - num(a, 'Salary'));
        return p(`🏆 Among the mentioned employees, <strong>${esc(sorted[0]['Name'])}</strong> earns the most at <strong>$${num(sorted[0], 'Salary').toLocaleString()}</strong>`) +
            miniTable(sorted, ['Name', 'Department', 'Role', 'Salary']);
    }
    if (mentionedPeople.length >= 2 && match(q, ['lowest', 'min', 'minimum', 'least'])) {
        const sorted = [...mentionedPeople].sort((a, b) => num(a, 'Salary') - num(b, 'Salary'));
        return p(`📉 Among the mentioned employees, <strong>${esc(sorted[0]['Name'])}</strong> earns the least at <strong>$${num(sorted[0], 'Salary').toLocaleString()}</strong>`) +
            miniTable(sorted, ['Name', 'Department', 'Role', 'Salary']);
    }

    // Single person salary query
    if (mentionedPeople.length === 1) {
        const person = mentionedPeople[0];
        const expCol = findExpCol();
        if (match(q, ['salary', 'earn', 'make', 'pay', 'paid', 'compensation', 'income'])) {
            return p(`💰 <strong>${esc(person['Name'])}</strong>'s salary is <strong>$${num(person, 'Salary').toLocaleString()}</strong>`) +
                p(`Role: ${person['Role']} | Department: ${person['Department']} | City: ${person['City']}`);
        }
        if (match(q, ['experience', 'years', 'tenure', 'how long'])) {
            if (expCol) {
                return p(`📈 <strong>${esc(person['Name'])}</strong> has <strong>${person[expCol]} years</strong> of experience`) +
                    p(`Role: ${person['Role']} | Department: ${person['Department']}`);
            }
        }
        if (match(q, ['department', 'team', 'which department', 'what department'])) {
            return p(`🏢 <strong>${esc(person['Name'])}</strong> works in the <strong>${person['Department']}</strong> department as a <strong>${person['Role']}</strong>`);
        }
        if (match(q, ['email', 'mail', 'contact'])) {
            return p(`📧 <strong>${esc(person['Name'])}</strong>'s email is <strong>${person['Email'] || 'N/A'}</strong>`);
        }
        if (match(q, ['city', 'location', 'where', 'based'])) {
            return p(`📍 <strong>${esc(person['Name'])}</strong> is based in <strong>${person['City']}</strong>`);
        }
    }

    // If 2+ people mentioned but no specific handler matched, show their details
    if (mentionedPeople.length >= 2) {
        const salaries = mentionedPeople.map(p => ({ name: p['Name'], salary: num(p, 'Salary') }));
        const total = salaries.reduce((a, b) => a + b.salary, 0);
        const avg = Math.round(total / salaries.length);
        let html = p(`👥 Found <strong>${mentionedPeople.length} employees</strong> mentioned:`);
        html += miniTable(mentionedPeople, ['Name', 'Department', 'Role', 'Salary', 'City']);
        html += p(`💰 Combined salary: <strong>$${total.toLocaleString()}</strong> | Average: <strong>$${avg.toLocaleString()}</strong>`);
        return html;
    }

    // If single person mentioned but no specific handler matched, show profile
    if (mentionedPeople.length === 1) {
        const person = mentionedPeople[0];
        let details = '<div style="margin:8px 0">';
        columns.forEach(col => {
            let val = person[col] ?? '';
            if (col === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
            details += `<p style="margin:2px 0"><strong>${esc(col)}:</strong> ${esc(String(val))}</p>`;
        });
        details += '</div>';
        return p(`👤 <strong>${esc(person['Name'])}</strong>`) + details;
    }

    // === TOTAL SALARY (with optional filter) ===
    if (mentionedPeople.length === 0 && match(q, ['total salary', 'sum salary', 'sum of salary', 'total of salary', 'combined salary', 'salary total', 'total salaries', 'sum of all salaries', 'total pay', 'payroll'])) {
        const dept = findValue(q, 'Department');
        const city = findValue(q, 'City');
        let subset = excelData;
        let label = 'all employees';
        if (dept) { subset = filterBy('Department', dept); label = `<strong>${dept}</strong> department`; }
        else if (city) { subset = filterBy('City', city); label = `<strong>${city}</strong>`; }

        const salaries = subset.map(r => Number(r['Salary'])).filter(s => !isNaN(s));
        if (salaries.length) {
            const total = salaries.reduce((a, b) => a + b, 0);
            return p(`💰 The <strong>total salary</strong> of ${label} (${salaries.length} employees) is <strong>$${total.toLocaleString()}</strong>`);
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

    // ===================== SQL-LIKE QUERY ENGINE =====================

    // --- GROUP BY ---
    const groupByMatch = q.match(/(?:group\s*by|grouped\s*by|breakdown\s*by|per|by each|for each|split\s*by|categorize\s*by|category)\s+(department|city|role|status)/i);
    if (groupByMatch || match(q, ['group by', 'grouped by', 'breakdown by'])) {
        let groupCol = null;
        if (groupByMatch) {
            groupCol = resolveColName(groupByMatch[1]);
        } else {
            // Try to detect the column
            for (const col of ['Department', 'City', 'Role', 'Status']) {
                if (q.includes(col.toLowerCase())) { groupCol = col; break; }
            }
        }
        if (groupCol) {
            const groups = unique(groupCol);
            // Determine what aggregation
            const wantSum = match(q, ['sum', 'total']);
            const wantAvg = match(q, ['avg', 'average', 'mean']);
            const wantMax = match(q, ['max', 'highest', 'maximum', 'top']);
            const wantMin = match(q, ['min', 'lowest', 'minimum']);
            const wantCount = match(q, ['count', 'how many', 'number']);

            const data = groups.map(g => {
                const emps = filterBy(groupCol, g);
                const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
                const row = { [groupCol]: g, Count: emps.length };
                if (wantSum || (!wantAvg && !wantMax && !wantMin && !wantCount)) {
                    row['Total Salary'] = '$' + sals.reduce((a, b) => a + b, 0).toLocaleString();
                }
                if (wantAvg || (!wantSum && !wantMax && !wantMin && !wantCount)) {
                    row['Avg Salary'] = '$' + (sals.length ? Math.round(sals.reduce((a, b) => a + b, 0) / sals.length) : 0).toLocaleString();
                }
                if (wantMax) row['Max Salary'] = '$' + (sals.length ? Math.max(...sals) : 0).toLocaleString();
                if (wantMin) row['Min Salary'] = '$' + (sals.length ? Math.min(...sals) : 0).toLocaleString();
                return row;
            });

            // Sort if requested
            if (match(q, ['order by salary', 'sort by salary', 'order by avg', 'sort by avg', 'descending', 'desc'])) {
                data.sort((a, b) => parseInt((b['Avg Salary'] || b['Total Salary'] || '0').replace(/[$,]/g, '')) - parseInt((a['Avg Salary'] || a['Total Salary'] || '0').replace(/[$,]/g, '')));
            }

            return p(`📊 Group by <strong>${groupCol}</strong>:`) + objTable(data);
        }
    }

    // --- ORDER BY / SORT BY ---
    const orderByMatch = q.match(/(?:order\s*by|sort\s*by|sorted\s*by|arrange\s*by|rank\s*by)\s+(salary|name|experience|department|city|role|join\s*date|email)/i);
    if (orderByMatch) {
        let col = resolveColName(orderByMatch[1]);
        const isDesc = match(q, ['desc', 'descending', 'high to low', 'highest first', 'top', 'most']);
        const isAsc = match(q, ['asc', 'ascending', 'low to high', 'lowest first', 'least']);
        const descending = isDesc || (!isAsc && (col === 'Salary' || col === findExpCol()));

        let sorted;
        const numCols = ['Salary'];
        const expCol = findExpCol();
        if (expCol) numCols.push(expCol);

        if (numCols.includes(col)) {
            sorted = [...excelData].sort((a, b) => descending ? num(b, col) - num(a, col) : num(a, col) - num(b, col));
        } else {
            sorted = [...excelData].sort((a, b) => {
                const av = String(a[col] || ''), bv = String(b[col] || '');
                return descending ? bv.localeCompare(av) : av.localeCompare(bv);
            });
        }

        // Apply TOP N / LIMIT
        const limitMatch = q.match(/(?:top|first|limit|show)\s+(\d+)/);
        const limit = limitMatch ? parseInt(limitMatch[1]) : sorted.length;
        const result = sorted.slice(0, limit);
        const dir = descending ? 'descending ↓' : 'ascending ↑';

        return p(`📋 Ordered by <strong>${col}</strong> (${dir})${limit < sorted.length ? `, showing top ${limit}` : ''}:`) +
            miniTable(result, ['Name', 'Department', 'Role', 'Salary', 'City']);
    }

    // --- TOP N / BOTTOM N ---
    const topNMatch = q.match(/(?:top|first|best|highest)\s+(\d+)\s*(?:employees?|people|earners?|paid)?(?:\s+(?:by|in|from)\s+(\w+))?/);
    if (topNMatch) {
        const n = parseInt(topNMatch[1]);
        let subset = excelData;
        const filterCol = topNMatch[2] ? resolveColName(topNMatch[2]) : null;
        const filterVal = filterCol ? findValue(q, filterCol) : null;
        if (filterVal) subset = filterBy(filterCol, filterVal);

        const sorted = [...subset].sort((a, b) => num(b, 'Salary') - num(a, 'Salary'));
        return p(`🏆 Top ${n} earners${filterVal ? ' in <strong>' + filterVal + '</strong>' : ''}:`) +
            miniTable(sorted.slice(0, n), ['Name', 'Department', 'Role', 'Salary', 'City']);
    }
    const bottomNMatch = q.match(/(?:bottom|last|lowest|least)\s+(\d+)\s*(?:employees?|people|earners?|paid)?/);
    if (bottomNMatch) {
        const n = parseInt(bottomNMatch[1]);
        const sorted = [...excelData].sort((a, b) => num(a, 'Salary') - num(b, 'Salary'));
        return p(`📉 Bottom ${n} earners:`) +
            miniTable(sorted.slice(0, n), ['Name', 'Department', 'Role', 'Salary', 'City']);
    }

    // --- WHERE clause (multi-condition) ---
    const whereMatch = q.match(/(?:where|filter|show\s+(?:me\s+)?(?:employees?|people|records?)\s+(?:where|with|whose|having))\s+(.+)/i);
    if (whereMatch) {
        const conditions = whereMatch[1];
        let result = [...excelData];

        // Parse conditions: "salary > 80000 and department = engineering"
        const condParts = conditions.split(/\s+and\s+/i);
        for (const cond of condParts) {
            result = applyCondition(result, cond.trim());
        }

        if (result.length > 0) {
            return p(`🔍 Found <strong>${result.length} employees</strong> matching your conditions:`) +
                miniTable(result, ['Name', 'Department', 'Role', 'Salary', 'City']);
        }
        return p('No employees match those conditions.');
    }

    // --- DISTINCT ---
    if (match(q, ['distinct', 'unique', 'unique values', 'distinct values'])) {
        for (const col of columns) {
            if (q.includes(col.toLowerCase()) || q.includes(col.toLowerCase().replace(/[()]/g, ''))) {
                const vals = unique(col);
                return p(`📋 <strong>${vals.length}</strong> unique values in <strong>${col}</strong>:`) +
                    '<ul>' + vals.map(v => `<li>${esc(String(v))}</li>`).join('') + '</ul>';
            }
        }
        // Default: show all columns with their distinct counts
        const data = columns.map(c => ({
            Column: c,
            'Unique Values': unique(c).length,
            Sample: unique(c).slice(0, 3).join(', ')
        }));
        return p('📋 Distinct value counts per column:') + objTable(data);
    }

    // --- COUNT WHERE ---
    if (match(q, ['count where', 'how many where', 'count of', 'number where'])) {
        const afterCount = q.replace(/.*(?:count\s+where|how\s+many\s+where|count\s+of|number\s+where)\s*/i, '');
        if (afterCount) {
            const result = applyCondition([...excelData], afterCount);
            return p(`📊 <strong>${result.length}</strong> employees match your condition.`);
        }
    }

    // --- SUM/AVG/MIN/MAX of a column with optional WHERE ---
    const aggMatch = q.match(/(?:sum|total|average|avg|minimum|min|maximum|max)\s+(?:of\s+)?(?:the\s+)?(salary|experience)/i);
    if (aggMatch) {
        const aggType = q.match(/(sum|total|average|avg|minimum|min|maximum|max)/i)[1].toLowerCase();
        let col = resolveColName(aggMatch[1]);
        let subset = excelData;
        let label = '';

        // Check for WHERE-like filter
        const dept = findValue(q, 'Department');
        const city = findValue(q, 'City');
        const role = findValue(q, 'Role');
        if (dept) { subset = filterBy('Department', dept); label = ` in <strong>${dept}</strong>`; }
        else if (city) { subset = filterBy('City', city); label = ` in <strong>${city}</strong>`; }
        else if (role) { subset = subset.filter(r => String(r['Role'] || '').toLowerCase().includes(role.toLowerCase())); label = ` for <strong>${role}</strong>`; }

        const vals = subset.map(r => num(r, col)).filter(v => v > 0);
        if (vals.length === 0) return p('No numeric data found for that query.');

        let result, aggLabel;
        if (['sum', 'total'].includes(aggType)) {
            result = vals.reduce((a, b) => a + b, 0); aggLabel = 'Sum';
        } else if (['average', 'avg'].includes(aggType)) {
            result = Math.round(vals.reduce((a, b) => a + b, 0) / vals.length); aggLabel = 'Average';
        } else if (['minimum', 'min'].includes(aggType)) {
            result = Math.min(...vals); aggLabel = 'Minimum';
        } else {
            result = Math.max(...vals); aggLabel = 'Maximum';
        }

        const display = col === 'Salary' ? '$' + result.toLocaleString() : result.toLocaleString();
        return p(`📊 <strong>${aggLabel}</strong> of ${col}${label}: <strong>${display}</strong> (${vals.length} employees)`);
    }

    // --- HAVING (group by + condition) ---
    if (match(q, ['having', 'departments with more than', 'cities with more than', 'departments having', 'cities having'])) {
        const havingNum = q.match(/(?:more than|greater than|over|above|at least|>=?)\s*(\d+)/);
        if (havingNum) {
            const threshold = parseInt(havingNum[1]);
            const isAboutSalary = match(q, ['salary', 'earn', 'pay']);

            for (const col of ['Department', 'City']) {
                if (q.includes(col.toLowerCase()) || q.includes(col.toLowerCase() + 's')) {
                    const groups = unique(col);
                    let filtered;
                    if (isAboutSalary) {
                        filtered = groups.filter(g => {
                            const emps = filterBy(col, g);
                            const avg = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
                            return avg.length && (avg.reduce((a, b) => a + b, 0) / avg.length) > threshold;
                        });
                        const data = filtered.map(g => {
                            const emps = filterBy(col, g);
                            const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
                            return { [col]: g, Employees: emps.length, 'Avg Salary': '$' + Math.round(sals.reduce((a, b) => a + b, 0) / sals.length).toLocaleString() };
                        });
                        return p(`📊 ${col}s with average salary above $${threshold.toLocaleString()}:`) + objTable(data);
                    } else {
                        filtered = groups.filter(g => filterBy(col, g).length > threshold);
                        const data = filtered.map(g => ({ [col]: g, Employees: filterBy(col, g).length }));
                        return p(`📊 ${col}s with more than ${threshold} employees:`) + objTable(data);
                    }
                }
            }
        }
    }

    // ===================== SMART AUTO-DETECT (bare input) =====================
    // If user just types a name, city, department, role, ID, email, year, etc.

    // --- EMPLOYEE ID MATCH (e.g. "EMP001") ---
    const idMatch = q.match(/^emp\s*0*(\d+)$/i);
    if (idMatch) {
        const emp = excelData.find(r => String(r['Employee ID'] || '').toLowerCase().replace(/\s/g, '') === q.replace(/\s/g, ''));
        if (emp) {
            let details = '<div style="margin:8px 0">';
            columns.forEach(col => {
                let val = emp[col] ?? '';
                if (col === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
                details += `<p style="margin:2px 0"><strong>${esc(col)}:</strong> ${esc(String(val))}</p>`;
            });
            details += '</div>';
            return p(`👤 <strong>${esc(emp['Name'])}</strong>`) + details;
        }
    }

    // --- EMAIL MATCH (e.g. "guru.charan@company.com") ---
    if (q.includes('@')) {
        const emp = excelData.find(r => String(r['Email'] || '').toLowerCase().includes(q.trim()));
        if (emp) {
            let details = '<div style="margin:8px 0">';
            columns.forEach(col => {
                let val = emp[col] ?? '';
                if (col === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
                details += `<p style="margin:2px 0"><strong>${esc(col)}:</strong> ${esc(String(val))}</p>`;
            });
            details += '</div>';
            return p(`👤 <strong>${esc(emp['Name'])}</strong>`) + details;
        }
    }

    // --- EXACT DEPARTMENT MATCH (e.g. just "Engineering" or "HR") ---
    const exactDept = unique('Department').find(d => d.toLowerCase() === q.trim());
    if (exactDept) {
        const emps = filterBy('Department', exactDept);
        const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
        const total = sals.reduce((a, b) => a + b, 0);
        const avg = sals.length ? Math.round(total / sals.length) : 0;
        return p(`🏢 <strong>${exactDept}</strong> Department (${emps.length} employees)`) +
            p(`💰 Total salary: <strong>$${total.toLocaleString()}</strong> | Average: <strong>$${avg.toLocaleString()}</strong>`) +
            miniTable(emps, ['Name', 'Role', 'Salary', 'City', 'Experience (Years)']);
    }

    // --- EXACT CITY MATCH (e.g. just "Hyderabad") ---
    const exactCity = unique('City').find(c => c.toLowerCase() === q.trim());
    if (exactCity) {
        const emps = filterBy('City', exactCity);
        const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
        const total = sals.reduce((a, b) => a + b, 0);
        const avg = sals.length ? Math.round(total / sals.length) : 0;
        return p(`📍 Employees in <strong>${exactCity}</strong> (${emps.length})`) +
            p(`💰 Total salary: <strong>$${total.toLocaleString()}</strong> | Average: <strong>$${avg.toLocaleString()}</strong>`) +
            miniTable(emps, ['Name', 'Department', 'Role', 'Salary', 'Experience (Years)']);
    }

    // --- EXACT ROLE MATCH (e.g. just "Developer" or "HR Manager") ---
    const exactRole = unique('Role').find(r => r.toLowerCase() === q.trim() || q.trim().includes(r.toLowerCase()));
    if (exactRole) {
        const emps = excelData.filter(r => r['Role'] === exactRole);
        return p(`💼 Employees with role "<strong>${exactRole}</strong>" (${emps.length})`) +
            miniTable(emps, ['Name', 'Department', 'Salary', 'City', 'Experience (Years)']);
    }

    // --- YEAR MATCH (e.g. just "2023" or "2021") ---
    const yearOnly = q.match(/^(20\d{2})$/);
    if (yearOnly) {
        const year = yearOnly[1];
        const dateCol = columns.find(c => c.toLowerCase().includes('join') || c.toLowerCase().includes('date'));
        if (dateCol) {
            const emps = excelData.filter(r => String(r[dateCol] || '').includes(year));
            if (emps.length > 0) {
                return p(`📅 Employees who joined in <strong>${year}</strong> (${emps.length}):`) +
                    miniTable(emps, ['Name', 'Department', 'Role', 'Salary', dateCol]);
            }
            return p(`No employees found who joined in ${year}.`);
        }
    }

    // --- PURE NUMBER (e.g. just "80000" → who earns around that?) ---
    const pureNum = q.match(/^\$?([\d,]+)$/);
    if (pureNum) {
        const target = parseInt(pureNum[1].replace(/,/g, ''));
        if (target > 1000) {
            // Salary range: show employees earning ±20% of this amount
            const margin = target * 0.2;
            const emps = excelData.filter(r => {
                const s = num(r, 'Salary');
                return s >= (target - margin) && s <= (target + margin);
            });
            const exact = excelData.filter(r => num(r, 'Salary') === target);
            if (exact.length > 0) {
                return p(`💰 <strong>${exact.length}</strong> employee(s) earning exactly <strong>$${target.toLocaleString()}</strong>:`) +
                    miniTable(exact, ['Name', 'Department', 'Role', 'Salary', 'City']);
            }
            if (emps.length > 0) {
                return p(`💰 <strong>${emps.length}</strong> employee(s) earning around <strong>$${target.toLocaleString()}</strong> (±20%):`) +
                    miniTable(emps, ['Name', 'Department', 'Role', 'Salary', 'City']);
            }
        }
        if (target <= 20) {
            // Might be experience
            const expCol = findExpCol();
            if (expCol) {
                const emps = excelData.filter(r => num(r, expCol) === target);
                if (emps.length) {
                    return p(`📈 Employees with <strong>${target} years</strong> of experience:`) +
                        miniTable(emps, ['Name', 'Department', 'Role', 'Salary', expCol]);
                }
            }
        }
    }

    // --- NAME-ONLY INPUT (just typing a person's name) ---
    // This catches cases where findMentionedPeople didn't trigger earlier
    if (mentionedPeople.length === 0) {
        // Try a more fuzzy name search
        const cleanQ = q.replace(/[?.,!'"]/g, '').trim();
        const nameMatches = excelData.filter(r => {
            const name = String(r['Name'] || '').toLowerCase();
            // Full or partial match
            return name.includes(cleanQ) || cleanQ.split(/\s+/).some(w => w.length >= 3 && name.includes(w));
        });
        if (nameMatches.length === 1) {
            const person = nameMatches[0];
            let details = '<div style="margin:8px 0">';
            columns.forEach(col => {
                let val = person[col] ?? '';
                if (col === 'Salary' && !isNaN(val)) val = '$' + Number(val).toLocaleString();
                details += `<p style="margin:2px 0"><strong>${esc(col)}:</strong> ${esc(String(val))}</p>`;
            });
            details += '</div>';
            return p(`👤 <strong>${esc(person['Name'])}</strong>`) + details;
        }
        if (nameMatches.length > 1) {
            return p(`👥 Found <strong>${nameMatches.length}</strong> matches for "${esc(cleanQ)}":`) +
                miniTable(nameMatches, ['Name', 'Department', 'Role', 'Salary', 'City', 'Email']);
        }
    }

    // --- PARTIAL MATCH ON ANY COLUMN VALUE ---
    const cleanWords = q.replace(/[?.,!'"]/g, '').split(/\s+/).filter(w => w.length >= 2);
    if (cleanWords.length > 0) {
        // Try to match exact column values first (department, city, role as partial)
        for (const col of ['Department', 'City', 'Role', 'Status']) {
            const vals = unique(col);
            for (const val of vals) {
                if (cleanWords.some(w => val.toLowerCase().includes(w) && w.length >= 3)) {
                    const emps = filterBy(col, val);
                    if (emps.length > 0) {
                        const sals = emps.map(r => num(r, 'Salary')).filter(s => s > 0);
                        const total = sals.reduce((a, b) => a + b, 0);
                        return p(`🔍 Found <strong>${emps.length}</strong> employees in <strong>${col}: ${val}</strong>`) +
                            (sals.length ? p(`💰 Total salary: <strong>$${total.toLocaleString()}</strong>`) : '') +
                            miniTable(emps, ['Name', 'Department', 'Role', 'Salary', 'City']);
                    }
                }
            }
        }

        // Generic keyword search across all columns
        const keywords = cleanWords.filter(w => w.length > 2 && !['the', 'and', 'for', 'are', 'was', 'what', 'who', 'how', 'can', 'you', 'show', 'tell', 'give', 'get', 'from', 'with', 'this', 'that', 'have', 'has', 'does', 'will', 'all', 'any'].includes(w));
        if (keywords.length > 0) {
            const matches = excelData.filter(row =>
                keywords.some(kw => columns.some(col => String(row[col] || '').toLowerCase().includes(kw)))
            );
            if (matches.length > 0) {
                const displayCols = matches.length <= 5 ? columns.slice(0, 8) : ['Name', 'Department', 'Role', 'Salary', 'City'];
                return p(`🔍 Found <strong>${matches.length} result${matches.length > 1 ? 's' : ''}</strong> matching "${esc(query)}":`) +
                    miniTable(matches.slice(0, 20), displayCols);
            }
        }
    }

    // === NO MATCH ===
    return p(`🤔 I couldn't find results for "<em>${esc(query)}</em>".`) +
        p('Try any of these:') +
        '<ul>' +
        '<li>Just type a <strong>name</strong> → "Guru Charan"</li>' +
        '<li>Just type a <strong>city</strong> → "Hyderabad"</li>' +
        '<li>Just type a <strong>department</strong> → "Engineering"</li>' +
        '<li>Just type an <strong>employee ID</strong> → "EMP001"</li>' +
        '<li>Just type a <strong>salary amount</strong> → "80000"</li>' +
        '<li>Just type a <strong>year</strong> → "2023"</li>' +
        '<li>"Group by department" – stats per department</li>' +
        '<li>"Sort by salary desc" – salary ranking</li>' +
        '<li>"Where salary > 80000 and city = Hyderabad"</li>' +
        '<li>"rajesh and charan total salary"</li>' +
        '<li>"Summary" – full overview</li>' +
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

function findMentionedPeople(query) {
    const q = query.toLowerCase();
    const found = [];
    const usedIndices = new Set();
    // Sort names by length (longest first) to match full names before partial
    const sortedByLength = [...excelData].sort((a, b) => String(b['Name'] || '').length - String(a['Name'] || '').length);
    for (const row of sortedByLength) {
        const name = String(row['Name'] || '').toLowerCase();
        if (!name) continue;
        const idx = excelData.indexOf(row);
        if (usedIndices.has(idx)) continue;
        // Check full name
        if (q.includes(name)) { found.push(row); usedIndices.add(idx); continue; }
        // Check individual parts of name (first name / last name)
        const parts = name.split(/\s+/);
        for (const part of parts) {
            if (part.length >= 3 && q.includes(part)) {
                // Make sure this part isn't a common word
                const commonWords = ['the', 'and', 'for', 'are', 'salary', 'total', 'sum', 'average', 'who', 'what', 'how'];
                if (!commonWords.includes(part)) {
                    found.push(row); usedIndices.add(idx); break;
                }
            }
        }
    }
    return found;
}

function resolveColName(input) {
    const lower = input.toLowerCase().replace(/\s+/g, '');
    const map = {
        'salary': 'Salary', 'pay': 'Salary', 'compensation': 'Salary', 'income': 'Salary', 'wage': 'Salary',
        'name': 'Name', 'employee': 'Name',
        'department': 'Department', 'dept': 'Department', 'team': 'Department',
        'city': 'City', 'location': 'City', 'place': 'City',
        'role': 'Role', 'position': 'Role', 'jobtitle': 'Role', 'title': 'Role', 'designation': 'Role',
        'email': 'Email', 'mail': 'Email',
        'status': 'Status',
        'experience': null, // Will resolve dynamically
        'joindate': null,
    };
    if (lower === 'experience' || lower === 'exp') {
        return findExpCol() || 'Experience (Years)';
    }
    if (lower === 'joindate' || lower === 'date') {
        return columns.find(c => c.toLowerCase().includes('join') || c.toLowerCase().includes('date')) || 'Join Date';
    }
    // Direct lookup
    if (map[lower]) return map[lower];
    // Try matching column names
    const found = columns.find(c => c.toLowerCase().includes(lower));
    if (found) return found;
    // Default: capitalize first letter
    return input.charAt(0).toUpperCase() + input.slice(1);
}

function applyCondition(data, condition) {
    const cond = condition.trim().toLowerCase();

    // Pattern: column operator value
    // e.g. "salary > 80000", "department = engineering", "name contains guru"
    const operators = [
        { regex: /(.+?)\s*>=\s*(.+)/, op: '>=' },
        { regex: /(.+?)\s*<=\s*(.+)/, op: '<=' },
        { regex: /(.+?)\s*!=\s*(.+)/, op: '!=' },
        { regex: /(.+?)\s*<>\s*(.+)/, op: '!=' },
        { regex: /(.+?)\s*>\s*(.+)/, op: '>' },
        { regex: /(.+?)\s*<\s*(.+)/, op: '<' },
        { regex: /(.+?)\s*=\s*(.+)/, op: '=' },
        { regex: /(.+?)\s+(?:is|equals?)\s+(.+)/, op: '=' },
        { regex: /(.+?)\s+(?:contains?|like|includes?|has)\s+(.+)/, op: 'contains' },
        { regex: /(.+?)\s+(?:not|isn't|isnt)\s+(.+)/, op: '!=' },
        { regex: /(.+?)\s+(?:starts?\s*with)\s+(.+)/, op: 'startswith' },
        { regex: /(.+?)\s+(?:ends?\s*with)\s+(.+)/, op: 'endswith' },
    ];

    for (const { regex, op } of operators) {
        const m = cond.match(regex);
        if (m) {
            const col = resolveColName(m[1].trim());
            const rawVal = m[2].trim().replace(/['"]/g, '').replace(/[$,]/g, '');

            return data.filter(row => {
                const cellVal = row[col];
                if (cellVal === undefined) return false;
                const cellStr = String(cellVal).toLowerCase();
                const cellNum = Number(String(cellVal).replace(/[$,]/g, ''));
                const valNum = Number(rawVal);
                const valStr = rawVal.toLowerCase();

                switch (op) {
                    case '>': return !isNaN(cellNum) && !isNaN(valNum) && cellNum > valNum;
                    case '<': return !isNaN(cellNum) && !isNaN(valNum) && cellNum < valNum;
                    case '>=': return !isNaN(cellNum) && !isNaN(valNum) && cellNum >= valNum;
                    case '<=': return !isNaN(cellNum) && !isNaN(valNum) && cellNum <= valNum;
                    case '=': return !isNaN(cellNum) && !isNaN(valNum) ? cellNum === valNum : cellStr === valStr;
                    case '!=': return !isNaN(cellNum) && !isNaN(valNum) ? cellNum !== valNum : cellStr !== valStr;
                    case 'contains': return cellStr.includes(valStr);
                    case 'startswith': return cellStr.startsWith(valStr);
                    case 'endswith': return cellStr.endsWith(valStr);
                    default: return true;
                }
            });
        }
    }

    // If no operator found, try contains for any column
    return data.filter(row =>
        columns.some(col => String(row[col] || '').toLowerCase().includes(cond))
    );
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
