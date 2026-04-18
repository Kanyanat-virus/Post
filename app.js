const API_URL = 'https://script.google.com/macros/s/AKfycbxw4dEm75lbDzl6rEEVdDR9uRXqYnKUnG3Rv-TUf2Hvf077VtG5heqCOIam_WTXd6zz-w/exec';

// Global Data variables
let rawData = [];
let processedData = [];
let filteredData = [];

// TomSelect instances
let tsProvince, tsPostOffice, tsCustomer;
let selectedEndDates = new Set(); // Custom date filter state

// Chart.js instances
let barChartInstance = null;
let donutChartInstance = null;

const DOM = {
    loading: document.getElementById('loading'),
    dashboardContent: document.getElementById('dashboardContent'),
    errorMsg: document.getElementById('errorMsg'),
    errorText: document.getElementById('errorText'),
    cardsContainer: document.getElementById('cardsContainer'),
    recordCount: document.getElementById('recordCount'),
    noDataResult: document.getElementById('noDataResult'),
    refreshBtn: document.getElementById('refreshBtn'),
    resetFilterBtn: document.getElementById('resetFilterBtn')
};

// Column Names mapping based on the CSV
const COL = {
    timestamp: 'ประทับเวลา',
    province: 'จังหวัดที่สังกัด',
    postOffice: 'จุดที่เปิดให้บริการ (สามารถพิมพ์รหัส ปณ.เพื่อค้นหา ในช่องด้านล่างได้)',
    customer: 'ชื่อลูกค้าที่อนุมัติสัญญาฯ (กรณีต้องการอัพเดทข้อมูล โปรดระบุชื่อลูกค้าให้ถูกต้องตรงกับชื่อเดิมทุกตัวอักษร)',
    package: 'Package ที่อนุมัติสัญญาฯ',
    file: 'แนบไฟล์รวมบันทึกทั้ง 3 หัวข้อ :\n1.แบบร่างอนุมัติ + สัญญาขนส่ง\n2.แจ้งปรับอัตราค่าบริการเรียกเก็บค่าขนส่งเพิ่มเติม\n3.หนังสือยืนยันการใช้บริการ Fuel Surcharge',
    itemType: 'ประเภทสิ่งของที่ฝากส่ง (ห้ามใช้กับสินค้าประเภท ผลไม้ ต้นไม้)',
    itemTypeOther: 'โปรดระบุประเภทสิ่งของที่ฝากส่งกรณีเลือก สินค้าประเภทอื่น ๆ',
    endDate: 'วันที่สิ้นสุดสัญญาฯ'
};

document.addEventListener('DOMContentLoaded', () => {
    initSelects();
    initAuth(); // Check auth before loading data

    DOM.refreshBtn.addEventListener('click', fetchData);
    DOM.resetFilterBtn.addEventListener('click', resetFilters);
    initDateFilter();

    // Logout button
    document.getElementById('logoutBtn')?.addEventListener('click', () => {
        sessionStorage.removeItem('dashboardAuth');
        sessionStorage.removeItem('dashboardUser');
        sessionStorage.removeItem('dashboardEmail');
        sessionStorage.removeItem('dashboardPass');
        location.reload();
    });

    setupAdminFeatures();

    // FAB scroll to top logic
    const fab = document.getElementById('scrollToTop');
    if (fab) {
        window.addEventListener('scroll', () => {
            if (window.scrollY > 300) fab.classList.add('visible');
            else fab.classList.remove('visible');
        });
        fab.addEventListener('click', () => {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }

    setupIdleTimeout();
});

// ─── IDLE TIMEOUT ──────────────────────────────────────────────────────────
let idleTimeoutTimer;
function setupIdleTimeout() {
    const handleUserActivity = () => {
        // ทำงานเฉพาะเมื่อ Login แล้วเท่านั้น
        if (sessionStorage.getItem('dashboardAuth') !== 'true') return;
        
        clearTimeout(idleTimeoutTimer);
        // ระยะเวลา 1 ชั่วโมง = 60 * 60 * 1000 = 3600000 ms
        idleTimeoutTimer = setTimeout(() => {
            sessionStorage.removeItem('dashboardAuth');
            sessionStorage.removeItem('dashboardUser');
            alert("ระบบได้ออกจากระบบอัตโนมัติ เนื่องจากไม่มีการใช้งานเกิน 1 ชั่วโมง กรุณาเข้าสู่ระบบใหม่อีกครั้ง");
            location.reload();
        }, 3600000);
    };

    window.addEventListener('mousemove', handleUserActivity);
    window.addEventListener('keydown', handleUserActivity);
    window.addEventListener('click', handleUserActivity);
    window.addEventListener('scroll', handleUserActivity, true);
    
    // เริ่มจับเวลาครั้งแรกทันทีที่รัน
    handleUserActivity();
}

// ─── AUTH ──────────────────────────────────────────────────────────────────

async function initAuth() {
    // Already logged in this session
    if (sessionStorage.getItem('dashboardAuth') === 'true') {
        const userName = sessionStorage.getItem('dashboardUser') || '';
        hideLoginOverlay(userName);
        fetchData();
        return;
    }
    setupLoginForm();
}

function setupLoginForm() {
    const form = document.getElementById('loginForm');
    const pwdInput = document.getElementById('inputPassword');
    const toggleBtn = document.getElementById('togglePassword');

    toggleBtn?.addEventListener('click', () => {
        const show = pwdInput.type === 'password';
        pwdInput.type = show ? 'text' : 'password';
        document.getElementById('eyeIcon').setAttribute('data-lucide', show ? 'eye-off' : 'eye');
        lucide.createIcons();
    });

    form?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const email = document.getElementById('inputEmail').value.trim().toLowerCase();
        const password = pwdInput.value.trim();
        const loginBtn = document.getElementById('loginBtn');
        const errEl = document.getElementById('loginError');

        loginBtn.disabled = true;
        loginBtn.innerHTML = '<i data-lucide="loader-circle" class="spin"></i> กำลังตรวจสอบ...';
        lucide.createIcons();
        errEl.style.display = 'none';

        try {
            // Fetch backend API to login
            const url = `${API_URL}?action=login&email=${encodeURIComponent(email)}&pass=${encodeURIComponent(password)}`;
            const res = await fetch(url);
            const result = await res.json();

            loginBtn.disabled = false;
            loginBtn.innerHTML = '<i data-lucide="log-in"></i> เข้าสู่ระบบ';
            lucide.createIcons();

            if (result.ok) {
                // Save user session
                sessionStorage.setItem('dashboardAuth', 'true');
                sessionStorage.setItem('dashboardUser', result.name);
                sessionStorage.setItem('dashboardEmail', email); // Save for admin check
                sessionStorage.setItem('dashboardPass', password); // Support API queries

                hideLoginOverlay(result.name);
                setupAdminFeatures();
                fetchData();
            } else {
                errEl.textContent = result.msg || 'เข้าสู่ระบบล้มเหลว';
                errEl.style.display = 'block';
                const card = document.querySelector('.login-card');
                card.classList.add('shake');
                setTimeout(() => card.classList.remove('shake'), 500);
            }
        } catch (err) {
            console.error('Login error:', err);
            loginBtn.disabled = false;
            loginBtn.innerHTML = '<i data-lucide="log-in"></i> เข้าสู่ระบบ';
            lucide.createIcons();
            
            errEl.textContent = 'เกิดข้อผิดพลาดในการเชื่อมต่อเซิร์ฟเวอร์';
            errEl.style.display = 'block';
        }
    });
}

function hideLoginOverlay(userName) {
    const overlay = document.getElementById('loginOverlay');
    if (!overlay) return;
    if (userName) {
        const greetEl = document.getElementById('userGreeting');
        if (greetEl) greetEl.textContent = `สวัสดี, ${userName}`;
    }
    overlay.classList.add('fade-out');
    setTimeout(() => overlay.style.display = 'none', 400);
}

// ==== Admin & Logs Logic ====


function setupAdminFeatures() {
    const email = sessionStorage.getItem('dashboardEmail');
    const adminBtn = document.getElementById('dashboardIcon');
    const logsModal = document.getElementById('logsModalOverlay');
    const closeBtn = document.getElementById('closeLogsBtn');
    
    if (!adminBtn || !logsModal) return;

    // Check if admin
    if (email === 'sales.reg6@gmail.com' || email === 'kanyanat.ra@thailandpost.com') {
        adminBtn.classList.add('admin-active');
        adminBtn.style.cursor = 'pointer';
        adminBtn.title = 'ดูประวัติการเข้าใช้งาน (Admin)';
        
        // Remove old listeners to avoid duplicates
        const newBtn = adminBtn.cloneNode(true);
        adminBtn.parentNode.replaceChild(newBtn, adminBtn);
        
        newBtn.addEventListener('click', () => {
            logsModal.style.display = 'flex';
            fetchLogData();
        });
    } else {
        adminBtn.classList.remove('admin-active');
        adminBtn.style.cursor = 'default';
        adminBtn.title = '';
        
        const newBtn = adminBtn.cloneNode(true);
        adminBtn.parentNode.replaceChild(newBtn, adminBtn);
    }

    if (closeBtn) {
        closeBtn.addEventListener('click', () => {
            logsModal.style.display = 'none';
        });
    }

    const modalScrollBtn = document.getElementById('modalScrollTopBtn');
    const logsContainer = document.getElementById('logsContainer');
    
    if (modalScrollBtn && logsContainer) {
        logsContainer.addEventListener('scroll', () => {
            if (logsContainer.scrollTop > 150) {
                modalScrollBtn.style.opacity = '1';
                modalScrollBtn.style.pointerEvents = 'auto';
                modalScrollBtn.style.transform = 'translateY(-10px)';
            } else {
                modalScrollBtn.style.opacity = '0';
                modalScrollBtn.style.pointerEvents = 'none';
                modalScrollBtn.style.transform = 'translateY(0)';
            }
        });
        
        modalScrollBtn.addEventListener('click', () => {
            logsContainer.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        });
    }
}

let currentLogData = [];

async function fetchLogData() {
    const container = document.getElementById('logsContainer');
    if (!container) return;
    
    container.innerHTML = '<div style="text-align: center; padding: 2rem;"><i data-lucide="loader-circle" class="spin"></i> กำลังโหลดประวัติ...</div>';
    lucide.createIcons();
    
    try {
        const email = sessionStorage.getItem('dashboardEmail') || '';
        const pass = sessionStorage.getItem('dashboardPass') || '';
        const url = `${API_URL}?action=getLogs&email=${encodeURIComponent(email)}&pass=${encodeURIComponent(pass)}`;
        
        const res = await fetch(url);
        const result = await res.json();
        
        if (!result.ok || !result.data || result.data.length <= 1) {
            container.innerHTML = '<div style="text-align: center; padding: 2rem; color: var(--text-muted);">ไม่มีประวัติการเข้าใช้งาน</div>';
            return;
        }
        
        // Remove header row and sort descending by timestamp (newest first)
        currentLogData = result.data.slice(1).reverse(); 
        renderLogs(false);
        
    } catch (err) {
        console.error('Fetch Logs Error:', err);
        container.innerHTML = '<div style="text-align: center; padding: 2rem; color: var(--danger);">เกิดข้อผิดพลาดในการโหลดประวัติ</div>';
    }
}

function renderLogs(showAll) {
    const container = document.getElementById('logsContainer');
    if (!container || currentLogData.length === 0) return;
    
    let html = '';
    const limit = showAll ? currentLogData.length : 10;
    const displayData = currentLogData.slice(0, limit);
    
    displayData.forEach(row => {
        // GAS returns arrays: row[0] is timestamp, row[1] is email
        const tsRaw = row[0];
        const emailRaw = row[1];
        if (!emailRaw) return;
        
        // The display string is already handled by getDisplayValues() in GAS
        let timeStr = String(tsRaw);
        
        html += `
            <div style="display: flex; gap: 1rem; align-items: center; padding: 0.75rem; border-bottom: 1px solid var(--border-color);">
                <div class="cell-avatar bg-c3" style="width: 32px; height: 32px;"><i data-lucide="user" style="width: 16px; height: 16px;"></i></div>
                <div style="flex: 1;">
                    <div style="font-weight: 600; color: var(--text-dark); font-size: 0.95rem;">${emailRaw}</div>
                    <div style="font-size: 0.8rem; color: var(--text-muted);"><i data-lucide="clock" style="width:12px;height:12px;display:inline;margin-right:4px;"></i> ${timeStr}</div>
                </div>
            </div>
        `;
    });
    
    if (!showAll && currentLogData.length > 10) {
        html += `
            <div style="text-align: center; padding: 1rem;">
                <button id="showAllLogsBtn" class="action-btn" style="border: 1px solid var(--border-color); background: #fff; padding: 0.5rem 1rem; margin: 0 auto; width: fit-content;">ดูประวัติทั้งหมด (${currentLogData.length})</button>
            </div>
        `;
    }
    
    container.innerHTML = html;
    lucide.createIcons();
    
    if (!showAll && currentLogData.length > 10) {
        document.getElementById('showAllLogsBtn')?.addEventListener('click', () => {
            renderLogs(true);
        });
    }
}


function initSelects() {
    const commonConfig = {
        create: false,
        maxItems: 1, // Enforce single selection to remove empty cursor rows
        sortField: {
            field: "text",
            direction: "asc"
        },
        onChange: handleFilterChange,
        onItemAdd: function() {
            this.blur(); // Hide the blinking cursor immediately after selecting
        }
    };

    tsCustomer = new TomSelect('#filterCustomer', { ...commonConfig, placeholder: "พิมพ์ค้นหาชื่อลูกค้า..." });
    tsProvince = new TomSelect('#filterProvince', { ...commonConfig, placeholder: "จังหวัด" });
    tsPostOffice = new TomSelect('#filterPostOffice', { ...commonConfig, placeholder: "ที่ทำการ/รหัส ปณ." });
}

// --- Date Filter (custom checkbox dropdown) ---
let allAvailableEndDates = [];

function initDateFilter() {
    const btn = document.getElementById('dateFilterBtn');
    const dropdown = document.getElementById('dateFilterDropdown');

    btn?.addEventListener('click', (e) => {
        e.stopPropagation();
        const isOpen = dropdown.style.display !== 'none';
        dropdown.style.display = isOpen ? 'none' : 'block';
        lucide.createIcons();
    });

    document.addEventListener('click', (e) => {
        if (!e.target.closest('.date-filter-wrap')) {
            if (dropdown) dropdown.style.display = 'none';
        }
    });

    document.getElementById('dateSelectAll')?.addEventListener('click', () => {
        selectedEndDates.clear();
        renderDateCheckboxList(allAvailableEndDates);
        updateDateFilterLabel();
        handleFilterChange();
    });
}

function renderDateCheckboxList(dates) {
    const container = document.getElementById('dateCheckboxList');
    if (!container) return;
    container.innerHTML = '';
    const sorted = [...dates]; // Already chronologically sorted ahead of time
    sorted.forEach(d => {
        const isChecked = selectedEndDates.size === 0 || selectedEndDates.has(d);
        const item = document.createElement('label');
        item.className = 'date-checkbox-item';
        
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = d;
        cb.checked = isChecked;
        
        const span = document.createElement('span');
        span.textContent = d;
        span.style.flex = "1";
        
        const onlyBtn = document.createElement('button');
        onlyBtn.type = 'button';
        onlyBtn.className = 'date-only-btn';
        onlyBtn.textContent = 'เฉพาะวันนี้';
        
        onlyBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            selectedEndDates.clear();
            selectedEndDates.add(d);
            updateDateFilterLabel();
            handleFilterChange();
            renderDateCheckboxList(allAvailableEndDates);
        });

        cb.addEventListener('change', () => {
            if (selectedEndDates.size === 0) {
                allAvailableEndDates.forEach(od => { if (od !== d) selectedEndDates.add(od); });
            } else {
                if (cb.checked) selectedEndDates.add(d);
                else selectedEndDates.delete(d);
                if (selectedEndDates.size === allAvailableEndDates.length) selectedEndDates.clear();
            }
            updateDateFilterLabel();
            handleFilterChange();
            renderDateCheckboxList(allAvailableEndDates);
        });
        
        item.appendChild(cb);
        item.appendChild(span);
        item.appendChild(onlyBtn);
        
        container.appendChild(item);
    });
}

function updateDateFilterLabel() {
    var label = document.getElementById('dateFilterLabel');
    var btn = document.getElementById('dateFilterBtn');
    if (!label) return;
    if (selectedEndDates.size === 0) {
        label.textContent = 'วันที่สิ้นสุดสัญญา';
        if (btn) btn.classList.remove('has-selection');
    } else if (selectedEndDates.size === 1) {
        label.textContent = [...selectedEndDates][0];
        if (btn) btn.classList.add('has-selection');
    } else {
        label.textContent = 'เลือก ' + selectedEndDates.size + ' วันที่';
        if (btn) btn.classList.add('has-selection');
    }
}

// --- AI Alert Section ---
function renderAIAlert(data) {
    var body = document.getElementById('aiAlertBody');
    var monthLabel = document.getElementById('aiAlertMonthLabel');
    if (!body) return;

    var now = new Date();
    var thisMonth = now.getMonth();
    var thisYear = now.getFullYear();
    var thMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

    // Dynamic month based on date filter
    if (selectedEndDates && selectedEndDates.size > 0) {
        var firstDateStr = Array.from(selectedEndDates)[0];
        if (firstDateStr && firstDateStr !== 'ไม่ระบุ') {
            var parts = firstDateStr.split('/');
            if (parts.length === 3) {
                thisMonth = parseInt(parts[1]) - 1;
                thisYear = parseInt(parts[2]);
            }
        }
    }

    if (monthLabel) monthLabel.textContent = 'ประจำเดือน ' + thMonths[thisMonth] + ' ' + (thisYear + 543);

    var expiring = data.filter(function(item) {
        var d = item.processedContractEndDate;
        if (!d || d === 'ไม่ระบุ') return false;
        var parts = d.split('/');
        if (parts.length !== 3) return false;
        var month = parseInt(parts[1]) - 1;
        var year = parseInt(parts[2]);
        return month === thisMonth && year === thisYear;
    });

    var a11 = expiring.filter(function(i) { return i.processedPackage && i.processedPackage.includes('A11'); });
    var a12 = expiring.filter(function(i) { return i.processedPackage && i.processedPackage.includes('A12'); });
    var other = expiring.filter(function(i) { return !(i.processedPackage && (i.processedPackage.includes('A11') || i.processedPackage.includes('A12'))); });

    if (expiring.length === 0) {
        body.innerHTML = '<p class="ai-no-expiry">✅ เดือนนี้ไม่มีลูกค้าที่สัญญาสิ้นสุด ทุกสัญญายังคงมีผลอยู่</p>';
        return;
    }

    var otherHTML = other.length > 0 ? ', อื่นๆ: <strong>' + other.length + ' ราย</strong>' : '';

    body.innerHTML =
        '<div class="ai-alert-stat-card main-stat">' +
            '<div class="ai-stat-label">สิ้นสุดเดือนนี้</div>' +
            '<div class="ai-stat-num" style="color:var(--primary)">' + expiring.length + '</div>' +
            '<div class="ai-stat-sub">ราย</div>' +
        '</div>' +
        '<div class="ai-alert-stat-card">' +
            '<div class="ai-stat-label">Package A11</div>' +
            '<div class="ai-stat-num" style="color:#E87A13">' + a11.length + '</div>' +
            '<div class="ai-stat-sub">ราย</div>' +
        '</div>' +
        '<div class="ai-alert-stat-card">' +
            '<div class="ai-stat-label">Package A12</div>' +
            '<div class="ai-stat-num" style="color:#2F8E6E">' + a12.length + '</div>' +
            '<div class="ai-stat-sub">ราย</div>' +
        '</div>' +
        '<div class="ai-alert-message">' +
            '📋 <strong>สรุปสัญญาสิ้นสุดเดือน' + thMonths[thisMonth] + ':</strong><br>' +
            'มีลูกค้า <strong>' + expiring.length + ' ราย</strong> ' +
            '(Package A11: <strong>' + a11.length + ' ราย</strong>, ' +
            'Package A12: <strong>' + a12.length + ' ราย</strong>' + otherHTML + ')' +
        '</div>';
}

// Helper to convert DD/MM/YYYY, HH:mm:ss to timestamp
function parseDateStrToTimestamp(dateStr) {
    if (!dateStr) return 0;
    try {
        const parts = dateStr.split(', ');
        if (parts.length === 2) {
            const [d, m, y] = parts[0].split('/');
            const [hh, mm, ss] = parts[1].split(':');
            return new Date(y, parseInt(m) - 1, d, hh, mm, ss).getTime();
        }
    } catch (e) {
        console.warn("Error parsing date:", dateStr);
    }
    return 0; // fallback
}

// Fetch data securely using Google Apps Script Web App
async function fetchData() {
    showLoading(true);
    
    try {
        const email = sessionStorage.getItem('dashboardEmail') || '';
        const pass = sessionStorage.getItem('dashboardPass') || '';
        const url = `${API_URL}?action=getData&email=${encodeURIComponent(email)}&pass=${encodeURIComponent(pass)}`;
        
        const res = await fetch(url);
        const result = await res.json();
        
        if (!result.ok) {
            if (result.msg && result.msg.includes('Unauthorized')) {
                alert('เซสชั่นของคุณหมดอายุ หรือรหัสผ่านมีการเปลี่ยนแปลง กรุณาเข้าสู่ระบบใหม่');
                forceLogout();
                return;
            }
            showError("ไม่สามารถอ่านข้อมูลจากตารางได้ หรืออาจยังไม่มีข้อมูลลูกค้า");
            showLoading(false);
            return;
        }

        if (!result.data || result.data.length < 2) {
            showError("ไม่พบข้อมูลลูกค้าในระบบ");
            showLoading(false);
            return;
        }
        
        const headers = result.data[0];
        const rows = result.data.slice(1);
        
        // Convert 2D array to array of objects to maintain compatibility with existing logic
        const formattedData = rows.map(row => {
            let rowObj = {};
            row.forEach((cell, i) => {
                const header = headers[i];
                if (header) {
                    rowObj[header] = cell ? String(cell) : '';
                }
            });
            return rowObj;
        });

        processRawData(formattedData);
        showLoading(false);
    } catch (error) {
        console.error("fetch API error:", error);
        showError(`การดึงข้อมูล Google API ล้มเหลวสุดๆ 😅: ${error.message} \nบราวเซอร์อาจถูกตัดการเชื่อมต่อ หรือเซิร์ฟเวอร์ไม่ตอบสนอง`);
        showLoading(false);
    }
}

function processRawData(data) {
    rawData = data;
    
    // Group by customer to find the latest file
    const recordsMap = {};
    
    rawData.forEach(row => {
        const cName = row[COL.customer]?.trim();
        if (!cName) return; // Skip empty rows

        const ts = parseDateStrToTimestamp(row[COL.timestamp]);
        
        // Use customer name as key. If name matches, we only keep the latest one.
        if (!recordsMap[cName] || recordsMap[cName].parsedTs < ts) {
            recordsMap[cName] = {
                ...row,
                parsedTs: ts,
                // Sanitize string values
                processedProvince: row[COL.province]?.trim() || 'ไม่ระบุ',
                processedPostOffice: row[COL.postOffice]?.trim() || 'ไม่ระบุ',
                processedCustomer: cName,
                processedPackage: row[COL.package]?.trim() || '-',
                processedItemType: row[COL.itemType]?.trim() || '-',
                processedItemTypeOther: row[COL.itemTypeOther]?.trim() || '',
                processedContractEndDate: row[COL.endDate]?.trim() || 'ไม่ระบุ',
                processedFile: row[COL.file]?.trim() || ''
            };
        }
    });

    processedData = Object.values(recordsMap);
    
    // Initialize filters and render
    updateFilterOptions(processedData);
    renderCards(processedData);
    
    // Init charts on first load, update after
    if (!barChartInstance) initCharts();
    updateCharts(processedData);
    renderAIAlert(processedData);
}

// Update the options available in dropdowns based on the available dataset
// but retain the current selection if it exists.
function updateFilterOptions(dataSet) {
    const provinces = new Set();
    const postOffices = new Set();
    const customers = new Set();
    const endDates = new Set();

    dataSet.forEach(item => {
        if(item.processedProvince) provinces.add(item.processedProvince);
        if(item.processedPostOffice) postOffices.add(item.processedPostOffice);
        if(item.processedCustomer) customers.add(item.processedCustomer);
        if(item.processedContractEndDate && item.processedContractEndDate !== 'ไม่ระบุ') endDates.add(item.processedContractEndDate);
    });

    // Update available date list for checkbox filter
    allAvailableEndDates = [...endDates].sort((a, b) => {
        const p1 = a.split('/');
        const p2 = b.split('/');
        if (p1.length === 3 && p2.length === 3) {
            const d1 = new Date(p1[2], parseInt(p1[1]) - 1, p1[0]);
            const d2 = new Date(p2[2], parseInt(p2[1]) - 1, p2[0]);
            return d1.getTime() - d2.getTime();
        }
        return a.localeCompare(b);
    });
    renderDateCheckboxList(allAvailableEndDates);

    // Helper to update TomSelect silently
    const updateTS = (tsInstance, items) => {
        const currentVal = tsInstance.getValue();
        tsInstance.clearOptions();
        const optionsList = Array.from(items).map(i => ({value: i, text: i}));
        tsInstance.addOptions(optionsList);
        if (Array.isArray(currentVal)) {
            const validVals = currentVal.filter(v => items.has(v));
            if (validVals.length > 0) tsInstance.setValue(validVals, true);
            else tsInstance.clear(true);
        } else {
            if (currentVal && items.has(currentVal)) tsInstance.setValue(currentVal, true);
            else tsInstance.clear(true);
        }
    };

    updateTS(tsProvince, provinces);
    updateTS(tsPostOffice, postOffices);
    updateTS(tsCustomer, customers);
}

let isUpdatingFilters = false;

function handleFilterChange() {
    if (isUpdatingFilters) return;

    const sProv = tsProvince.getValue();
    const sPO = tsPostOffice.getValue();
    const sCust = tsCustomer.getValue();
    // Use custom checkbox set; empty = all selected
    const sEndDates = selectedEndDates.size > 0 ? [...selectedEndDates] : [];

    filteredData = processedData.filter(item => {
        const matchProv = !sProv || item.processedProvince === sProv;
        const matchPO = !sPO || item.processedPostOffice === sPO;
        const matchCust = !sCust || item.processedCustomer === sCust;
        const matchEndDate = sEndDates.length === 0 || sEndDates.includes(item.processedContractEndDate);
        return matchProv && matchPO && matchCust && matchEndDate;
    });

    isUpdatingFilters = true;
    updateCascadingOptions(sProv, sPO, sCust);
    isUpdatingFilters = false;

    renderCards(filteredData);
    updateCharts(filteredData);
    renderAIAlert(filteredData);
}

// Function specifically to update dropdowns based on selections without infinite loops
function updateCascadingOptions(sProv, sPO, sCust) {
    let availableForProv = processedData;
    if (sPO) availableForProv = availableForProv.filter(i => i.processedPostOffice === sPO);
    if (sCust) availableForProv = availableForProv.filter(i => i.processedCustomer === sCust);

    let availableForPO = processedData;
    if (sProv) availableForPO = availableForPO.filter(i => i.processedProvince === sProv);
    if (sCust) availableForPO = availableForPO.filter(i => i.processedCustomer === sCust);
    
    let availableForCust = processedData;
    if (sProv) availableForCust = availableForCust.filter(i => i.processedProvince === sProv);
    if (sPO) availableForCust = availableForCust.filter(i => i.processedPostOffice === sPO);

    const provSet = new Set(availableForProv.map(i => i.processedProvince));
    updateTSVisualOnly(tsProvince, provSet, sProv);
    const poSet = new Set(availableForPO.map(i => i.processedPostOffice));
    updateTSVisualOnly(tsPostOffice, poSet, sPO);
    const custSet = new Set(availableForCust.map(i => i.processedCustomer));
    updateTSVisualOnly(tsCustomer, custSet, sCust);
}

function updateTSVisualOnly(tsInstance, itemsSet, currentVal) {
    let targetVal = currentVal;
    if (itemsSet.size === 1 && (!currentVal || currentVal.length === 0)) {
        targetVal = Array.from(itemsSet)[0]; // Auto-select if there's exactly 1 valid option left
    }

    const opts = Array.from(itemsSet).map(i => ({value: i, text: i}));
    
    // This will trigger 'change' but it is safely trapped by isUpdatingFilters
    tsInstance.clearOptions();
    tsInstance.addOptions(opts);
    
    if (Array.isArray(targetVal)) {
        const validVals = targetVal.filter(v => itemsSet.has(v));
        if (validVals.length > 0) tsInstance.setValue(validVals, true);
        else tsInstance.clear(true);
    } else {
        if (targetVal && itemsSet.has(targetVal)) {
            tsInstance.setValue(targetVal, true);
        } else {
            tsInstance.clear(true);
        }
    }
}

function resetFilters() {
    tsProvince.clear(true);
    tsPostOffice.clear(true);
    tsCustomer.clear(true);
    selectedEndDates.clear();
    renderDateCheckboxList(allAvailableEndDates);
    updateDateFilterLabel();
    
    updateFilterOptions(processedData);
    renderCards(processedData);
    updateCharts(processedData);
    renderAIAlert(processedData);
}

// ─── CHARTS ────────────────────────────────────────────────────────────────

const CHART_PALETTE = [
    '#F94144', // red
    '#F3722C', // orange-red
    '#F8961E', // orange
    '#F9844A', // peachy-orange
    '#F9C74F', // yellow
    '#90BE6D', // light green
    '#43AA8B', // green
    '#4D908E', // teal
    '#577590', // blue-grey
    '#277DA1'  // blue
];

// ใช้ Palette เดียวกันกับกราฟวงกลม เพื่อความคุมโทน
const DONUT_PALETTE = [
    '#F94144',
    '#F3722C',
    '#F8961E',
    '#F9844A',
    '#F9C74F',
    '#90BE6D',
    '#43AA8B',
    '#4D908E',
    '#577590',
    '#277DA1'
];

function initCharts() {
    const barCtx = document.getElementById('barChart').getContext('2d');
    const donutCtx = document.getElementById('donutChart').getContext('2d');

    barChartInstance = new Chart(barCtx, {
        type: 'bar',
        plugins: [ChartDataLabels], // แสดงตัวเลขบนแท่ง
        data: { labels: [], datasets: [{ label: 'จำนวนลูกค้า', data: [], backgroundColor: CHART_PALETTE, borderRadius: 8 }] },
        options: {
            responsive: true, maintainAspectRatio: false,
            layout: { padding: { top: 24 } }, // เว้นไว้สำหรับ datalabel ด้านบน
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: ctx => ` ${ctx.parsed.y} ราย` } },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: '#6B7280',
                    font: { family: 'Prompt', size: 11, weight: 'bold' },
                    formatter: (value) => value
                }
            },
            scales: {
                x: { ticks: { font: { family: 'Prompt', size: 11 }, maxRotation: 45 }, grid: { display: false } },
                y: { ticks: { font: { family: 'Prompt' }, stepSize: 1 }, grid: { color: '#f1f5f9' } }
            }
        }
    });


    donutChartInstance = new Chart(donutCtx, {
        type: 'doughnut',
        plugins: [ChartDataLabels], // register only for this chart, not bar chart
        data: { labels: [], datasets: [{ data: [], backgroundColor: CHART_PALETTE, borderWidth: 2, borderColor: '#fff', hoverOffset: 8 }] },
        options: {
            responsive: true, maintainAspectRatio: false,
            cutout: '60%',
            layout: { padding: { top: 4, bottom: 4 } },
            plugins: {
                legend: {
                    position: 'bottom',
                    align: 'start',
                    onClick: null, // ปิด click-to-hide segment
                    labels: {
                        font: { family: 'Prompt', size: 11 },
                        padding: 10,
                        boxWidth: 12,
                        boxHeight: 12,
                        usePointStyle: true,
                        pointStyle: 'rectRounded'
                    }
                },
                datalabels: {
                    color: '#1F2937',
                    font: { family: 'Prompt', size: 11, weight: 'bold' },
                    textShadowBlur: 3,
                    textShadowColor: 'rgba(255,255,255,0.85)',
                    formatter: (value, ctx) => {
                        const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                        if (!total) return '';
                        const pct = ((value / total) * 100);
                        return pct >= 7 ? pct.toFixed(1) + '%' : '';
                    }
                },
                tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} ราย (${((ctx.parsed / ctx.dataset.data.reduce((a,b)=>a+b,0))*100).toFixed(1)}%)` } }
            }
        }
    });

}

function updateCharts(data) {
    const sProv = tsProvince ? tsProvince.getValue() : '';
    const sPO   = tsPostOffice ? tsPostOffice.getValue() : '';

    // Stats card - total
    const total = data.length;
    document.getElementById('statTotal').textContent = total;
    const ctx = sProv ? `จังหวัด ${sProv}` : (sPO ? `ที่ทำการ ${sPO}` : 'ทั้งระบบ');
    document.getElementById('statContext').textContent = ctx;

    // Package breakdown
    const pkgGroups = {};
    data.forEach(item => {
        const key = item.processedPackage || '-';
        pkgGroups[key] = (pkgGroups[key] || 0) + 1;
    });
    const pkgContainer = document.getElementById('statPackages');
    if (pkgContainer) {
        pkgContainer.innerHTML = Object.entries(pkgGroups)
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([pkg, count]) => {
                const pct = total ? ((count / total) * 100).toFixed(0) : 0;
                const isA11 = pkg.includes('A11');
                const isA12 = pkg.includes('A12');
                const bgObj = isA11
                    ? { bg: 'rgba(232, 122, 19, 0.12)', border: 'rgba(232, 122, 19, 0.3)', text: '#D16D11', unit: '#E87A13', sep: 'rgba(232, 122, 19, 0.4)' }
                    : isA12
                    ? { bg: 'rgba(47, 142, 110, 0.12)', border: 'rgba(47, 142, 110, 0.3)', text: '#277A5E', unit: '#2F8E6E', sep: 'rgba(47, 142, 110, 0.4)' }
                    : { bg: '#F8FAFC', border: '#E2E8F0', text: '#334155', unit: '#64748B', sep: '#CBD5E1' };

                return `<div class="pkg-row" style="background:${bgObj.bg}; border:1.5px solid ${bgObj.border}; margin-bottom: 1rem; padding: 1rem; display: flex; flex-direction: column; align-items: center; justify-content: center; border-radius: 12px; gap: 0.25rem;">
                    <span class="pkg-label" style="color:${bgObj.text}; font-size: 0.95rem; font-weight: 700; letter-spacing: 0.3px;">${pkg}</span>
                    <div class="pkg-stats" style="display: flex; align-items: baseline; justify-content: center; gap: 0.3rem;">
                        <span class="pkg-count-num" style="color:${bgObj.text}; font-weight: 800; font-size: 1.5rem;">${count}</span>
                        <span class="pkg-unit" style="color:${bgObj.unit}; font-size: 0.9rem; font-weight: 600;"> ราย</span>
                        <span class="pkg-sep" style="color:${bgObj.sep}; font-size: 1rem; font-weight: bold; margin: 0 0.2rem;">·</span>
                        <span class="pkg-pct" style="color:${bgObj.text}; font-size: 1.1rem; font-weight: 700;">${pct}%</span>
                    </div>
                </div>`;
            }).join('');
    }

    // Bar Chart: group by Province normally, by PostOffice if province is selected
    const groupByPO = !(sProv === '');
    const barLabel = groupByPO ? 'ที่ทำการ' : 'จังหวัด';
    document.getElementById('barChartLabel').textContent = barLabel;

    const barGroups = {};
    data.forEach(item => {
        const key = groupByPO ? item.processedPostOffice : item.processedProvince;
        barGroups[key] = (barGroups[key] || 0) + 1;
    });
    const barEntries = Object.entries(barGroups).sort((a,b) => b[1]-a[1]);
    barChartInstance.data.labels = barEntries.map(e => e[0]);
    barChartInstance.data.datasets[0].data = barEntries.map(e => e[1]);
    barChartInstance.data.datasets[0].backgroundColor = barEntries.map((_,i) => CHART_PALETTE[i % CHART_PALETTE.length]);
    barChartInstance.data.datasets[0].label = `ลูกค้า (${barLabel})`;
    barChartInstance.update();

    // Donut Chart: group by item type (all 'สินค้าประเภทอื่น ๆ' grouped together)
    const typeGroups = {};
    data.forEach(item => {
        const key = item.processedItemType; // Always use base type, ignore the specified detail
        typeGroups[key] = (typeGroups[key] || 0) + 1;
    });
    const typeEntries = Object.entries(typeGroups).sort((a,b) => b[1]-a[1]);
    donutChartInstance.data.labels = typeEntries.map(e => e[0]);
    donutChartInstance.data.datasets[0].data = typeEntries.map(e => e[1]);
    donutChartInstance.data.datasets[0].backgroundColor = typeEntries.map(e => getItemTypeChartColor(e[0]));
    donutChartInstance.update();
}

const PASTEL_COLORS = [
    { bg: '#DBEAFE', text: '#1E40AF' }, // Blue
    { bg: '#D1FAE5', text: '#065F46' }, // Green
    { bg: '#FCE7F3', text: '#9D174D' }, // Pink
    { bg: '#EDE9FE', text: '#5B21B6' }, // Purple
    { bg: '#FFEDD5', text: '#9A3412' }, // Orange
    { bg: '#E0E7FF', text: '#3730A3' }, // Indigo
    { bg: '#CCFBF1', text: '#115E59' }  // Teal
];

function getPkgBadgeStyle(pkg) {
    if (!pkg || pkg === '-') return '';
    if (pkg.includes('A11')) return 'background:#FEF3C7;color:#92400E;border-color:#FCD34D;';
    if (pkg.includes('A12')) return 'background:#D1FAE5;color:#065F46;border-color:#6EE7B7;';
    return '';
}

// ใช้ DONUT_PALETTE โทนใหม่!
function getItemTypeChartColor(typeStr) {
    if (!typeStr || typeStr === '-') return '#E5E7EB';
    if (typeStr.includes('สินค้าประเภทอื่น ๆ')) return '#F9C74F'; // yellow
    let hash = 0;
    for (let i = 0; i < typeStr.length; i++) hash = typeStr.charCodeAt(i) + ((hash << 5) - hash);
    return DONUT_PALETTE[Math.abs(hash) % DONUT_PALETTE.length];
}

function getItemTypeTextColor(typeStr) {
    if (!typeStr || typeStr === '-') return '#6B7280';
    if (typeStr.includes('สินค้าประเภทอื่น ๆ')) return '#92400E';
    let hash = 0;
    for (let i = 0; i < typeStr.length; i++) hash = typeStr.charCodeAt(i) + ((hash << 5) - hash);
    return PASTEL_COLORS[Math.abs(hash) % PASTEL_COLORS.length].text;
}

function getBadgeStyle(typeStr) {
    const bg = getItemTypeChartColor(typeStr);
    const isLightBg = bg === '#F9C74F' || bg === '#90BE6D' || bg === '#E5E7EB';
    const text = isLightBg ? '#1F2937' : '#FFFFFF';
    return `background-color: ${bg}; color: ${text};`;
}

function renderCards(data, showAll = false) {
    DOM.recordCount.textContent = data.length;
    DOM.cardsContainer.innerHTML = '';

    if (data.length === 0) {
        DOM.noDataResult.style.display = 'flex';
    } else {
        DOM.noDataResult.style.display = 'none';

        const fragment = document.createDocumentFragment();
        
        // Sort descending by parsedTs (newest first)
        const sortedData = [...data].sort((a,b) => b.parsedTs - a.parsedTs);
        
        // Limit rendering to 20 unless showAll is true
        const limit = showAll ? sortedData.length : 20;
        const displayData = sortedData.slice(0, limit);

        displayData.forEach((item, index) => {
            const card = document.createElement('div');
            card.className = 'row-card';
            
            // Item Type mapping & Dynamic colors
            let displayItemType = item.processedItemType;
            let dynamicStyle = getBadgeStyle(displayItemType);

            if (displayItemType === 'สินค้าประเภทอื่น ๆ' && item.processedItemTypeOther) {
                displayItemType = `สินค้าประเภทอื่น ๆ<br><span style="font-size: 0.8em; font-weight: 500; opacity: 0.85;">👉 ${item.processedItemTypeOther}</span>`;
            }

            // Contract End Date Logic
            let endDateHTML = '';
            if (item.processedContractEndDate && item.processedContractEndDate !== 'ไม่ระบุ') {
                const parts = item.processedContractEndDate.split('/');
                if (parts.length === 3) {
                    const d = parseInt(parts[0]);
                    const m = parseInt(parts[1]) - 1;
                    const y = parseInt(parts[2]);
                    const endDt = new Date(y, m, d);
                    const now = new Date();
                    const diffDays = Math.ceil((endDt - now) / (1000 * 60 * 60 * 24));
                    
                    let badgeClass = '';
                    let icon = '<i data-lucide="calendar" style="width:14px;height:14px;"></i>';
                    if (diffDays < 0) {
                        badgeClass = 'expired';
                        icon = '<i data-lucide="alert-circle" style="width:14px;height:14px;"></i>';
                    } else if (diffDays <= 30) {
                        badgeClass = 'expiring-soon';
                        icon = '<i data-lucide="clock" style="width:14px;height:14px;"></i>';
                    }
                    endDateHTML = `<span class="enddate-badge ${badgeClass}">${icon}${item.processedContractEndDate}</span>`;
                } else {
                    endDateHTML = `<span class="enddate-badge"><i data-lucide="calendar" style="width:14px;height:14px;"></i>${item.processedContractEndDate}</span>`;
                }
            } else {
                endDateHTML = `<span style="color:var(--text-muted);font-size:0.8rem;">-</span>`;
            }

            // File Actions
            const hasFile = item.processedFile && item.processedFile !== '-';
            const actionHTML = hasFile
                ? `<a href="${item.processedFile}" target="_blank" class="action-btn" title="ดูไฟล์แนบ"><i data-lucide="paperclip"></i></a>`
                : `<span class="action-btn disabled" title="ไม่มีไฟล์"><i data-lucide="file-x"></i></span>`;

            card.innerHTML = `
                <div class="cell-client">
                    <div class="client-name" title="${item.processedCustomer}">${item.processedCustomer}</div>
                </div>
                
                <div class="cell-province">
                    <div class="cell-text">${item.processedProvince}</div>
                </div>
                
                <div class="cell-office">
                    <div class="cell-text">${item.processedPostOffice}</div>
                </div>

                <div class="cell-package">
                    <span class="pkg-badge" style="${getPkgBadgeStyle(item.processedPackage)}">${item.processedPackage}</span>
                </div>
                
                <div class="cell-type">
                    <span class="status-badge" style="${dynamicStyle}">${displayItemType}</span>
                </div>
                
                <div class="cell-enddate">
                    ${endDateHTML}
                </div>
                
                <div class="cell-action">
                    ${actionHTML}
                </div>
            `;
            fragment.appendChild(card);
        });
        
        // Load More Toggle Logic using the Top Button
        const topBtn = document.getElementById('showAllBtnTop');
        if (!showAll && sortedData.length > 20) {
            topBtn.style.display = 'inline-flex';
            topBtn.innerHTML = `ดูรายละเอียดทั้งหมด`;
            topBtn.onclick = () => renderCards(data, true);
        } else if (showAll && sortedData.length > 20) {
            topBtn.style.display = 'inline-flex';
            topBtn.innerHTML = `แสดงเฉพาะรายการล่าสุด (20 รายการ)`;
            topBtn.onclick = () => renderCards(data, false);
        } else {
            topBtn.style.display = 'none';
        }

        DOM.cardsContainer.appendChild(fragment);
        lucide.createIcons();
    }
}

// UI State Helpers
function showLoading(isLoading) {
    if (isLoading) {
        DOM.loading.style.display = 'flex';
        DOM.dashboardContent.style.display = 'none';
        DOM.errorMsg.style.display = 'none';
    } else {
        DOM.loading.style.display = 'none';
        DOM.dashboardContent.style.display = 'block';
    }
}

function showError(msg) {
    DOM.loading.style.display = 'none';
    DOM.dashboardContent.style.display = 'none';
    DOM.errorMsg.style.display = 'flex';
    DOM.errorText.textContent = msg;
}
