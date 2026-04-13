$file = 'e:\Gravity\scratch\customer-files-dashboard\app.js'
$enc = [System.Text.Encoding]::UTF8
$allLines = [System.IO.File]::ReadAllLines($file, $enc)

# Find the bad line (line index 274, which is line number 275)
$badIdx = 274

# Split the file into 3 parts: before bad line, after bad line (line 275+)
$before = $allLines[0..273]  # lines 1-274 (indices 0-273)
$after  = $allLines[275..($allLines.Length - 1)]  # line 276+ onwards

# The new code to insert between the two parts
$newCode = @'
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
    const sorted = [...dates].sort();
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
        item.appendChild(cb);
        item.appendChild(span);
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
        });
        container.appendChild(item);
    });
}

function updateDateFilterLabel() {
    const label = document.getElementById('dateFilterLabel');
    if (!label) return;
    if (selectedEndDates.size === 0) {
        label.textContent = 'วันที่สิ้นสุดสัญญาฯ';
    } else if (selectedEndDates.size === 1) {
        label.textContent = [...selectedEndDates][0];
    } else {
        label.textContent = 'เลือก ' + selectedEndDates.size + ' วันที่';
    }
}

// --- AI Alert Section ---
function renderAIAlert(data) {
    const body = document.getElementById('aiAlertBody');
    const monthLabel = document.getElementById('aiAlertMonthLabel');
    if (!body) return;

    const now = new Date();
    const thisMonth = now.getMonth();
    const thisYear = now.getFullYear();
    const thMonths = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                      'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

    if (monthLabel) monthLabel.textContent = 'ประจำเดือน ' + thMonths[thisMonth] + ' ' + (thisYear + 543);

    const expiring = data.filter(function(item) {
        const d = item.processedContractEndDate;
        if (!d || d === 'ไม่ระบุ') return false;
        const parts = d.split('/');
        if (parts.length !== 3) return false;
        const month = parseInt(parts[1]) - 1;
        const year = parseInt(parts[2]);
        return month === thisMonth && year === thisYear;
    });

    const a11 = expiring.filter(function(i) { return i.processedPackage && i.processedPackage.includes('A11'); });
    const a12 = expiring.filter(function(i) { return i.processedPackage && i.processedPackage.includes('A12'); });
    const other = expiring.filter(function(i) { return !(i.processedPackage && (i.processedPackage.includes('A11') || i.processedPackage.includes('A12'))); });

    if (expiring.length === 0) {
        body.innerHTML = '<p class="ai-no-expiry">✅ เดือนนี้ไม่มีลูกค้าที่สัญญาสิ้นสุด ทุกสัญญายังคงมีผลอยู่</p>';
        return;
    }

    const otherHTML = other.length > 0 ? '<br>· อื่นๆ: <strong>' + other.length + ' ราย</strong>' : '';
    const names = expiring.slice(0, 5).map(function(i) { return '• ' + i.processedCustomer; }).join('<br>');
    const moreText = expiring.length > 5 ? '<br>และอีก ' + (expiring.length - 5) + ' รายการ...' : '';

    body.innerHTML =
        '<div class="ai-alert-stat-card">' +
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
            '📋 <strong>สรุปสัญญาสิ้นสุดเดือน' + thMonths[thisMonth] + '</strong><br>' +
            'มีลูกค้าทั้งหมด <strong>' + expiring.length + ' ราย</strong> ที่สัญญากำลังสิ้นสุด:<br>' +
            '· Package A11: <strong>' + a11.length + ' ราย</strong><br>' +
            '· Package A12: <strong>' + a12.length + ' ราย</strong>' + otherHTML + '<br><br>' +
            '📌 <strong>รายชื่อ:</strong><br>' + names + moreText +
        '</div>';
}

'@

$newLines = $newCode -split "`n"
$combined = $before + $newLines + $after
[System.IO.File]::WriteAllLines($file, $combined, $enc)
Write-Host "Done. Total lines: $($combined.Length)"
