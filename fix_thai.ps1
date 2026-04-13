$file = 'e:\Gravity\scratch\customer-files-dashboard\app.js'
$enc = New-Object System.Text.UTF8Encoding($false)  # UTF8 no BOM
$allLines = [System.IO.File]::ReadAllLines($file, $enc)

# Identify the range to replace: from "function updateDateFilterLabel" to end of "renderAIAlert"
$startIdx = -1
$endIdx = -1
for ($i = 0; $i -lt $allLines.Length; $i++) {
    if ($allLines[$i] -match 'function updateDateFilterLabel') { $startIdx = $i }
    if ($startIdx -gt -1 -and $i -gt $startIdx -and $allLines[$i] -match '^}$') { $endIdx = $i; break }
}
Write-Host "updateDateFilterLabel: lines $($startIdx+1) to $($endIdx+1)"

# Also find renderAIAlert end
$aiStart = -1; $aiEnd = -1
for ($i = 0; $i -lt $allLines.Length; $i++) {
    if ($allLines[$i] -match '// --- AI Alert Section ---') { $aiStart = $i }
}
if ($aiStart -gt -1) {
    for ($i = $aiStart; $i -lt $allLines.Length; $i++) {
        if ($i -gt $aiStart -and $allLines[$i] -match '^}$') { $aiEnd = $i }
        if ($i -gt $aiEnd -and $aiEnd -gt -1 -and $allLines[$i].Trim() -eq '') { break }
    }
}
Write-Host "renderAIAlert: from line $($aiStart+1) to $($aiEnd+1)"

# Build replacement text using unicode escapes only (no raw Thai chars in this script)
$vatTiSiSutSanYa = "'\u0E27\u0E31\u0E19\u0E17\u0E35\u0E48\u0E2A\u0E34\u0E49\u0E19\u0E2A\u0E38\u0E14\u0E2A\u0E31\u0E0D\u0E0D\u0E32\u0E2F'"
$luakStr = "'\u0E40\u0E25\u0E37\u0E2D\u0E01 ' + selectedEndDates.size + ' \u0E27\u0E31\u0E19\u0E17\u0E35\u0E48'"

$newBlock = @"
function updateDateFilterLabel() {
    var label = document.getElementById('dateFilterLabel');
    if (!label) return;
    if (selectedEndDates.size === 0) {
        label.textContent = $vatTiSiSutSanYa;
    } else if (selectedEndDates.size === 1) {
        label.textContent = [...selectedEndDates][0];
    } else {
        label.textContent = $luakStr;
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
    var thMonths = [
        '\u0E21\u0E01\u0E23\u0E32\u0E04\u0E21','\u0E01\u0E38\u0E21\u0E20\u0E32\u0E1E\u0E31\u0E19\u0E18\u0E4C','\u0E21\u0E35\u0E19\u0E32\u0E04\u0E21',
        '\u0E40\u0E21\u0E29\u0E32\u0E22\u0E19','\u0E1E\u0E24\u0E29\u0E20\u0E32\u0E04\u0E21','\u0E21\u0E34\u0E16\u0E38\u0E19\u0E32\u0E22\u0E19',
        '\u0E01\u0E23\u0E01\u0E0E\u0E32\u0E04\u0E21','\u0E2A\u0E34\u0E07\u0E2B\u0E32\u0E04\u0E21','\u0E01\u0E31\u0E19\u0E22\u0E32\u0E22\u0E19',
        '\u0E15\u0E38\u0E25\u0E32\u0E04\u0E21','\u0E1E\u0E24\u0E28\u0E08\u0E34\u0E01\u0E32\u0E22\u0E19','\u0E18\u0E31\u0E19\u0E27\u0E32\u0E04\u0E21'
    ];

    if (monthLabel) monthLabel.textContent = '\u0E1B\u0E23\u0E30\u0E08\u0E33\u0E40\u0E14\u0E37\u0E2D\u0E19 ' + thMonths[thisMonth] + ' ' + (thisYear + 543);

    var expiring = data.filter(function(item) {
        var d = item.processedContractEndDate;
        if (!d || d === '\u0E44\u0E21\u0E48\u0E23\u0E30\u0E1A\u0E38') return false;
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
        body.innerHTML = '<p class="ai-no-expiry">\u2705 \u0E40\u0E14\u0E37\u0E2D\u0E19\u0E19\u0E35\u0E49\u0E44\u0E21\u0E48\u0E21\u0E35\u0E25\u0E39\u0E01\u0E04\u0E49\u0E32\u0E17\u0E35\u0E48\u0E2A\u0E31\u0E0D\u0E0D\u0E32\u0E2A\u0E34\u0E49\u0E19\u0E2A\u0E38\u0E14 \u0E17\u0E38\u0E01\u0E2A\u0E31\u0E0D\u0E0D\u0E32\u0E22\u0E31\u0E07\u0E04\u0E07\u0E21\u0E35\u0E1C\u0E25\u0E2D\u0E22\u0E39\u0E48</p>';
        return;
    }

    var otherHTML = other.length > 0 ? '<br>\u00B7 \u0E2D\u0E37\u0E48\u0E19\u0E46: <strong>' + other.length + ' \u0E23\u0E32\u0E22</strong>' : '';
    var names = expiring.slice(0, 5).map(function(i) { return '\u2022 ' + i.processedCustomer; }).join('<br>');
    var moreText = expiring.length > 5 ? '<br>\u0E41\u0E25\u0E30\u0E2D\u0E35\u0E01 ' + (expiring.length - 5) + ' \u0E23\u0E32\u0E22\u0E01\u0E32\u0E23...' : '';

    body.innerHTML =
        '<div class="ai-alert-stat-card">' +
            '<div class="ai-stat-label">\u0E2A\u0E34\u0E49\u0E19\u0E2A\u0E38\u0E14\u0E40\u0E14\u0E37\u0E2D\u0E19\u0E19\u0E35\u0E49</div>' +
            '<div class="ai-stat-num" style="color:var(--primary)">' + expiring.length + '</div>' +
            '<div class="ai-stat-sub">\u0E23\u0E32\u0E22</div>' +
        '</div>' +
        '<div class="ai-alert-stat-card">' +
            '<div class="ai-stat-label">Package A11</div>' +
            '<div class="ai-stat-num" style="color:#E87A13">' + a11.length + '</div>' +
            '<div class="ai-stat-sub">\u0E23\u0E32\u0E22</div>' +
        '</div>' +
        '<div class="ai-alert-stat-card">' +
            '<div class="ai-stat-label">Package A12</div>' +
            '<div class="ai-stat-num" style="color:#2F8E6E">' + a12.length + '</div>' +
            '<div class="ai-stat-sub">\u0E23\u0E32\u0E22</div>' +
        '</div>' +
        '<div class="ai-alert-message">' +
            '\u{1F4CB} <strong>\u0E2A\u0E23\u0E38\u0E1B\u0E2A\u0E31\u0E0D\u0E0D\u0E32\u0E2A\u0E34\u0E49\u0E19\u0E2A\u0E38\u0E14\u0E40\u0E14\u0E37\u0E2D\u0E19' + thMonths[thisMonth] + '</strong><br>' +
            '\u0E21\u0E35\u0E25\u0E39\u0E01\u0E04\u0E49\u0E32\u0E17\u0E31\u0E49\u0E07\u0E2B\u0E21\u0E14 <strong>' + expiring.length + ' \u0E23\u0E32\u0E22</strong> \u0E17\u0E35\u0E48\u0E2A\u0E31\u0E0D\u0E0D\u0E32\u0E01\u0E33\u0E25\u0E31\u0E07\u0E2A\u0E34\u0E49\u0E19\u0E2A\u0E38\u0E14:<br>' +
            '\u00B7 Package A11: <strong>' + a11.length + ' \u0E23\u0E32\u0E22</strong><br>' +
            '\u00B7 Package A12: <strong>' + a12.length + ' \u0E23\u0E32\u0E22</strong>' + otherHTML + '<br><br>' +
            '\u{1F4CC} <strong>\u0E23\u0E32\u0E22\u0E0A\u0E37\u0E48\u0E2D:</strong><br>' + names + moreText +
        '</div>';
}
"@

$newLines = $newBlock -split "`r?`n"

# Replace lines from startIdx to aiEnd
$before = $allLines[0..($startIdx-1)]
$after  = $allLines[($aiEnd+1)..($allLines.Length - 1)]
$combined = $before + $newLines + $after

[System.IO.File]::WriteAllLines($file, $combined, $enc)
Write-Host "Fixed! Total lines: $($combined.Length)"
