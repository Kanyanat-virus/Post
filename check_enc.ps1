$file = 'e:\Gravity\scratch\customer-files-dashboard\app.js'
$content = Get-Content $file -Raw -Encoding UTF8

# Fix Thai encoding issue if present (re-read with correct encoding)
$bytes = [System.IO.File]::ReadAllBytes($file)
$text = [System.Text.Encoding]::UTF8.GetString($bytes)

# Check if bad encoding by looking for replacement chars
if ($text -match '\xEF\xBF\xBD' -or $text -match 'เธง') {
    Write-Host "Encoding issue detected - file may have wrong encoding"
} else {
    Write-Host "Encoding OK"
}
Write-Host "File size: $($bytes.Length) bytes"
Write-Host "First Thai check line 341: " ($text.Split("`n")[340].Substring(0, [Math]::Min(80, $text.Split("`n")[340].Length)))
