# SIAM JWD - Car Carrier Transport Optimization System
# รันเวอร์ชันแก้ไขกลุ่มล่าสุด (รายคัน + เลือกทั้งหมด)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

Write-Host "=== Car Carrier Grouping App ===" -ForegroundColor Cyan
Write-Host "โฟลเดอร์โปรเจกต์: $PSScriptRoot"

if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "ไม่พบ python ใน PATH กรุณาติดตั้ง Python 3 แล้วลองใหม่" -ForegroundColor Red
    exit 1
}

Write-Host "กำลังติดตั้งไลบรารีจาก requirements.txt ..."
python -m pip install -r "requirements.txt"

Write-Host "กำลังเปิด Streamlit (main.py) ..." -ForegroundColor Green
python -m streamlit run "main.py" --server.headless false
