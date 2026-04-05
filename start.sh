#!/bin/bash
# 啟動出勤查詢系統（Sheets Proxy + n8n）
# 使用方式：bash start.sh

cd "$(dirname "$0")"

echo "🔄 停止舊程序..."
pkill -f "sheets-proxy" 2>/dev/null
pkill -f "n8n start" 2>/dev/null
sleep 2

echo "🚀 啟動 Sheets Proxy（port 3456）..."
GMAIL_APP_PASSWORD="hhdgvyhdalmwmsqd" node sheets-proxy.js &
sleep 2

echo "🚀 啟動 n8n（port 5678）..."
N8N_DEFAULT_LOCALE=zh n8n start &
sleep 5

echo ""
echo "✅ 系統已啟動！"
echo "   n8n:   http://localhost:5678"
echo "   Proxy: http://localhost:3456"
echo ""
echo "📋 測試指令："
echo '   curl -X POST http://localhost:5678/webhook/attendance-query \'
echo '     -H "Content-Type: application/json" \'
echo '     -d '"'"'{"query": "4/2請假狀況"}'"'"''
echo ""
echo "按 Ctrl+C 停止所有服務"
wait
