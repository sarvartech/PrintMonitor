#!/bin/bash
# Print Monitor - Serverda avtomatik sozlash va ishga tushirish skripti
# Ishlatish: chmod +x deploy.sh && ./deploy.sh

set -e

echo "=== System yangilanmoqda va kerakli paketlar o'rnatilmoqda ==="
sudo apt update
sudo apt install -y python3-pip python3-venv unzip nginx

# Virtual muhit (venv) yaratish va sozlash
if [ ! -d "venv" ]; then
    echo "=== Virtual muhit yaratilmoqda (venv) ==="
    python3 -m venv venv
fi

echo "=== Kutubxonalar o'rnatilmoqda ==="
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

# Hozirgi katalogni aniqlash
APP_DIR=$(pwd)

echo "=== Systemd Service yaratilmoqda ==="
sudo tee /etc/systemd/system/print_monitor.service > /dev/null <<EOF
[Unit]
Description=Canon Print Monitor Service
After=network.target

[Service]
User=$USER
WorkingDirectory=$APP_DIR
ExecStart=$APP_DIR/venv/bin/python canon_server.py
Restart=always
RestartSec=5
StandardOutput=syslog
StandardError=syslog
SyslogIdentifier=print_monitor

[Install]
WantedBy=multi-user.target
EOF

echo "=== Tizim xizmatlari qayta yuklanmoqda va ishga tushirilmoqda ==="
sudo systemctl daemon-reload
sudo systemctl enable print_monitor.service
sudo systemctl restart print_monitor.service

# Firewall sozlash (agar ufw yoqilgan bo'lsa)
if command -v ufw > /dev/null; then
    echo "=== Firewall sozlanmoqda (Port 5000 va UDP 5140) ==="
    sudo ufw allow 5000/tcp comment 'Print Monitor Web Dashboard'
    sudo ufw allow 5140/udp comment 'Print Monitor Syslog Receiver'
fi

echo "=== Xizmat holati tekshirilmoqda ==="
sudo systemctl status print_monitor.service --no-pager

echo ""
echo "=========================================================="
echo " Muvaffaqiyatli yakunlandi!"
echo " Web Dashboard porti: http://<server_ip>:5000"
echo " Syslog UDP porti:    5140"
echo "=========================================================="
