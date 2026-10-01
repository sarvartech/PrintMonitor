#!/bin/bash
# ====================================================================
# Print Monitor — Docker Production Auto-Deploy Script
# Ishlatish: chmod +x deploy_docker.sh && ./deploy_docker.sh
# ====================================================================

set -e

echo "=========================================================="
echo "🖨️  Print Monitor Docker Deploy boshlanmoqda..."
echo "=========================================================="

# 1. Docker o'rnatilganligini tekshirish
if ! command -v docker &> /dev/null; then
    echo "⚠️ Docker topilmadi. Docker va kerakli paketlar o'rnatilmoqda..."
    sudo apt-get update
    sudo apt-get install -y ca-certificates curl gnupg lsb-release
    sudo install -m 0755 -d /etc/apt/keyrings
    if [ ! -f /etc/apt/keyrings/docker.gpg ]; then
        curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
        sudo chmod a+r /etc/apt/keyrings/docker.gpg
    fi
    echo \
      "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] https://download.docker.com/linux/ubuntu \
      $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
    sudo apt-get update
    sudo apt-get install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin docker-compose-v2
    sudo usermod -aG docker "$USER"
    echo "✅ Docker muvaffaqiyatli o'rnatildi!"
fi

# 2. Docker compose komandasini aniqlash
COMPOSE_CMD=""
if docker compose version &> /dev/null; then
    COMPOSE_CMD="docker compose"
elif command -v docker-compose &> /dev/null; then
    COMPOSE_CMD="docker-compose"
else
    echo "⚠️ Docker Compose o'rnatilmoqda..."
    sudo apt-get update && sudo apt-get install -y docker-compose-v2 docker-compose
    COMPOSE_CMD="docker compose"
fi

# 3. Fayllar mavjudligini ta'minlash (Docker papka qilib yaratib yubormasligi uchun)
if [ ! -f canon_logs.db ]; then
    echo "ℹ️ canon_logs.db yaratilmoqda..."
    touch canon_logs.db
fi

if [ ! -f toner_cache.json ]; then
    echo "ℹ️ toner_cache.json yaratilmoqda..."
    echo "{}" > toner_cache.json
fi

# 4. Firewall sozlash (agar UFW yoqilgan bo'lsa)
if command -v ufw &> /dev/null; then
    UFW_STATUS=$(sudo ufw status | head -n 1)
    if [[ "$UFW_STATUS" == *"active"* ]]; then
        echo "🛡️ Firewall (UFW) portlari ochilmoqda..."
        sudo ufw allow 5000/tcp comment 'Print Monitor Web Dashboard'
        sudo ufw allow 5140/udp comment 'Print Monitor Syslog UDP'
    fi
fi

# 5. Konteynerni to'xtatish (agar avval ishlayotgan bo'lsa) va qayta build qilish
echo "📦 Docker konteyneri yaratilmoqda va ishga tushirilmoqda..."
$COMPOSE_CMD down 2>/dev/null || true
$COMPOSE_CMD up -d --build

# 6. Salomatlik tekshiruvi (Healthcheck)
echo "⏳ Tizim ishga tushishi kutilmoqda..."
sleep 5
for i in {1..12}; do
    if curl -s http://localhost:5000/api/health > /dev/null 2>&1; then
        echo "✅ Print Monitor muvaffaqiyatli ishga tushdi va ishlamoqda!"
        break
    fi
    echo -n "."
    sleep 3
done

SERVER_IP=$(hostname -I 2>/dev/null | awk '{print $1}')
if [ -z "$SERVER_IP" ]; then
    SERVER_IP="<server_ip>"
fi

echo ""
echo "=========================================================="
echo "🎉 O'rnatish muvaffaqiyatli yakunlandi!"
echo "🌐 Veb Dashboard: http://${SERVER_IP}:5000"
echo "📡 Syslog UDP:    Port 5140"
echo "=========================================================="
echo "Konteyner holatini tekshirish: $COMPOSE_CMD ps"
echo "Loglarni ko'rish:             $COMPOSE_CMD logs -f"
echo "=========================================================="
