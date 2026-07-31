# 🖨️ Canon Print Monitor — Professional Print & Syslog Monitor

Canon imageRUNNER ADVANCE C3922i va boshqa Canon printerlaridan chop etish loglarini yig'ish, tahlil qilish, Syslog orqali real vaqtda monitoring qilish hamda rasmiy **3 varaqli KPI Excel hisobotlarini** shakllantirish uchun mo'ljallangan avtomatlashtirilgan tizim.

---

## 🌟 Asosiy imkoniyatlar

- **📡 Real vaqtli Syslog Receiver (UDP:5140)**: Canon printerlaridan jo'natiladigan chop etish hodisalarini lahzada qabul qilish va bazaga yozish.
- **⏰ Har 10 minutda Avto-Sinxronlash (Auto-Sync)**: Printerlarning Remote UI / CSV interfeysi orqali o'tkazib yuborilgan loglarni avtomatik tortib kelish.
- **📊 Rasmiy 3 Varaqli KPI Hisoboti (`.xlsx`)**:
  - `KPI otchet <oy>`: 61 nafar faol xodimlar ro'yxati, chop etilgan varaqlar va hujjatlar soni, 220 limit va formulalar (`=G{row}/E{row}`).
  - `Сводка`: Umumiy oy jamlamasi, bajarilish ko'rsatkichlari, max/avg ko'rsatkichlari.
  - `Норма и список`: Xodimlar shtat jadvali.
- **🎨 Modern Glassmorphism Dashboard**: Dark mode, real-vaqt grafiklar (Chart.js), toner holati monitoringi va ko'p tilli interfeys (UZ / RU).
- **🐳 Full Dockerization**: Docker & Docker Compose orqali 1 daqiqada ishga tushirish.

---

## 🐳 Docker orqali ishga tushirish (Tavsiya etiladi)

### 1. Omborni klonlash
```bash
git clone https://github.com/USERNAME/PrintMonitor.git
cd PrintMonitor
```

### 2. Docker Compose orqali konteynerni ishga tushirish
```bash
docker compose up -d --build
```

Dastur quyidagi portlarda ishga tushadi:
- **Veb Dashboard**: `http://<server_ip>:5000`
- **Syslog UDP Receiver**: `5140/udp`

---

## 🛠 Manual (Servissiz) Ishga Tushirish

```bash
# Virtual muhit yaratish va kutubxonalarni o'rnatish
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Serverni ishga tushirish
python canon_server.py
```

Yoki avtomatik Linux systemd servisi o'rnatish uchun:
```bash
chmod +x deploy.sh
./deploy.sh
```

---

## 🖨️ Canon Printerda Syslog Sozlash

Printer Remote UI sahifasiga kiring (`http://<PRINTER_IP>:8000`):
1. **Settings/Registration → Device Management → Export/Clear Audit Log → Syslog Settings**
2. **"Use Syslog Send"** ni yoqing (`✅`).
3. **Syslog Server Address**: Server IP manzili.
4. **Syslog Server Port Number**: `5140`
5. **Connection Type**: `UDP`
