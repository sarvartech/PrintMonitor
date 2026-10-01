# 🖨️ Canon Print Monitor — Professional Print & Syslog Monitor

Canon imageRUNNER ADVANCE C3922i va boshqa Canon printerlaridan chop etish loglarini yig'ish, tahlil qilish, Syslog orqali real vaqtda monitoring qilish hamda rasmiy **3 varaqli KPI Excel hisobotlarini** shakllantirish uchun mo'ljallangan to'liq avtomatlashtirilgan tizim.

---

## 🌟 Asosiy imkoniyatlar

- **📡 Real vaqtli Syslog Receiver (UDP:5140)**: Canon printerlaridan yuboriladigan chop etish hodisalarini lahzada qabul qilish va SQLite bazaga yozish.
- **⏰ Har 10 daqiqada Avto-Sinxronlash (Auto-Sync)**: Printerlarning Remote UI / CSV interfeysi orqali tarmoq uzilishlarida ham loglarni yo'qotmasdan avtomatik to'ldirib borish.
- **📊 Rasmiy 3 Varaqli KPI Hisoboti (`.xlsx`)**:
  - `KPI otchet <oy>`: 61 nafar faol xodimlar ro'yxati, chop etilgan varaqlar va hujjatlar soni, 220 limit va formulalar (`=G{row}/E{row}`).
  - `Сводка`: Umumiy oy jamlamasi, bajarilish ko'rsatkichlari, eng faol xodimlar, min/max/avg ko'rsatkichlar.
  - `Норма и список`: Xodimlar shtat jadvali va lavozimlari.
- **🎨 Zamonaviy Glassmorphism Dashboard**: Dark mode, real-vaqt interaktiv grafiklar (Chart.js), toner darajalari monitoringi (Cyan, Magenta, Yellow, Black) va ko'p tilli interfeys (UZ / RU).
- **🐳 Full Production Dockerization**: Docker, Docker Compose, Healthcheck va avtomatik deploy skriptlari.

---

## 🚀 Ishlab chiqarish (Production) serveriga o'rnatish

### Usul 1: Docker orqali avtomatik o'rnatish (Eng qulay va tavsiya etiladigan)

Serverga kirib bitta buyruq bilan o'rnatish:

```bash
# 1. Loyihani yuklab oling yoki serverga ko'chiring:
git clone https://github.com/USERNAME/PrintMonitor.git
cd PrintMonitor

# 2. Avtomatik deploy skriptini ishga tushiring:
chmod +x deploy_docker.sh
./deploy_docker.sh
```

Yoki to'g'ridan-to'g'ri `docker compose` orqali:
```bash
docker compose up -d --build
```

Konteyner holatini tekshirish:
```bash
docker compose ps
docker compose logs -f
```

---

### Usul 2: Nginx Reverse Proxy (80 / 443 port va Domen orqali)

Agar saytni standart HTTP 80 yoki HTTPS (SSL) orqali ochmoqchi bo'lsangiz:

1. Konfiguratsiyani Nginx ga ko'chiring:
```bash
sudo cp nginx.conf /etc/nginx/sites-available/print_monitor
sudo ln -s /etc/nginx/sites-available/print_monitor /etc/nginx/sites-enabled/
```

2. `/etc/nginx/sites-available/print_monitor` faylida `server_name` ga domeningizni yoki server IP sini yozing.

3. Nginx ni qayta ishga tushiring:
```bash
sudo nginx -t
sudo systemctl reload nginx
```

4. *(Ixtiyoriy)* Bepul SSL (Let's Encrypt) o'rnatish:
```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d print.sizningdomen.uz
```

---

### Usul 3: Linux Systemd servisi orqali (Docker-siz)

```bash
chmod +x deploy.sh
./deploy.sh
```

Servis boshqaruvi:
```bash
sudo systemctl status print_monitor
sudo systemctl restart print_monitor
sudo journalctl -u print_monitor -f
```

---

## 🖨️ Canon Printerda Syslog Sozlash

Printer Remote UI sahifasiga kiring (`http://<PRINTER_IP>:8000` yoki `http://<PRINTER_IP>`):
1. **Settings/Registration → Device Management → Export/Clear Audit Log → Syslog Settings**
2. **"Use Syslog Send"** ni yoqing (`✅`).
3. **Syslog Server Address**: Print Monitor o'rnatilgan server IP manzili.
4. **Syslog Server Port Number**: `5140`
5. **Connection Type**: `UDP`
6. Sozlamalarni saqlang va printerni qayta yuklang (agar talab qilinsa).

---

## 💾 Ma'lumotlar bazasi va zaxira nusxa (Backup)

Barcha loglar `canon_logs.db` SQLite faylida saqlanadi. Zaxira nusxa olish uchun:
```bash
# Bazani arxivlab olish:
sqlite3 canon_logs.db ".backup 'backup_$(date +%Y%m%d).db'"
# Yoki oddiy nusxa ko'chirish:
cp canon_logs.db "canon_logs_backup_$(date +%Y%m%d).db"
```

---

## 📁 Loyiha tarkibi

- `canon_server.py` — Flask asosidagi backend API, Syslog UDP qabul qiluvchi va avto-sinxronlash tizimi.
- `canon_dashboard.html` — Glassmorphism uslubidagi zamonaviy interaktiv frontend dashboard.
- `kpi_export.py` — Rasmiy 3 varaqli KPI Excel hisoboti generatori.
- `kpi_template.xlsx` — KPI andoza fayli.
- `Dockerfile` & `docker-compose.yml` — Production Docker konteyner konfiguratsiyasi.
- `deploy_docker.sh` — 1-klikli to'liq avtomatlashtirilgan Docker deploy skripti.
- `nginx.conf` — Production darajasidagi Nginx reverse proxy konfiguratsiyasi.
