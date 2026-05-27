# Porra Mundial 2026

## Instalación local
```bash
pip3 install flask openpyxl
python3 app.py
```

## Variables de entorno (Render)
- `ADMIN_USER` — usuario del panel admin (por defecto: admin)
- `ADMIN_PASS` — contraseña del panel admin
- `SECRET_KEY` — clave secreta para sesiones (genera una aleatoria)

## Rutas
- `/` — Clasificación pública
- `/admin` — Panel de administración (requiere login)
