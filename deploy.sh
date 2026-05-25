#!/usr/bin/env bash
set -euo pipefail

APP_NAME="scsptsp"
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_BIN="python3"
VENV_DIR="$APP_DIR/.venv"
SERVICE_FILE="/etc/systemd/system/${APP_NAME}.service"
NGINX_SITE="/etc/nginx/sites-available/${APP_NAME}"
NGINX_ENABLED="/etc/nginx/sites-enabled/${APP_NAME}"
DJANGO_PORT="${DJANGO_PORT:-8000}"
DJANGO_ALLOWED_HOSTS="${DJANGO_ALLOWED_HOSTS:-127.0.0.1,localhost}"
DJANGO_DEBUG="${DJANGO_DEBUG:-False}"
DJANGO_SECRET_KEY="${DJANGO_SECRET_KEY:-change-me-in-production}"

if ! command -v apt-get >/dev/null 2>&1; then
  echo "This script is written for Ubuntu/Debian EC2 instances (apt-get required)."
  exit 1
fi

cd "$APP_DIR"

# Install System dependencies
sudo apt-get update -y
sudo apt-get install -y "$PYTHON_BIN" "${PYTHON_BIN}-venv" "${PYTHON_BIN}-pip" nginx

# Setup Python Virtual Environment
"$PYTHON_BIN" -m venv "$VENV_DIR"
source "$VENV_DIR/bin/activate"

python -m pip install --upgrade pip wheel

# Clean windows line endings if present in requirements.txt
TMP_REQUIREMENTS="$(mktemp)"
tr -d '\000' < requirements.txt | sed 's/\r$//' > "$TMP_REQUIREMENTS"
python -m pip install -r "$TMP_REQUIREMENTS"
python -m pip install gunicorn
rm -f "$TMP_REQUIREMENTS"

# Write Production Environment File
cat > "$APP_DIR/.env" <<EOF
DJANGO_ALLOWED_HOSTS=${DJANGO_ALLOWED_HOSTS}
DJANGO_DEBUG=${DJANGO_DEBUG}
DJANGO_SECRET_KEY=${DJANGO_SECRET_KEY}
EOF

export DJANGO_ALLOWED_HOSTS DJANGO_DEBUG DJANGO_SECRET_KEY

# Run migrations
python manage.py migrate --noinput

# Manage Static Files
rm -rf "$APP_DIR/staticfiles"
mkdir -p "$APP_DIR/staticfiles"
if [ -f "$APP_DIR/static/index.css" ]; then
  cp -f "$APP_DIR/static/index.css" "$APP_DIR/staticfiles/index.css"
fi

LOGO_FILE="$(find "$APP_DIR/static" -maxdepth 1 -type f -iname 'logo.*' | head -n 1 || true)"
if [ -n "$LOGO_FILE" ]; then
  cp -f "$LOGO_FILE" "$APP_DIR/staticfiles/"
fi

find "$APP_DIR/staticfiles" -type f ! -iname 'index.css' ! -iname 'logo.*' -delete

# Create Systemd Service File (Fix: Ensure it runs as the actual user, not root)
CURRENT_USER="${SUDO_USER:-$(whoami)}"
cat | sudo tee "$SERVICE_FILE" >/dev/null <<EOF
[Unit]
Description=${APP_NAME} Django service
After=network.target

[Service]
User=${CURRENT_USER}
Group=www-data
WorkingDirectory=${APP_DIR}
EnvironmentFile=${APP_DIR}/.env
ExecStart=${VENV_DIR}/bin/gunicorn --workers 3 --bind 127.0.0.1:${DJANGO_PORT} student.wsgi:application
Restart=always

[Install]
WantedBy=multi-user.target
EOF

# Setup Nginx configuration
cat | sudo tee "$NGINX_SITE" >/dev/null <<EOF
server {
    listen 80;
    server_name _;

    location = /favicon.ico { access_log off; log_not_found off; }

    location /static/ {
        alias ${APP_DIR}/staticfiles/;
    }

    location / {
        proxy_pass http://127.0.0.1:${DJANGO_PORT};
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
EOF

# Link Nginx sites and disable default Apache/Nginx welcoming page
sudo ln -sf "$NGINX_SITE" "$NGINX_ENABLED"
sudo rm -f /etc/nginx/sites-enabled/default

# Restart processes to apply updates
sudo systemctl daemon-reload
sudo systemctl enable "${APP_NAME}.service"
sudo systemctl restart "${APP_NAME}.service"
sudo nginx -t
sudo systemctl restart nginx

echo "Deployment prep complete."
