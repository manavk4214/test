#!/usr/bin/env bash
set -euo pipefail

APP_NAME="SCSPTSP"
TARGET_DIR="/home/ubuntu/${APP_NAME}"
DOMAIN="scsptsp.nielitdelhiforum.online"
PYTHON_BIN="python3"
VENV_DIR="${TARGET_DIR}/.venv"
SERVICE_FILE="/etc/systemd/system/${APP_NAME}.service"
NGINX_SITE="/etc/nginx/sites-available/${APP_NAME}"
NGINX_ENABLED="/etc/nginx/sites-enabled/${APP_NAME}"

if ! command -v apt-get >/dev/null 2>&1; then
  echo "This script is written for Ubuntu/Debian EC2 instances (apt-get required)."
  exit 1
fi

REQUIRED_PKGS=("$PYTHON_BIN" "${PYTHON_BIN}-venv" "${PYTHON_BIN}-pip" "nginx")
MISSING_PKGS=()
for pkg in "${REQUIRED_PKGS[@]}"; do
  if ! dpkg -s "$pkg" >/dev/null 2>&1; then
    MISSING_PKGS+=("$pkg")
  fi
done
if [ ${#MISSING_PKGS[@]} -gt 0 ]; then
  sudo apt-get update -y
  sudo apt-get install -y "${MISSING_PKGS[@]}"
fi

if [ ! -d "$TARGET_DIR" ]; then
  echo "Target directory doesn't exist. Cloning repository..."
  git clone "https://github.com/piyush1205nielit/SCSPTSP.git" "$TARGET_DIR"
fi

cd "$TARGET_DIR"
git fetch --all
git reset --hard origin/main

if [ ! -d "$VENV_DIR" ]; then
  "$PYTHON_BIN" -m venv "$VENV_DIR"
fi
source "$VENV_DIR/bin/activate"

python -m pip install --upgrade pip wheel
python -m pip install -r requirements.txt
python -m pip install gunicorn

if [ ! -f "${TARGET_DIR}/.env" ]; then
  cat > "${TARGET_DIR}/.env" <<EOF
DJANGO_ALLOWED_HOSTS=127.0.0.1,localhost,${DOMAIN}
DJANGO_DEBUG=False
DJANGO_SECRET_KEY=$(python -c "import secrets; print(secrets.token_urlsafe(50))")
EOF
fi

set -a; source "${TARGET_DIR}/.env"; set +a

python manage.py migrate --noinput

python manage.py collectstatic --noinput --clear

CURRENT_USER="${SUDO_USER:-$(whoami)}"
cat | sudo tee "$SERVICE_FILE" >/dev/null <<EOF
[Unit]
Description=${APP_NAME} Django service
After=network.target

[Service]
User=${CURRENT_USER}
Group=www-data
WorkingDirectory=${TARGET_DIR}
EnvironmentFile=${TARGET_DIR}/.env
ExecStart=${VENV_DIR}/bin/gunicorn --workers 3 --bind 127.0.0.1:8000 student.wsgi:application
Restart=always

[Install]
WantedBy=multi-user.target
EOF

sudo rm -f /etc/nginx/sites-enabled/SCSPTSP

CERT_DIR="/etc/letsencrypt/live/${DOMAIN}"
if [ -f "${CERT_DIR}/fullchain.pem" ]; then
  cat | sudo tee "$NGINX_SITE" >/dev/null <<EOF
server {
    listen 80;
    server_name ${DOMAIN};
    return 301 https://\$host\$request_uri;
}

server {
    listen 443 ssl;
    server_name ${DOMAIN};

    ssl_certificate ${CERT_DIR}/fullchain.pem;
    ssl_certificate_key ${CERT_DIR}/privkey.pem;
    include /etc/letsencrypt/options-ssl-nginx.conf;
    ssl_dhparam /etc/letsencrypt/ssl-dhparams.pem;

    location = /favicon.ico { access_log off; log_not_found off; }

    location /static/ {
        alias ${TARGET_DIR}/staticfiles/;
    }

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
EOF
else
  cat | sudo tee "$NGINX_SITE" >/dev/null <<EOF
server {
    listen 80;
    server_name ${DOMAIN};

    location = /favicon.ico { access_log off; log_not_found off; }

    location /static/ {
        alias ${TARGET_DIR}/staticfiles/;
    }

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
EOF
fi

sudo ln -sf "$NGINX_SITE" "$NGINX_ENABLED"
sudo rm -f /etc/nginx/sites-enabled/default

sudo systemctl daemon-reload
sudo systemctl enable "${APP_NAME}.service"
sudo systemctl restart "${APP_NAME}.service"
sudo nginx -t
sudo systemctl restart nginx

echo "Deployment complete."
