#!/bin/bash

# Ensure the script is run as root
if [ "$EUID" -ne 0 ]; then
  echo "Please run this script as root (e.g., sudo ./install_linux_service.sh)"
  exit 1
fi

APP_DIR=$(pwd)
SERVICE_FILE="/etc/systemd/system/nexuscrm.service"
USER_NAME=$(logname || echo $SUDO_USER)

echo "Setting up NexusCRM service for directory: $APP_DIR"
echo "Running under user: $USER_NAME"

# Create the systemd service file
cat <<EOF > $SERVICE_FILE
[Unit]
Description=NexusCRM Backend Server
After=network.target

[Service]
User=$USER_NAME
WorkingDirectory=$APP_DIR
ExecStart=$APP_DIR/venv/bin/python3 $APP_DIR/app.py
Restart=always
RestartSec=5
# Output logs to standard journalctl
StandardOutput=syslog
StandardError=syslog
SyslogIdentifier=nexuscrm

[Install]
WantedBy=multi-user.target
EOF

# Reload systemd, enable the service to start on boot, and start it now
systemctl daemon-reload
systemctl enable nexuscrm.service
systemctl start nexuscrm.service

echo ""
echo "✅ NexusCRM has been successfully installed as a background service!"
echo "It will now start automatically every time the computer boots."
echo ""
echo "Commands to manage the service:"
echo "  Check Status : sudo systemctl status nexuscrm"
echo "  Stop App     : sudo systemctl stop nexuscrm"
echo "  Restart App  : sudo systemctl restart nexuscrm"
echo "  View Logs    : sudo journalctl -u nexuscrm -f"
