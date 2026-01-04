#!/bin/bash
# CCI Design VPS Setup Script for Ubuntu 24.04
# Run as: sudo bash setup-vps.sh

set -e

# =============================================================================
# CONFIGURATION
# =============================================================================
DEPLOY_USER="sysadmin"
DEPLOY_PATH="/home/${DEPLOY_USER}/cci-design"
LOG_DIR="${DEPLOY_PATH}/logs"
GITHUB_REPO="https://github.com/NatiCha/cci-design.git"

# =============================================================================
# COLORS FOR OUTPUT
# =============================================================================
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

log_info() { echo -e "${GREEN}[INFO]${NC} $1"; }
log_warn() { echo -e "${YELLOW}[WARN]${NC} $1"; }
log_error() { echo -e "${RED}[ERROR]${NC} $1"; }
log_step() { echo -e "${BLUE}[STEP]${NC} $1"; }

# =============================================================================
# CHECK PREREQUISITES
# =============================================================================
log_step "Checking prerequisites..."

if [ "$EUID" -ne 0 ]; then
    log_error "Please run as root (use sudo)"
    exit 1
fi

if [ ! -d "/home/${DEPLOY_USER}" ]; then
    log_error "User ${DEPLOY_USER} does not exist"
    exit 1
fi

# =============================================================================
# INSTALL SYSTEM PACKAGES
# =============================================================================
log_step "Installing system packages..."

apt-get update
apt-get install -y \
    git \
    curl \
    unzip \
    logrotate

# =============================================================================
# INSTALL UV (PYTHON PACKAGE MANAGER)
# =============================================================================
log_step "Installing uv package manager..."

if ! sudo -u ${DEPLOY_USER} bash -c 'command -v uv' &> /dev/null; then
    sudo -u ${DEPLOY_USER} bash -c 'curl -LsSf https://astral.sh/uv/install.sh | sh'
    log_info "uv installed for ${DEPLOY_USER}"
else
    log_info "uv already installed"
fi

# =============================================================================
# INSTALL CLOUDFLARED (Ubuntu 24.04 Noble)
# =============================================================================
log_step "Installing Cloudflare Tunnel (cloudflared)..."

if ! command -v cloudflared &> /dev/null; then
    # Add Cloudflare GPG key
    mkdir -p --mode=0755 /usr/share/keyrings
    curl -fsSL https://pkg.cloudflare.com/cloudflare-main.gpg | tee /usr/share/keyrings/cloudflare-main.gpg >/dev/null

    # Add Cloudflare repository for Ubuntu 24.04 (Noble)
    echo 'deb [signed-by=/usr/share/keyrings/cloudflare-main.gpg] https://pkg.cloudflare.com/cloudflared noble main' | tee /etc/apt/sources.list.d/cloudflared.list

    # Install cloudflared
    apt-get update
    apt-get install -y cloudflared
    log_info "cloudflared installed"
else
    log_info "cloudflared already installed"
fi

# =============================================================================
# CLONE/UPDATE REPOSITORY
# =============================================================================
log_step "Setting up repository..."

if [ ! -d "${DEPLOY_PATH}" ]; then
    log_info "Cloning repository..."
    sudo -u ${DEPLOY_USER} git clone ${GITHUB_REPO} ${DEPLOY_PATH}
else
    log_info "Repository already exists, pulling latest..."
    sudo -u ${DEPLOY_USER} bash -c "cd ${DEPLOY_PATH} && git pull origin main"
fi

# =============================================================================
# CREATE DIRECTORY STRUCTURE
# =============================================================================
log_step "Creating directory structure..."

sudo -u ${DEPLOY_USER} mkdir -p ${DEPLOY_PATH}/data/db
sudo -u ${DEPLOY_USER} mkdir -p ${DEPLOY_PATH}/data/templates
sudo -u ${DEPLOY_USER} mkdir -p ${DEPLOY_PATH}/output/reports/weekly
sudo -u ${DEPLOY_USER} mkdir -p ${DEPLOY_PATH}/output/reports/monthly
sudo -u ${DEPLOY_USER} mkdir -p ${DEPLOY_PATH}/output/invoices
sudo -u ${DEPLOY_USER} mkdir -p ${LOG_DIR}
sudo -u ${DEPLOY_USER} mkdir -p /home/${DEPLOY_USER}/.cloudflared

log_info "Directories created"

# =============================================================================
# INSTALL PYTHON DEPENDENCIES
# =============================================================================
log_step "Installing Python dependencies..."

sudo -u ${DEPLOY_USER} bash -c "source ~/.local/bin/env && cd ${DEPLOY_PATH} && uv sync"
log_info "Python dependencies installed"

# =============================================================================
# SETUP SYSTEMD SERVICES
# =============================================================================
log_step "Setting up systemd services..."

# CCI API Service
cp ${DEPLOY_PATH}/deploy/cci-api.service /etc/systemd/system/cci-api.service
log_info "cci-api.service installed"

# Cloudflared Service
cp ${DEPLOY_PATH}/deploy/cloudflared.service /etc/systemd/system/cloudflared.service
log_info "cloudflared.service installed"

# Reload systemd
systemctl daemon-reload
systemctl enable cci-api.service
log_info "cci-api.service enabled"

# =============================================================================
# SETUP CRON WRAPPER SCRIPTS
# =============================================================================
log_step "Setting up cron wrapper scripts..."

# Make wrapper scripts executable
chmod +x ${DEPLOY_PATH}/deploy/run-weekly-report.sh
chmod +x ${DEPLOY_PATH}/deploy/run-monthly-report.sh
chown ${DEPLOY_USER}:${DEPLOY_USER} ${DEPLOY_PATH}/deploy/run-weekly-report.sh
chown ${DEPLOY_USER}:${DEPLOY_USER} ${DEPLOY_PATH}/deploy/run-monthly-report.sh

log_info "Cron wrapper scripts configured"

# =============================================================================
# SETUP CRON JOBS
# =============================================================================
log_step "Setting up cron jobs..."

# Install crontab for deploy user
CRON_FILE="/tmp/cci-crontab"
cat > ${CRON_FILE} << 'CRONEOF'
# CCI Design Timesheet Reports
# Weekly report: Every Friday at 8PM (America/New_York)
0 20 * * 5 /home/sysadmin/cci-design/deploy/run-weekly-report.sh

# Monthly report: 1st of each month at 1AM (America/New_York)
0 1 1 * * /home/sysadmin/cci-design/deploy/run-monthly-report.sh
CRONEOF

sudo -u ${DEPLOY_USER} crontab ${CRON_FILE}
rm ${CRON_FILE}

log_info "Cron jobs installed"

# =============================================================================
# SETUP LOG ROTATION
# =============================================================================
log_step "Configuring log rotation..."

cat > /etc/logrotate.d/cci-design << LOGEOF
${LOG_DIR}/*.log {
    daily
    missingok
    rotate 14
    compress
    delaycompress
    notifempty
    create 0640 ${DEPLOY_USER} ${DEPLOY_USER}
    dateext
    dateformat -%Y%m%d
}
LOGEOF

log_info "Log rotation configured (14 day retention)"

# =============================================================================
# SETUP SUDOERS FOR DEPLOYMENT
# =============================================================================
log_step "Configuring sudoers for deployment..."

cp ${DEPLOY_PATH}/deploy/sudoers-cci /etc/sudoers.d/cci-deploy
chmod 440 /etc/sudoers.d/cci-deploy

# Validate sudoers file
if visudo -c -f /etc/sudoers.d/cci-deploy; then
    log_info "Sudoers configuration valid"
else
    log_error "Sudoers configuration invalid, removing..."
    rm /etc/sudoers.d/cci-deploy
    exit 1
fi

# =============================================================================
# GENERATE SSH DEPLOY KEY FOR GITHUB ACTIONS
# =============================================================================
log_step "Generating SSH deploy key..."

SSH_DIR="/home/${DEPLOY_USER}/.ssh"
SSH_KEY_PATH="${SSH_DIR}/github_deploy_key"

# Ensure .ssh directory exists
sudo -u ${DEPLOY_USER} mkdir -p ${SSH_DIR}
chmod 700 ${SSH_DIR}
chown ${DEPLOY_USER}:${DEPLOY_USER} ${SSH_DIR}

if [ ! -f "${SSH_KEY_PATH}" ]; then
    sudo -u ${DEPLOY_USER} ssh-keygen -t ed25519 -f ${SSH_KEY_PATH} -N "" -C "cci-deploy-key"
    log_info "Deploy key generated"
else
    log_info "Deploy key already exists"
fi

# Add deploy key to authorized_keys if not already there
AUTHORIZED_KEYS="${SSH_DIR}/authorized_keys"
if [ ! -f "${AUTHORIZED_KEYS}" ] || ! grep -q "cci-deploy-key" ${AUTHORIZED_KEYS} 2>/dev/null; then
    cat ${SSH_KEY_PATH}.pub >> ${AUTHORIZED_KEYS}
    chmod 600 ${AUTHORIZED_KEYS}
    chown ${DEPLOY_USER}:${DEPLOY_USER} ${AUTHORIZED_KEYS}
    log_info "Deploy key added to authorized_keys"
fi

# =============================================================================
# FINAL INSTRUCTIONS
# =============================================================================
echo ""
echo "=============================================="
log_info "VPS Setup Complete!"
echo "=============================================="
echo ""
log_warn "MANUAL STEPS REQUIRED:"
echo ""
echo "1. Create .env file:"
echo "   nano ${DEPLOY_PATH}/.env"
echo ""
echo "   Add these variables:"
echo "   MICROSOFT_GRAPH_TENANT_ID=your-tenant-id"
echo "   MICROSOFT_GRAPH_APP_ID=your-app-id"
echo "   MICROSOFT_GRAPH_CLIENT_SECRET=your-client-secret"
echo "   CCI_API_KEY=\$(openssl rand -base64 32)"
echo ""
echo "2. Copy invoice template to server:"
echo "   scp data/templates/invoice-template.xlsx sysadmin@<server>:${DEPLOY_PATH}/data/templates/"
echo ""
echo "3. Initialize database:"
echo "   cd ${DEPLOY_PATH}"
echo "   source ~/.local/bin/env"
echo "   uv run python src/scripts/init_db.py"
echo ""
echo "4. Configure Cloudflare Tunnel:"
echo "   cloudflared tunnel login"
echo "   cloudflared tunnel create cci-api"
echo "   # Copy the Tunnel ID, then edit ~/.cloudflared/config.yml"
echo "   cp ${DEPLOY_PATH}/deploy/cloudflared-config.yml.template ~/.cloudflared/config.yml"
echo "   nano ~/.cloudflared/config.yml  # Replace <TUNNEL_ID> with actual ID"
echo "   cloudflared tunnel route dns cci-api cci.landslidelogic.com"
echo ""
echo "5. Start services:"
echo "   sudo systemctl start cci-api"
echo "   sudo systemctl enable cloudflared"
echo "   sudo systemctl start cloudflared"
echo ""
echo "6. Add GitHub Secrets (Settings > Secrets > Actions):"
echo "   SSH_HOST: your VPS IP address"
echo "   SSH_USER: ${DEPLOY_USER}"
echo "   SSH_PRIVATE_KEY: (see below)"
echo ""
echo "=============================================="
log_info "SSH Private Key for GitHub Actions:"
echo "=============================================="
echo ""
cat ${SSH_KEY_PATH}
echo ""
echo "=============================================="
log_info "SSH Public Key (already in authorized_keys):"
echo "=============================================="
echo ""
cat ${SSH_KEY_PATH}.pub
echo ""
