# CCI Timesheet Reporting System - Deployment Guide

This document describes how to deploy the CCI Timesheet Reporting System to an Ubuntu 24.04 VPS.

## Overview

The deployment consists of:
- **FastAPI** running as a systemd service via uvicorn
- **Cloudflare Tunnel** exposing the API
- **Cron jobs** for automated weekly and monthly report generation
- **GitHub Actions** for CI/CD deployment on push to main

## Prerequisites

- Ubuntu 24.04.3 LTS VPS
- User `sysadmin` with sudo access
- Python 3.12+ (pre-installed on Ubuntu 24.04)
- Cloudflare account with domain configured
- GitHub repository access

## Quick Start

### 1. Initial Server Setup

SSH into your VPS and run:

```bash
# Clone repository
cd /home/sysadmin
git clone https://github.com/NatiCha/cci-design.git
cd cci-design

# Run setup script
sudo bash deploy/setup-vps.sh
```

The setup script will:
- Install required packages (git, curl, unzip, logrotate)
- Install uv (Python package manager)
- Install cloudflared (Cloudflare Tunnel)
- Create directory structure
- Install Python dependencies
- Set up systemd services
- Configure cron jobs
- Set up log rotation
- Generate SSH deploy key

### 2. Configure Environment Variables

Create the `.env` file (this file is never overwritten by deployments):

```bash
nano /home/sysadmin/cci-design/.env
```

Add the following variables:

```env
# Azure/Microsoft Entra authentication
MICROSOFT_GRAPH_TENANT_ID=your-tenant-id
MICROSOFT_GRAPH_APP_ID=your-app-id
MICROSOFT_GRAPH_CLIENT_SECRET=your-client-secret

# API authentication
CCI_API_KEY=your-secure-api-key
```

Generate a strong API key:
```bash
openssl rand -base64 32
```

### 3. Copy Invoice Template

The invoice template must be copied manually (it's gitignored):

```bash
# From your local machine
scp data/templates/invoice-template.xlsx sysadmin@<server>:/home/sysadmin/cci-design/data/templates/
scp data/templates/cci-design-logo.png sysadmin@<server>:/home/sysadmin/cci-design/data/templates/
```

### 4. Initialize Database

```bash
cd /home/sysadmin/cci-design
source ~/.local/bin/env
uv run python src/scripts/init_db.py
```

### 5. Configure Cloudflare Tunnel

#### a. Login to Cloudflare

```bash
cloudflared tunnel login
```

This opens a browser for authentication.

#### b. Create Tunnel

```bash
cloudflared tunnel create cci-api
```

Note the Tunnel ID from the output (e.g., `a1b2c3d4-e5f6-7890-abcd-ef1234567890`).

#### c. Configure Tunnel

Copy the template and edit:

```bash
cp /home/sysadmin/cci-design/deploy/cloudflared-config.yml.template ~/.cloudflared/config.yml
nano ~/.cloudflared/config.yml
```

Replace `<TUNNEL_ID>` with your actual Tunnel ID:

```yaml
tunnel: a1b2c3d4-e5f6-7890-abcd-ef1234567890
credentials-file: /home/sysadmin/.cloudflared/a1b2c3d4-e5f6-7890-abcd-ef1234567890.json

ingress:
  - hostname: cci.landslidelogic.com
    service: http://localhost:8000
  - service: http_status:404
```

#### d. Configure DNS

```bash
cloudflared tunnel route dns cci-api cci.landslidelogic.com
```

#### e. Enable and Start Tunnel Service

```bash
sudo systemctl enable cloudflared
sudo systemctl start cloudflared
```

### 6. Start the API Service

```bash
sudo systemctl start cci-api
sudo systemctl status cci-api
```

### 7. Configure GitHub Actions

Add these secrets to your GitHub repository (Settings > Secrets and variables > Actions):

| Secret | Description |
|--------|-------------|
| `SSH_HOST` | Your VPS IP address |
| `SSH_USER` | `sysadmin` |
| `SSH_PRIVATE_KEY` | Contents of `/home/sysadmin/.ssh/github_deploy_key` |

The setup script outputs the private key content. If you need it again:

```bash
cat ~/.ssh/github_deploy_key
```

## Directory Structure

```
/home/sysadmin/cci-design/
├── .env                    # Environment variables (manual, never overwritten)
├── data/
│   ├── db/
│   │   └── cci-timesheets.db   # SQLite database
│   └── templates/
│       ├── invoice-template.xlsx # Invoice template (manual copy)
│       └── cci-design-logo.png   # Logo for invoices (manual copy)
├── output/
│   ├── reports/
│   │   ├── weekly/         # Generated weekly reports (.numbers)
│   │   └── monthly/        # Generated monthly reports (.xlsx)
│   └── invoices/           # Generated invoices (.xlsx)
├── logs/
│   ├── api.log             # API stdout
│   ├── api-error.log       # API stderr
│   ├── weekly-report.log   # Weekly cron output
│   └── monthly-report.log  # Monthly cron output
├── deploy/
│   ├── setup-vps.sh
│   ├── cci-api.service
│   ├── cloudflared.service
│   ├── cloudflared-config.yml.template
│   ├── run-weekly-report.sh
│   ├── run-monthly-report.sh
│   └── sudoers-cci
└── src/                    # Application source code
```

## Cron Jobs

Cron jobs run under the `sysadmin` user:

| Schedule | Script | Description |
|----------|--------|-------------|
| Friday 8PM | `deploy/run-weekly-report.sh` | Weekly timesheet report |
| 1st of month 1AM | `deploy/run-monthly-report.sh` | Monthly timesheet report |

### View Cron Configuration

```bash
crontab -l
```

### Manual Execution

```bash
# Run weekly report manually
/home/sysadmin/cci-design/deploy/run-weekly-report.sh

# Run monthly report manually
/home/sysadmin/cci-design/deploy/run-monthly-report.sh
```

## Service Management

### CCI API Service

```bash
# Start service
sudo systemctl start cci-api

# Stop service
sudo systemctl stop cci-api

# Restart service
sudo systemctl restart cci-api

# View status
sudo systemctl status cci-api

# View logs (live)
sudo journalctl -u cci-api -f

# View log file
tail -f /home/sysadmin/cci-design/logs/api.log
```

### Cloudflare Tunnel Service

```bash
# Start tunnel
sudo systemctl start cloudflared

# Stop tunnel
sudo systemctl stop cloudflared

# Restart tunnel
sudo systemctl restart cloudflared

# View status
sudo systemctl status cloudflared

# View logs
sudo journalctl -u cloudflared -f
```

## Log Rotation

Logs are rotated daily with 14 days retention.

Configuration file: `/etc/logrotate.d/cci-design`

## Troubleshooting

### API Not Starting

1. Check service status:
   ```bash
   sudo systemctl status cci-api
   ```

2. Check logs:
   ```bash
   sudo journalctl -u cci-api -n 50
   cat /home/sysadmin/cci-design/logs/api-error.log
   ```

3. Verify .env file exists and has correct permissions:
   ```bash
   ls -la /home/sysadmin/cci-design/.env
   ```

4. Test manually:
   ```bash
   cd /home/sysadmin/cci-design
   source ~/.local/bin/env
   uv run uvicorn src.api.main:app --host 0.0.0.0 --port 8000
   ```

### Cloudflare Tunnel Issues

1. Check tunnel status:
   ```bash
   cloudflared tunnel info cci-api
   ```

2. Test tunnel locally:
   ```bash
   cloudflared tunnel --config ~/.cloudflared/config.yml run
   ```

3. Verify DNS routing:
   ```bash
   cloudflared tunnel route dns cci-api cci.landslidelogic.com
   ```

4. Check tunnel logs:
   ```bash
   sudo journalctl -u cloudflared -n 50
   ```

### Database Issues

1. Verify database exists:
   ```bash
   ls -la /home/sysadmin/cci-design/data/db/
   ```

2. Reinitialize if needed:
   ```bash
   cd /home/sysadmin/cci-design
   source ~/.local/bin/env
   uv run python src/scripts/init_db.py
   ```

### Cron Jobs Not Running

1. Check cron service:
   ```bash
   sudo systemctl status cron
   ```

2. Check cron logs:
   ```bash
   grep CRON /var/log/syslog | tail -20
   ```

3. Verify crontab:
   ```bash
   crontab -l
   ```

4. Check script permissions:
   ```bash
   ls -la /home/sysadmin/cci-design/deploy/run-*.sh
   ```

5. Check script logs:
   ```bash
   tail -50 /home/sysadmin/cci-design/logs/weekly-report.log
   tail -50 /home/sysadmin/cci-design/logs/monthly-report.log
   ```

### Deployment Failures

1. Check GitHub Actions logs in the repository

2. Verify SSH access:
   ```bash
   ssh -i ~/.ssh/github_deploy_key sysadmin@<server>
   ```

3. Check sudoers configuration:
   ```bash
   sudo cat /etc/sudoers.d/cci-deploy
   ```

## Security Considerations

1. **Environment Variables**: The `.env` file contains secrets and should have restricted permissions:
   ```bash
   chmod 600 /home/sysadmin/cci-design/.env
   ```

2. **API Key**: Use a strong, randomly generated API key

3. **Firewall**: No ports need to be opened - Cloudflare Tunnel handles ingress

4. **SSH Keys**: Deploy keys are ed25519 for security

5. **Service Hardening**: The systemd service runs with `NoNewPrivileges=true` and `PrivateTmp=true`

## Backup Recommendations

1. **Database**: Backup `/home/sysadmin/cci-design/data/db/cci-timesheets.db`

2. **Environment**: Backup `/home/sysadmin/cci-design/.env`

3. **Cloudflare Credentials**: Backup `~/.cloudflared/` directory

4. **Generated Reports**: Consider periodic backup of `/home/sysadmin/cci-design/output/`

## Manual Deployment

If you need to deploy manually without GitHub Actions:

```bash
cd /home/sysadmin/cci-design
git pull origin main
source ~/.local/bin/env
uv sync
sudo systemctl restart cci-api
```

## System Updates

```bash
# Update system packages
sudo apt update && sudo apt upgrade -y

# Update uv
curl -LsSf https://astral.sh/uv/install.sh | sh

# Update cloudflared
sudo apt update && sudo apt install --only-upgrade cloudflared
```
