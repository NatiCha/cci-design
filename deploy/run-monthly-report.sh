#!/bin/bash
# Monthly Report Cron Wrapper
# Runs on the 1st of each month at 1AM via crontab
# Logs to: /home/sysadmin/cci-design/logs/monthly-report.log

set -e

LOG_FILE="/home/sysadmin/cci-design/logs/monthly-report.log"

echo "========================================" >> "$LOG_FILE"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Starting monthly report" >> "$LOG_FILE"

cd /home/sysadmin/cci-design
source ~/.local/bin/env

uv run python src/scripts/create_monthly_report.py >> "$LOG_FILE" 2>&1
EXIT_CODE=$?

if [ $EXIT_CODE -eq 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Monthly report completed successfully" >> "$LOG_FILE"
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Monthly report failed with exit code $EXIT_CODE" >> "$LOG_FILE"
fi

echo "========================================" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"

exit $EXIT_CODE
