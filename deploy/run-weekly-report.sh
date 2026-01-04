#!/bin/bash
# Weekly Report Cron Wrapper
# Runs every Friday at 8PM via crontab
# Logs to: /home/sysadmin/cci-design/logs/weekly-report.log

set -e

LOG_FILE="/home/sysadmin/cci-design/logs/weekly-report.log"

echo "========================================" >> "$LOG_FILE"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] Starting weekly report" >> "$LOG_FILE"

cd /home/sysadmin/cci-design
source ~/.local/bin/env

uv run python src/scripts/create_weekly_report.py >> "$LOG_FILE" 2>&1
EXIT_CODE=$?

if [ $EXIT_CODE -eq 0 ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Weekly report completed successfully" >> "$LOG_FILE"
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Weekly report failed with exit code $EXIT_CODE" >> "$LOG_FILE"
fi

echo "========================================" >> "$LOG_FILE"
echo "" >> "$LOG_FILE"

exit $EXIT_CODE
