#!/bin/bash
# Run all weekly Santo Taco scripts in sequence.
# Retries each script up to 3 times before giving up.
# Logs to weekly_run.log in the same directory.

DIR="$(cd "$(dirname "$0")" && pwd)"
LOG="$DIR/weekly_run.log"
VENV="$DIR/venv/bin/activate"
MAX_ATTEMPTS=5
RETRY_WAIT=60
FAILED_SCRIPTS=()

echo "=======================================" | tee -a "$LOG"
echo "Weekly run started: $(date)" | tee -a "$LOG"
echo "=======================================" | tee -a "$LOG"

TARGET_SUNDAY=$(python3 -c "from datetime import date, timedelta; d=date.today(); print(d - timedelta(days=(d.weekday()+1)%7))")
echo "Target week ending: $TARGET_SUNDAY" | tee -a "$LOG"

source "$VENV"

for SCRIPT in pl_report.py santo_taco_labor_check.py menu_engineering.py cogs_deep_dive.py; do
    echo "" | tee -a "$LOG"
    echo "--- $SCRIPT $(date) ---" | tee -a "$LOG"

    SUCCESS=0
    for ATTEMPT in $(seq 1 $MAX_ATTEMPTS); do
        if [ $ATTEMPT -gt 1 ]; then
            echo "[retry] Attempt $ATTEMPT of $MAX_ATTEMPTS — waiting ${RETRY_WAIT}s..." | tee -a "$LOG"
            sleep $RETRY_WAIT
        fi

        python "$DIR/$SCRIPT" --week-ending "$TARGET_SUNDAY" 2>&1 | tee -a "$LOG"
        STATUS=${PIPESTATUS[0]}

        if [ $STATUS -eq 0 ]; then
            SUCCESS=1
            break
        else
            echo "[ERROR] $SCRIPT attempt $ATTEMPT failed with status $STATUS" | tee -a "$LOG"
        fi
    done

    if [ $SUCCESS -eq 0 ]; then
        echo "[FAILED] $SCRIPT failed after $MAX_ATTEMPTS attempts" | tee -a "$LOG"
        FAILED_SCRIPTS+=("$SCRIPT")
    fi
done

echo "" | tee -a "$LOG"
echo "=======================================" | tee -a "$LOG"
echo "Weekly run finished: $(date)" | tee -a "$LOG"
echo "=======================================" | tee -a "$LOG"

if [ ${#FAILED_SCRIPTS[@]} -gt 0 ]; then
    FAILED_ARG=$(IFS=','; echo "${FAILED_SCRIPTS[*]}")
    python "$DIR/notify.py" --log "$LOG" --failed "$FAILED_ARG" 2>&1 | tee -a "$LOG"
else
    python "$DIR/notify.py" --log "$LOG" 2>&1 | tee -a "$LOG"
fi
