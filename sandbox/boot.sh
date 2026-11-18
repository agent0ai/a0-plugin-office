#!/bin/sh
set -eu
mkdir -p /var/lib/a0-office /run/a0-office
rm -f /var/lib/a0-office/ready.json /var/lib/a0-office/progress.json
export A0_OFFICE_BOOT="$1"
# Keep fast startup failures in the launcher's log too, before manager cleanup.
exec 3>&1
trap 'code=$?; if [ "$code" -ne 0 ]; then tail -c 6000 /var/lib/a0-office/startup.log >&3; fi' 0
exec > /var/lib/a0-office/startup.log 2>&1
fingerprint=$(cat /plugin/sandbox/install.sh | sha256sum | cut -d ' ' -f 1)
if [ ! -f /var/lib/a0-office/installed ] || [ "$(cat /var/lib/a0-office/installed)" != "$fingerprint" ]; then
    sh /plugin/sandbox/install.sh
    printf '%s\n' "$fingerprint" > /var/lib/a0-office/installed.tmp
    mv /var/lib/a0-office/installed.tmp /var/lib/a0-office/installed
fi
export PYTHONDONTWRITEBYTECODE=1
export PYTHONPATH=/support:/plugin/sandbox
/usr/bin/python3 /plugin/sandbox/service.py "$1"
