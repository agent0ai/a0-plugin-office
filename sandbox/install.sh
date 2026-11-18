#!/bin/sh
set -eu
export DEBIAN_FRONTEND=noninteractive
python /plugin/sandbox/progress.py installing 'Installing LibreOffice. First use can take several minutes.'
dpkg --configure -a
apt-get update
apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc libreoffice-impress libreoffice-gtk3 python3-uno python3-docx python3-openpyxl dbus-x11 matchbox-window-manager fonts-liberation fonts-dejavu-core
for tool in /opt/tuxee-gui/bin/*; do ln -sfn "$tool" "/usr/local/bin/$(basename "$tool")"; done
if [ ! -e /usr/bin/xkbcomp ]; then ln -s /opt/tuxee-gui/bin/xkbcomp /usr/bin/xkbcomp; fi
/usr/bin/python3 -c 'import uno, docx, openpyxl'
rm -rf /var/lib/apt/lists/*
