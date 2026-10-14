#!/bin/bash
# Хук старта облачной сессии Claude Code: Python-зависимости и библиотеки для Qt offscreen.
# Локально (Windows и любой не облачный запуск) ничего не делает.
set -euo pipefail

if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

cd "${CLAUDE_PROJECT_DIR:-$(dirname "$0")/../..}"

# Системные библиотеки, которые нужны PySide6 (QtGui/QtWidgets) в режиме offscreen,
# и шрифты с кириллицей для скриншотов.
APT_PKGS="libegl1 libgl1 libopengl0 libglib2.0-0 libfontconfig1 libfreetype6 libdbus-1-3 \
libxkbcommon0 libxkbcommon-x11-0 fonts-dejavu-core"
missing=""
for p in $APT_PKGS; do
  dpkg -s "$p" >/dev/null 2>&1 || missing="$missing $p"
done
if [ -n "$missing" ]; then
  SUDO=""
  [ "$(id -u)" -ne 0 ] && SUDO="sudo"
  export DEBIAN_FRONTEND=noninteractive
  $SUDO apt-get update -qq || true
  # shellcheck disable=SC2086
  $SUDO apt-get install -y -qq --no-install-recommends $missing
fi

python3 -m pip install --quiet --disable-pip-version-check -r requirements.txt

# Окна Qt в облаке рисуем без дисплея.
if [ -n "${CLAUDE_ENV_FILE:-}" ]; then
  grep -qs 'QT_QPA_PLATFORM' "$CLAUDE_ENV_FILE" || echo 'export QT_QPA_PLATFORM=offscreen' >> "$CLAUDE_ENV_FILE"
fi
