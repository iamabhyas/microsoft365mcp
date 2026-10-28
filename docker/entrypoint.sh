#!/bin/sh
set -eu

if [ "$(id -u)" = "0" ]; then
  mkdir -p /data/tokens
  chown -R app:app /data/tokens
  exec setpriv --reuid=app --regid=app --init-groups -- "$@"
fi

exec "$@"
