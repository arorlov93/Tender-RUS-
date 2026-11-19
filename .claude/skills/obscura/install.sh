#!/usr/bin/env bash
# Установка бинарника Obscura (github.com/h4ckf0r0day/obscura) в ~/.local/bin — контейнер одноразовый, запускать в новой сессии.
set -e
V=${1:-v0.2.7}; D=$(mktemp -d)
curl -sSL -o "$D/o.tgz" "https://github.com/h4ckf0r0day/obscura/releases/download/$V/obscura-x86_64-linux.tar.gz"
tar xzf "$D/o.tgz" -C "$D" && mkdir -p ~/.local/bin && install -m755 "$D/obscura" "$D/obscura-worker" ~/.local/bin/
~/.local/bin/obscura --version
