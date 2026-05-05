#!/bin/bash
# apo→FAS 自動同期の launchd インストール/アンインストールスクリプト
# 使い方:
#   bash install_apo_sync.sh install   # インストール（1時間ごとの自動同期開始）
#   bash install_apo_sync.sh uninstall # アンインストール
#   bash install_apo_sync.sh status    # 現在の状態確認
#   bash install_apo_sync.sh test      # 1回だけ手動同期

set -e
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PLIST_SRC="$SCRIPT_DIR/com.hokuso.fas.apo-sync.plist"
PLIST_DST="$HOME/Library/LaunchAgents/com.hokuso.fas.apo-sync.plist"
LABEL="com.hokuso.fas.apo-sync"

cmd="${1:-status}"

case "$cmd" in
  install)
    echo "→ launchd にインストールします..."
    mkdir -p "$HOME/Library/LaunchAgents"
    cp "$PLIST_SRC" "$PLIST_DST"
    launchctl unload "$PLIST_DST" 2>/dev/null || true
    launchctl load "$PLIST_DST"
    echo "✓ インストール完了。毎時05分に自動同期が走ります"
    echo "  ログ: ~/Library/Logs/apo_sync.log"
    ;;
  uninstall)
    echo "→ アンインストールします..."
    launchctl unload "$PLIST_DST" 2>/dev/null || true
    rm -f "$PLIST_DST"
    echo "✓ 自動同期を停止しました"
    ;;
  status)
    if [ -f "$PLIST_DST" ]; then
      echo "✓ インストール済み: $PLIST_DST"
      launchctl list | grep "$LABEL" || echo "  (未ロード)"
    else
      echo "× 未インストール"
    fi
    ;;
  test)
    echo "→ 1回だけ手動実行..."
    python3 "$SCRIPT_DIR/apo_sync.py"
    ;;
  *)
    echo "Usage: bash install_apo_sync.sh {install|uninstall|status|test}"
    exit 1
    ;;
esac
