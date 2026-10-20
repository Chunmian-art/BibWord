#!/bin/bash
set -e
cd "$(dirname "$0")"
if ! command -v mkcert >/dev/null 2>&1; then
  if ! command -v brew >/dev/null 2>&1; then echo '请先安装 Homebrew: https://brew.sh'; exit 1; fi
  brew install mkcert
fi
mkcert -install
mkcert -cert-file localhost.pem -key-file localhost-key.pem localhost 127.0.0.1
WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
mkdir -p "$WEF"
cp manifest.xml "$WEF/BibWord.xml"
echo '已复制 Word 插件清单。请重启 Word，打开 首页 > 加载项 > BibWord 引用。'
echo '请保持此终端窗口运行。'
python3 server.py
