#!/bin/bash
# Mac Mail Reports · 一键安装（双击运行）
# 把本地服务装到 ~/Library/Application Support/MailReports，
# 由 launchd 托管（开机自启、崩溃自动拉起），只监听 127.0.0.1。
cd "$(dirname "$0")"
APP_DIR="$HOME/Library/Application Support/MailReports"
PLIST="$HOME/Library/LaunchAgents/com.muse.mailreports.plist"

echo "=== Mac Mail Reports · 安装 ==="
echo

if ! /usr/bin/python3 -c "pass" >/dev/null 2>&1; then
  echo "这台 Mac 还没有 python3（系统命令行工具）。"
  echo "请先在终端运行：xcode-select --install ，装完后再双击本安装文件。"
  read -p "按回车退出…" _
  exit 1
fi

read -p "接收报告的邮箱地址（如 you@163.com）：" EMAIL
if [ -z "$EMAIL" ]; then echo "邮箱为空，安装取消。"; read -p "按回车退出…" _; exit 1; fi
read -p "报告发件人邮箱（只有这个地址发来的邮件会被当作报告）：" FROM_ADDR
if [ -z "$FROM_ADDR" ]; then echo "发件人为空，安装取消。"; read -p "按回车退出…" _; exit 1; fi
echo "请输入邮箱的 IMAP 客户端授权码（不是邮箱密码；"
echo "在邮箱网页版 设置 → POP3/SMTP/IMAP 里开启 IMAP 后生成）。"
echo "输入时屏幕不显示，是正常的："
read -s -p "授权码：" AUTHCODE
echo
if [ -z "$AUTHCODE" ]; then echo "授权码为空，安装取消。"; read -p "按回车退出…" _; exit 1; fi

mkdir -p "$APP_DIR" "$HOME/Library/LaunchAgents" "$HOME/选股报告"
cp app.py "$APP_DIR/app.py"
[ -f nav_wallpaper.jpg ] && cp nav_wallpaper.jpg "$APP_DIR/nav_wallpaper.jpg"

EMAIL="$EMAIL" FROM_ADDR="$FROM_ADDR" AUTHCODE="$AUTHCODE" APP_DIR="$APP_DIR" /usr/bin/python3 - <<'PYEOF'
import json, os
cfg = {
    "email": os.environ["EMAIL"],
    "auth_code": os.environ["AUTHCODE"],
    "from_addr": os.environ["FROM_ADDR"],
    "reports_dir": "~/选股报告",
    "imap_host": "imap.163.com",
    "imap_port": 993,
    "interval_sec": 300,
    "lookback_days": 4,
}
path = os.path.join(os.environ["APP_DIR"], "config.json")
with open(path, "w", encoding="utf-8") as f:
    json.dump(cfg, f, ensure_ascii=False, indent=2)
os.chmod(path, 0o600)
print("配置已写入（权限 600）：", path)
PYEOF

sed -e "s|__APP_DIR__|$APP_DIR|g" -e "s|__HOME__|$HOME|g" \
  com.muse.mailreports.plist > "$PLIST"
launchctl unload "$PLIST" 2>/dev/null
launchctl load "$PLIST"

# 可选：Hammerspoon 窗口壳（菜单栏「📈 报告」+ 标签页窗口）
if [ -d "$HOME/.hammerspoon" ]; then
  cp muse-reports.lua "$HOME/.hammerspoon/muse-reports.lua"
  echo
  echo "已把窗口壳 muse-reports.lua 放进 ~/.hammerspoon/。"
  echo "请在 ~/.hammerspoon/init.lua 末尾加一行："
  echo '  pcall(require, "muse-reports")'
  echo "然后在 Hammerspoon 菜单里点 Reload Config 生效。"
fi

sleep 2
open "http://127.0.0.1:8765/home"
echo
echo "安装完成：服务运行在 http://127.0.0.1:8765 （仅本机可访问）"
read -p "按回车关闭…" _
