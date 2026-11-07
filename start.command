#!/bin/bash
# Muse 选股报告 · 启动服务（退出后想再开，双击我）
# 服务平时是开机自启的；只有点了页面上「退出服务」之后才需要手动启动。
launchctl load "$HOME/Library/LaunchAgents/com.muse.mailreports.plist" 2>/dev/null
sleep 1
open "http://127.0.0.1:8765/"
