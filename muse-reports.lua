-- Muse 选股报告 · Hammerspoon 窗口壳（窗口版）
-- 菜单栏「📈 报告」点击 / Dock 里「选股报告.app」/ 新报告到达
-- → 弹出独立窗口（Webview）显示本地报告服务的标签页外壳 /shell
--   （外壳内是真标签页：列表固定第一个，「新标签」在同一窗口里开新标签）
-- 由 install.command 复制到 ~/.hammerspoon/ 并在 init.lua 里 require。

local M = {}

local SERVICE = "http://127.0.0.1:8765/"
local MANIFEST = os.getenv("HOME") .. "/选股报告/manifest.json"
local REPORTS_DIR = os.getenv("HOME") .. "/选股报告"

local wv = nil
local lastTotal = -1

local function readManifest()
  local f = io.open(MANIFEST, "r")
  if not f then return nil end
  local s = f:read("*a")
  f:close()
  if not s or s == "" then return nil end
  local ok, data = pcall(hs.json.decode, s)
  if ok then return data end
  return nil
end

local function counts()
  local m = readManifest()
  if not m then return 0 end
  return #m
end

local menubar = hs.menubar.new()

local function refreshMenubar()
  if menubar then
    menubar:setTitle("📈 报告")
  end
  return counts()
end

local function ensureWebview()
  if wv and wv:hswindow() then return wv end
  local screen = hs.screen.mainScreen():frame()
  local rect = {
    x = screen.x,
    y = screen.y,
    w = screen.w,
    h = screen.h,
  }
  wv = hs.webview.new(rect)
  -- 允许在网页输入框里打字（Hammerspoon 的 webview 默认不让输入，
  -- 不开这个，搜索框点进去也敲不进字）
  pcall(function() wv:allowTextEntry(true) end)
  pcall(function() wv:level(0) end)  -- 普通窗口层级，不压住其他 app
  local ok = pcall(function()
    wv:windowStyle({ "titled", "closable", "resizable", "miniaturizable" })
  end)
  if not ok then
    pcall(function()
      wv:windowStyle({ "titled", "closable", "resizable" })
    end)
  end
  wv:windowTitle("")
  wv:url(SERVICE .. "shell")
  return wv
end

local function showWindow()
  local v = ensureWebview()
  -- 窗口里是标签页外壳（/shell）：已在外壳就不重载，免得把开着的标签冲掉；
  -- 不在外壳（比如刚启动）才载入外壳
  local cur = nil
  pcall(function() cur = v:url() end)
  if not (cur and cur:find("shell")) then
    v:url(SERVICE .. "shell")
  end
  v:show()
  local win = v:hswindow()
  if win then win:focus() end
end

if menubar then
  menubar:setClickCallback(function()
    showWindow()
    refreshMenubar()
  end)
end

-- Dock 里的「选股报告.app」通过这个链接唤起窗口
hs.urlevent.bind("openreports", function()
  showWindow()
  refreshMenubar()
end)

-- 注：报告页「全屏」走的 hammerspoon://togglefs 绑定在 init.lua 末尾
-- （2026-11-07 挪过去的：模块内绑定在他机器上不触发，原因未明；
-- init.lua 版按窗口枚举找报告窗口，不依赖本模块内部变量。）

-- 盯住 manifest：新报告（总数变多）时通知 + 弹窗；标已读不弹窗
local function onManifestChanged()
  local total = refreshMenubar()
  if lastTotal >= 0 and total > lastTotal then
    hs.notify.new({
      title = "Muse 选股报告",
      informativeText = "有新报告到了，已打开窗口",
    }):send()
    showWindow()
  end
  lastTotal = total
end

local watcher = hs.pathwatcher.new(REPORTS_DIR, function(files)
  for _, p in ipairs(files) do
    if p:find("manifest%.json") then
      hs.timer.doAfter(0.5, onManifestChanged)
      return
    end
  end
end)
watcher:start()

-- 兜底轮询（pathwatcher 偶尔漏事件时靠它刷新）
hs.timer.doEvery(120, function() refreshMenubar() end)

lastTotal = refreshMenubar()
hs.notify.new({
  title = "Muse 选股报告",
  informativeText = "报告窗口已就绪，点菜单栏「📈 报告」打开",
}):send()

return M
