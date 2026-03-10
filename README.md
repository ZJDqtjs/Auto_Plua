# Auto-Plua (MVP)

Auto-Plua（Power & Launch Unified Automation）是一个面向 Windows 的自动化控制工具（MVP）。

项目名称含义：
- `P`：Power，电源系统管控（开机、休眠、关机、唤醒等）
- `L`：Launch，程序启动与任务调度
- `U`：Unified，统一管理入口
- `A`：Automation，自动化执行核心

当前版本包含：
- 程序启动/停止/重启与状态检测
- 定时任务调度（基于 APScheduler）
- 电源控制（关机/重启/睡眠/锁屏）
- 本地配置持久化与日志
- PySide6 图形界面
- OpenCV 识图流程编辑与执行（含等待模块、默认等待与超时退出）
- OpenCV 执行模式切换：前台模拟输入 / 后台窗口消息输入（不占用鼠标键盘）

## 1. 环境准备

- Python 3.10+
- Windows 10/11

## 2. 安装依赖

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 3. 运行

```powershell
python src/autoplua/main.py
```

也可以直接用模块方式启动：

```powershell
python -m src.autoplua.main
```

## 4. 项目结构

```text
Auto_plua/
  requirements.txt
  src/
    autoplua/
      main.py
      config.py
      logger.py
      models.py
      services/
        process_service.py
        scheduler_service.py
        power_service.py
      ui/
        main_window.py
```

## 5. 说明

- MVP 版本优先保证“可运行 + 可扩展”。
- 复杂自动化（如 GUI 图像识别）可后续在 `services` 中扩展。
- 用户配置默认保存在项目根目录 `autoplua.user.json`，可直接共享给其他人使用。
- 可通过环境变量 `AUTOPLUA_CONFIG_PATH` 指定配置文件路径。

## 6. 睡眠唤醒 + 息屏运行 + 不占用键鼠方案

### 6.1 当前已实现能力

- 睡眠唤醒：通过 `CreateWaitableTimerW + SetWaitableTimer(fResume=True)` 安排下一次唤醒。
- 自动登录配置：支持写入 Winlogon 注册表（需管理员权限）。
- 不占用键鼠输入：在程序配置中将“输入模式”设为“后台窗口消息（不抢鼠标键盘）”，可通过窗口消息发送点击、滚轮、文本和回车。
- 息屏保护日志：当截图源不可用（常见于息屏/锁屏）时，流程会直接返回明确错误，不再只显示模糊超时。
- 虚拟显示器驱动接入：可在电源页选择 `.inf` 驱动并一键安装启用（调用 `pnputil` + `DisplaySwitch /extend`）。

### 6.2 关键限制（Windows 机制）

- 真实息屏/锁屏后，桌面渲染通常会暂停，OpenCV 截图可能黑屏或空帧。
- 后台窗口消息输入只能解决“键鼠占用”，不能单独解决“息屏渲染中断”。
- 要在息屏下继续识图，建议使用虚拟显示器驱动保持桌面持续渲染。

### 6.3 推荐落地组合

1. 电源页开启“电源自动化”，配置开机时间与睡眠/关机动作。
2. 在电源页配置虚拟显示器：
  - 勾选“每次执行 OpenCV 前自动准备虚拟显示器”
  - 需要自动安装时再勾选“若未检测到虚拟显示器则自动安装驱动（需管理员）”
  - 选择虚拟显示器驱动 `INF` 路径并点击“安装并启用”测试
3. 将程序配置为“后台窗口消息”并填写准确窗口标题。
4. 以管理员身份运行 AutoPlua；长期运行建议封装为 Windows 服务或开机自启常驻。

### 6.5 重要说明

- AutoPlua 不能在本机“从零生成一个新的内核显示驱动”，但可以自动安装你提供的现成虚拟显示驱动包（INF）。
- 若你还没有驱动包，可先下载并解压一个可用的 IDD/Virtual Display Driver，再在电源页指定其 INF 文件。

### 6.4 常见故障排查

- 日志 `screen-capture-unavailable-possibly-screen-off-or-locked`：说明当前无有效截图源，优先检查虚拟显示器与锁屏状态。
- 日志 `target-window-not-found`：窗口标题不匹配，需使用系统实际窗口标题。
- 日志 `step-timeout-20s-click-target-not-found`：模板识别失败，检查模板清晰度、分辨率一致性、窗口尺寸和 DPI。
