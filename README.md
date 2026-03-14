# AutoPlua

AutoPlua（Power & Launch Unified Automation）是一个面向 Windows 的自动化控制工具，聚焦于“程序调度 + 电源自动化 + OpenCV 流程执行”。

项目仓库：`https://github.com/ZJDqtjs/Auto_Plua`

## 功能概览

- 程序启动、停止、重启与状态检测
- 定时任务调度（APScheduler）
- 电源控制（关机、重启、睡眠、锁屏）
- 本地配置持久化与日志记录
- OpenCV 流程编辑与执行（等待模块、默认等待、超时退出）
- 输入模式切换：前台模拟输入 / 后台窗口消息输入（不抢键鼠）
- 虚拟显示器接入：支持内置驱动安装与扩展显示检测
- 关于页新增仓库入口与“检查更新”按钮（GitHub Release/Tag）

## 环境要求

- Windows 10/11
- Python 3.10+

## 安装与运行

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python src/autoplua/main.py
```

也可以使用模块方式启动：

```powershell
python -m src.autoplua.main
```

## 界面说明

左侧导航包含：`开始`、`电源`、`日志`、`关于`。

- 开始：管理程序列表、运行状态、调度时间点
- 电源：配置自动开机登录、关机动作、虚拟显示器策略
- 日志：查看应用日志与被启动程序日志
- 关于：查看版本、打开仓库链接、检查更新

### 关于页更新检查机制

- 点击“检查更新”后，会请求 GitHub API：
  - 优先读取 `releases/latest`
  - 若无 Release，则回退读取 `tags`
- 若检测到远端版本高于本地版本，会展示可点击的版本链接
- 若无法联网或无版本数据，会给出明确提示并保留仓库直达入口

## 项目结构

```text
Auto_plua/
  autoplua.user.json
  requirements.txt
  drivers/
  src/
    autoplua/
      __init__.py
      main.py
      config.py
      logger.py
      models.py
      services/
      ui/
        main_window.py
```

## 配置说明

- 用户配置默认写入项目根目录：`autoplua.user.json`
- 可通过环境变量覆盖配置路径：`AUTOPLUA_CONFIG_PATH`
- 电源自动化配置集中在 `power_settings`

## 睡眠唤醒 + 息屏运行 + 不占键鼠

### 已实现能力

- 睡眠唤醒：优先 `pywin32 + Task Scheduler COM`，失败回退 WaitableTimer
- 自动登录：写入 Winlogon 注册表（需要管理员权限）
- 后台输入：窗口消息模式可避免抢占鼠标键盘
- 息屏保护日志：截图源不可用时给出明确错误
- 虚拟显示器：支持一键安装驱动并切换扩展显示

### 关键限制

- 锁屏/息屏后桌面渲染可能暂停，OpenCV 可能拿到黑屏或空帧
- 后台输入只能解决键鼠占用，不能单独解决渲染中断
- 若要在息屏时持续识图，建议接入虚拟显示器

### 推荐配置流程

1. 在电源页开启“电源自动化”，设置开机与关机时间。
2. 开启“每次执行 OpenCV 前自动准备虚拟显示器”。
3. 需要自动部署驱动时，开启“若未检测到虚拟显示器则自动安装驱动（需管理员）”。
4. 程序配置中输入模式选择“后台窗口消息（不抢鼠标键盘）”，并填写准确窗口标题。
5. 建议管理员权限运行 AutoPlua。

## 常见问题

- `screen-capture-unavailable-possibly-screen-off-or-locked`
  - 当前无可用截图源，优先检查虚拟显示器与锁屏状态。
- `target-window-not-found`
  - 目标窗口标题不匹配。
- `step-timeout-20s-click-target-not-found`
  - 模板识别失败，检查模板质量、分辨率、窗口尺寸与 DPI。
