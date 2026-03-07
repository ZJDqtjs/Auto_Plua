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
