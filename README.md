# AutoPlua (MVP)

AutoPlua 是一个 Windows 自动化控制工具（MVP），包含：
- 程序启动/停止/重启与状态检测
- 定时任务调度（基于 APScheduler）
- 电源控制（关机/重启/睡眠/锁屏）
- 本地配置持久化与日志
- PySide6 图形界面

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
python autoplua/src/autoplua/main.py
```

也可以先进入子目录后使用模块方式启动：

```powershell
cd autoplua
python -m src.autoplua.main
```

## 4. 项目结构

```text
Auto_plua/
  requirements.txt
  autoplua/
    README.md
    src/autoplua/
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
