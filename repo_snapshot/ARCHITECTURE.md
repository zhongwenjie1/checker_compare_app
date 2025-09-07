# 项目架构快照（自动生成）

- 生成时间：本地运行
- 根目录：`/Users/zhongwenjie/程序`
- 忽略目录：.git, .idea, .mypy_cache, .venv, .vscode, __pycache__, build, dist, venv
- 关键 Python 文件数量：4204

## 目录树
见 `TREE.txt`

## 模块概览（按路径粗分）
- **checker_ui/core/**：排程、导出逻辑（如 `tickets.py`）
- **checker_ui/ui/**：界面层（窗口、对话框、流程图视图等）
- **checker_ui/models/**：状态/数据模型（如 `state.py`）
- **infra/**：线程、IO、工具
- **.github/workflows/**：CI/构建脚本
- ***.spec**：PyInstaller 打包规格

## 入口建议
- `checker_ui/main.py`（若存在）作为图形入口
- CI 入口：`.github/workflows/*.yml`

## Python 文件摘要样例
详见 `SNIPPETS/`，每个文件包含：
- 类/函数签名（静态解析）
- 头部片段（便于快速定位风格与依赖）