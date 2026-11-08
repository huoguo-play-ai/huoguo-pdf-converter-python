#!/bin/bash
# Mac：双击运行。第一次会自动安装依赖（约 1～3 分钟），之后几秒就能打开。
cd "$(dirname "$0")" || exit 1
MIRROR="https://pypi.tuna.tsinghua.edu.cn/simple"

if ! command -v python3 >/dev/null 2>&1; then
  echo "没有找到 Python 3。请先到 https://www.python.org/downloads/ 安装（3.10 或更新版本），再双击本文件。"
  read -r -p "按回车退出"; exit 1
fi
if ! python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)'; then
  echo "当前 Python 版本太旧：$(python3 --version)。请安装 3.10 或更新版本。"
  read -r -p "按回车退出"; exit 1
fi

if [ ! -x .venv/bin/python ]; then
  echo "第一次运行：正在准备运行环境……"
  python3 -m venv .venv || { read -r -p "创建运行环境失败，按回车退出"; exit 1; }
fi
if ! .venv/bin/python -c 'import pdfplumber, pypdfium2, docx' 2>/dev/null; then
  echo "正在安装依赖（使用清华镜像）……"
  if ! .venv/bin/python -m pip install --disable-pip-version-check -q -i "$MIRROR" -r requirements.txt; then
    echo "依赖安装失败，请检查网络后重试。"; read -r -p "按回车退出"; exit 1
  fi
fi

.venv/bin/python app.py
