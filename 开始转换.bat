@echo off
chcp 65001 >nul
rem Windows：双击运行。第一次会自动安装依赖（约 1～3 分钟），之后几秒就能打开。
cd /d "%~dp0"
set MIRROR=https://pypi.tuna.tsinghua.edu.cn/simple

set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY (where python >nul 2>nul && set PY=python)
if not defined PY (
  echo 没有找到 Python 3。请先到 https://www.python.org/downloads/ 安装 3.10 或更新版本，安装时勾选 "Add python.exe to PATH"。
  pause & exit /b 1
)

if not exist .venv\Scripts\python.exe (
  echo 第一次运行：正在准备运行环境……
  %PY% -m venv .venv || (echo 创建运行环境失败 & pause & exit /b 1)
)
.venv\Scripts\python.exe -c "import pdfplumber, pypdfium2, docx" 2>nul
if errorlevel 1 (
  echo 正在安装依赖（使用清华镜像）……
  .venv\Scripts\python.exe -m pip install --disable-pip-version-check -q -i %MIRROR% -r requirements.txt || (echo 依赖安装失败，请检查网络后重试。 & pause & exit /b 1)
)

.venv\Scripts\python.exe app.py
pause
