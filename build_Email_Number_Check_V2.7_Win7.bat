@echo off
setlocal
chcp 65001 >nul

echo =================================================
echo Email附件检测分析工具 V2.7 - Win7/新电脑优化打包
echo =================================================

where py >nul 2>&1
if errorlevel 1 (
  echo [错误] 未找到 Python 启动器 py。
  echo 建议：Win7 SP1 使用 Python 3.8.x。
  pause
  exit /b 1
)

py -3.8 -c "import sys; print(sys.version)" >nul 2>&1
if errorlevel 1 (
  echo [错误] 未安装 Python 3.8。
  echo 请安装 Python 3.8.x 后再运行本脚本。
  pause
  exit /b 1
)

if not exist .venv_v27 (
  echo [1/7] 创建干净虚拟环境...
  py -3.8 -m venv .venv_v27
  if errorlevel 1 goto :fail
)

call .venv_v27\Scripts\activate.bat
if errorlevel 1 goto :fail

echo [2/7] 升级基础打包工具...
python -m pip install --upgrade pip setuptools wheel
if errorlevel 1 goto :fail

echo [3/7] 安装固定依赖...
python -m pip install "PyQt5==5.15.10" "openpyxl==3.1.5" "pyinstaller==5.13.2"
if errorlevel 1 goto :fail

echo [4/7] 检查源码...
python -m py_compile Email_Number_Check_V2.7.py
if errorlevel 1 goto :fail

if not exist "Email附件数据分析v1.0使用说明.docx" (
  echo [错误] 缺少使用手册：Email附件数据分析v1.0使用说明.docx
  echo 请把 DOCX 与本 BAT 和 PY 文件放在同一目录。
  goto :fail
)

echo [5/7] 清理旧构建目录...
if exist build\Email附件分析V2.7 rmdir /s /q build\Email附件分析V2.7
if exist dist\Email附件分析V2.7 rmdir /s /q dist\Email附件分析V2.7

echo [6/7] 开始构建 onedir 版本...
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name "Email附件分析V2.7" ^
  --exclude-module matplotlib ^
  --exclude-module IPython ^
  --exclude-module tkinter ^
  --exclude-module PySide ^
  --exclude-module PySide2 ^
  Email_Number_Check_V2.7.py
if errorlevel 1 goto :fail

echo [7/7] 复制使用手册到程序目录...
copy /Y "Email附件数据分析v1.0使用说明.docx" "dist\Email附件分析V2.7\Email附件数据分析v1.0使用说明.docx" >nul
if errorlevel 1 goto :fail

echo.
echo =================================================
echo 打包完成：
echo dist\Email附件分析V2.7\
echo.
echo 帮助菜单：
echo   使用手册 / 软件版本 / 系统信息
echo.
echo 默认邮件目录：
echo   程序目录的父级目录
echo.
echo Win7 老电脑建议直接运行整个 onedir 文件夹。
echo =================================================
pause
exit /b 0

:fail
echo.
echo [失败] 打包过程发生错误，请查看上方日志。
pause
exit /b 1
