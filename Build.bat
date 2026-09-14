@echo off
setlocal
chcp 65001 >nul

echo =================================================
echo Email附件检测分析工具 V2.7 - Anaconda 环境打包
echo =================================================

REM 检查 conda 是否可用
where conda >nul 2>&1
if errorlevel 1 (
  echo [错误] 未找到 conda，请确保 Anaconda 已安装并配置了环境变量。
  pause
  exit /b 1
)

REM 获取当前 conda 环境名
for /f "delims=" %%i in ('conda info --envs ^| find "*"') do set CURRENT_ENV=%%i
echo [1/6] 当前 conda 环境：%CURRENT_ENV%

REM 检查 Python 版本（至少 3.7）
python -c "import sys; assert sys.version_info >= (3,7,0)" >nul 2>&1
if errorlevel 1 (
  echo [错误] Python 版本过低，请使用 Python 3.7 或更高版本。
  pause
  exit /b 1
)

echo [2/6] 检查并安装打包依赖...
python -m pip install --upgrade pip setuptools wheel >nul 2>&1
python -m pip install "PyQt5" "openpyxl" "pyinstaller" >nul 2>&1
echo 依赖已安装（或已存在）。

echo [3/6] 检查源码...
python -m py_compile Email_Number_Check_V2.7.py
if errorlevel 1 (
  echo [错误] 源码编译失败，请检查 Email_Number_Check_V2.7.py 是否有语法错误。
  goto :fail
)

if not exist "Email附件数据分析v1.0使用说明.docx" (
  echo [警告] 缺少使用手册：Email附件数据分析v1.0使用说明.docx
  echo 继续打包，但手册不会被复制。
)

echo [4/6] 清理旧构建目录...
if exist build\Email附件分析V2.7 rmdir /s /q build\Email附件分析V2.7
if exist dist\Email附件分析V2.7 rmdir /s /q dist\Email附件分析V2.7

echo [5/6] 开始构建 onedir 版本...
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name "Email附件分析V2.7" ^
  --exclude-module gevent ^
  --exclude-module greenlet ^
  --exclude-module matplotlib ^
  --exclude-module IPython ^
  --exclude-module tkinter ^
  --exclude-module PySide ^
  --exclude-module PySide2 ^
  --exclude-module notebook ^
  --exclude-module pytest ^
  --hidden-import PyQt5 ^
  --hidden-import openpyxl ^
  Email_Number_Check_V2.7.py

if errorlevel 1 goto :fail

echo [6/6] 复制使用手册到程序目录...
if exist "Email附件数据分析v1.0使用说明.docx" (
  copy /Y "Email附件数据分析v1.0使用说明.docx" "dist\Email附件分析V2.7\Email附件数据分析v1.0使用说明.docx" >nul
)

echo.
echo =================================================
echo ? 打包完成！
echo ?? 输出目录：dist\Email附件分析V2.7\
echo.
echo ?? 使用方法：
echo   1. 将整个 Email附件分析V2.7 文件夹复制到目标电脑
echo   2. 双击运行 Email附件分析V2.7.exe
echo.
echo ?? 注意：Win7 需要安装 KB2533623 补丁
echo =================================================
pause
exit /b 0

:fail
echo.
echo ? 打包失败，请检查上方错误日志。
pause
exit /b 1