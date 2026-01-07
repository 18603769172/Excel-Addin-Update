@echo off
chcp 65001 >nul 2>&1  :: 设置编码为UTF-8，避免中文乱码

:: 定义要复制的文件和目标路径
set "XLAM_FILE=LROOT工具箱.xlam"
set "CHM_FILE=LROOT工具箱使用说明.CHM"
set "TARGET_DIR=%APPDATA%\Microsoft\AddIns\"

echo 正在检查安装文件...
:: 检查XLAM文件是否存在
if not exist "%XLAM_FILE%" (
    echo 错误：未找到文件 "%XLAM_FILE%"，请确认文件与脚本在同一目录！
    pause
    exit /b 1
)
:: 检查CHM帮助文件是否存在
if not exist "%CHM_FILE%" (
    echo 错误：未找到帮助文件 "%CHM_FILE%"，请确认文件与脚本在同一目录！
    pause
    exit /b 1
)

echo 正在安装Excel工具箱（包含帮助文件）...
:: 执行XLAM文件复制，并检查结果
copy /Y "%XLAM_FILE%" "%TARGET_DIR%" >nul 2>&1
if %errorlevel% neq 0 (
    echo 错误：工具箱文件复制失败，请检查权限或目标目录是否存在！
    pause
    exit /b 1
)

:: 执行CHM帮助文件复制，并检查结果
copy /Y "%CHM_FILE%" "%TARGET_DIR%" >nul 2>&1
if %errorlevel% neq 0 (
    echo 警告：工具箱文件已安装，但帮助文件复制失败！
    echo 你可手动将"%CHM_FILE%"复制到"%TARGET_DIR%"目录。
)

echo 安装完成！请重启Excel以启用工具箱，帮助文件已同步至加载项目录。
pause