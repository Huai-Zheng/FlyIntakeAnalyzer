@echo off
setlocal

:: 创建干净的虚拟环境
python -m venv venv
call venv\Scripts\activate.bat

:: 安装基础依赖
python -m pip install --upgrade pip setuptools wheel

:: 精准安装项目依赖（跳过冲突包）
pip install -r requirements.txt --no-deps

:: 单独安装必需子依赖
pip install "FuzzyTM>=0.4.0" "python-pptx==0.6.19"

:: 打包操作
pyinstaller build.spec --noconfirm

:: 复制运行时文件
xcopy "venv\Lib\site-packages\scipy\.libs\*" "dist\BCAAnalyzer\scipy\.libs\" /s /y

echo 打包成功！
pause
