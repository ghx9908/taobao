#!/bin/bash

# 激活虚拟环境
source venv/bin/activate

# 检查Chrome浏览器是否可用
echo "检查Chrome浏览器配置..."
python3 -c "from selenium import webdriver; from selenium.webdriver.chrome.service import Service; from webdriver_manager.chrome import ChromeDriverManager; from selenium.webdriver.chrome.options import Options; print('Chrome驱动配置正常！')"

# 运行Python程序
echo "启动爬虫程序..."
python3 华盛顿.py

# 保持虚拟环境激活状态，方便调试
echo "程序运行完成。虚拟环境仍然激活中。"
echo "要退出虚拟环境，请运行: deactivate"
