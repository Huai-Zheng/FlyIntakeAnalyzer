# FlyIntakeAnalyzer
High-throughput Drosophila feeding behavior analysis system
# FlyIntakeAnalyzer - 果蝇摄食量数据分析工具

![GUI界面截图](screenshot.png) 

## 简介

本工具是一个基于PyQt5开发的图形化应用程序，专门用于处理Flyplate-BCA果蝇摄食实验中的吸光度数据。通过自动化计算标准曲线、样品浓度和摄食量，帮助研究人员快速生成标准化分析报告。

## 主要功能

- 🖼️ 友好的GUI界面操作
- 📊 自动计算标准曲线和R²值
- ⚡ 一键式批量处理样品数据
- 📈 自动生成带误差线的柱状图
- 📉 标准曲线图表与趋势线绘制
- 📁 结果输出为格式规范的Excel文件
- ✅ 自动调整列宽和样式美化

## 快速开始

### 环境要求
- Python 3.7+
- 依赖库：PyQt5, pandas, numpy, scipy, openpyxl

### 安装步骤
```bash
# 克隆仓库
git clone https://github.com/Huai-Zheng/FlyIntakeAnalyzer.git

# 安装依赖
pip install -r requirements.txt

# Windows用户可直接运行打包版本
