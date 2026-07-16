# Metadata Split Tool

一个用于 Excel 指标名称分列处理的 Web 工具。

## 项目背景

在数据指标整理过程中，经常会遇到类似：
江西:固定资产投资:累计同比

这样的层级指标名称。

传统方式需要人工拆分字段，效率较低。

因此开发了 Metadata Split Tool，通过浏览器完成指标名称自动拆分，提高数据整理效率。


## 功能特点

### Excel 文件处理
- 支持 Excel / CSV 文件上传
- 支持拖拽上传

### 智能分列
- 自动识别常见分隔符：
  - :
  - -
  - /
  - _

- 支持手动选择分隔符

### 数据处理
- 支持两种拆分模式：
  - 从左往右填充
  - 首尾对齐模式

### 结果导出
- 支持预览拆分结果
- 支持导出 Excel 文件


## 技术栈

- HTML5
- CSS3
- JavaScript
- SheetJS


## 项目结构
metadata-split-tool
│
├── index.html
├── css
│ └── style.css
├── js
│ └── app.js
└── README.md


## 本地运行

1. 下载项目

2. 使用 VS Code 打开

3. 使用 Live Server 运行：


## 后续计划

- 支持多个 Excel Sheet
- 保存用户配置
- 优化指标名称智能解析
- 增加更多数据清洗能力

## 项目预览

![项目预览](assets/preview.png)