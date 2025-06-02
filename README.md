# 智能翻译工具 (Smart Translation Tool)

一个功能强大的多功能翻译工具，支持文本翻译、文档翻译和字幕翻译。

## 功能特点

- 文本翻译：支持实时文本翻译，带有术语管理功能
- 文档翻译：支持Excel文件的批量翻译，可同时翻译多种目标语言
- 字幕翻译：支持SRT和ASS格式字幕文件的翻译
- 术语管理：可以创建和管理翻译术语表，确保专业术语翻译的一致性
- 深色/浅色主题：支持界面主题切换，保护视觉体验
- 翻译历史：记录所有翻译任务的历史记录
- 批量处理：支持多线程翻译，提高处理效率
- 进度显示：实时显示翻译进度和预计剩余时间

## 安装要求

- Python 3.8+
- 依赖包：见 requirements.txt

## 安装步骤

1. 克隆仓库：
```bash
git clone https://github.com/Amadeus9029/smart-translation-tool.git
cd smart-translation-tool
```

2. 安装依赖：
```bash
pip install -r requirements.txt
```

3. 配置API密钥：
   - 在应用设置中配置您的DeepSeek API密钥

## 使用说明

1. 启动应用：
```bash
python translate_app.py
```

2. 文本翻译：
   - 在"文本翻译"页面输入要翻译的文本
   - 选择源语言和目标语言
   - 点击"翻译"按钮开始翻译

3. 文档翻译：
   - 在"文档翻译"页面选择Excel文件
   - 选择源语言和目标语言
   - 设置保存路径
   - 点击"开始翻译"按钮

4. 字幕翻译：
   - 在"字幕翻译"页面选择字幕文件
   - 选择源语言和目标语言
   - 点击"翻译"按钮开始翻译
   - 翻译完成的文件会保存在subtitle_result目录下

## 配置说明

- API设置：配置DeepSeek API密钥
- 翻译参数：可调整并发线程数、批处理大小等参数
- 界面主题：可选择明亮或深色主题
- 存储设置：设置翻译结果的保存位置

## 注意事项

- 请确保有足够的API余额
- 大文件翻译时建议使用合适的并发线程数
- 定期备份重要的翻译术语表

## 更新日志

### v1.0.0 (2025-06-02)
- 初始版本发布
- 支持文本、文档和字幕翻译
- 实现术语管理功能
- 支持主题切换
- 添加翻译历史记录

## 贡献指南

欢迎提交问题和改进建议！请遵循以下步骤：

1. Fork 本仓库
2. 创建您的特性分支 (git checkout -b feature/AmazingFeature)
3. 提交您的更改 (git commit -m 'Add some AmazingFeature')
4. 推送到分支 (git push origin feature/AmazingFeature)
5. 开启一个 Pull Request

## 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件 
