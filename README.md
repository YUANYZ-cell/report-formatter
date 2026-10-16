# Report Formatter · 报告格式化器

Rule-based DOCX formatting and PDF report analysis.

用可解释的规则检查并整理实验报告中的图表、公式编号与引用。

**技术栈 / Stack:** Python · Streamlit · python-docx · PyMuPDF

## 能做什么

- 支持 IEEE、模板匹配与自定义格式，提供中英文界面。
- 分析图题、表题、公式和文内交叉引用，保留报告正文。
- Word 可分析或格式化，PDF 提供分析报告；真实页面预览按本机转换器可用性启用。

## 安装与运行

Python 3.11+。

```powershell
python -m venv .venv
.venv/Scripts/python -m pip install -r requirements.txt
.venv/Scripts/python scripts/make_samples.py
.venv/Scripts/python -m streamlit run app.py
```

高保真 Word 预览需要 Microsoft Word 或 LibreOffice，PDF 页图渲染使用 Poppler。缺少转换器时使用应用提供的降级预览。PDF 是分析模式；扫描 PDF 需先做 OCR。公开仓库不包含个人实验报告；`make_samples.py` 生成合成样例。

## 验证

```powershell
python scripts/make_samples.py
python -m unittest discover -s tests
```

依赖私人文档的集成用例会明确跳过，其余合成文档测试可独立执行。

命令结果与验收范围见 [VALIDATION.md](VALIDATION.md)。这些检查不能替代目标机器上的原生桌面验收。

## 发布范围

[公开源码说明](PUBLICATION.md) · [全部作品](https://github.com/YUANYZ-cell)
