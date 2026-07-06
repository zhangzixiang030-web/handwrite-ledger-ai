# Handwrite Ledger AI

Handwrite Ledger AI 是一个本地桌面端/浏览器端工具，用于把拍照得到的手写台账、固定格式表格、业务登记表等图片，转换为结构化 Excel 结果。

提供一套通用的“图片表格 -> OCR -> 字段拆解 -> 视觉复核 -> 人工审核 -> 模板导出”流程。适合处理姓名、证件号、地址、电话、金额等字段需要准确落表的手写材料。

## 功能

- 支持一次上传多张图片，按任务批量处理。
- 每张图片 OCR 识别 3 次，并按一致性和置信度选择结果。
- 支持姓名、身份证号、地址、电话、金额等字段的行级抽取。
- 支持可选的 LLM 和 LLM_VL 后处理，提高复杂手写场景下的结构化准确率。
- 支持逐行审核模式：左侧选择行，右侧展示对应行图和字段结果。
- 审核时可直接修改字段，保存后重新导出。
- 支持按 `.xls` 模板导出，尽量保留原模板版式和单元格样式。
- 输出文件、审核明细、任务历史均保存在本地。

## 三种识别策略

本项目支持按能力逐步增强的三层处理方式：

1. **规则拆解 OCR 结果**

   先对 OCR 文本和坐标进行行聚类，再用规则拆解字段。优点是快、稳定、成本低；缺点是遇到手写粘连、错列、姓名误识别、边栏标记混入时，纠错能力有限。

2. **LLM 拆解 OCR 结果**

   将 OCR 文本、行号、坐标和初步字段结果交给 LLM 进行结构化纠错。相比纯规则方式，它能更好处理字段粘连、缺字、噪声和格式异常，但仍主要依赖 OCR 文本本身。

3. **LLM+VL 拆解 OCR 结果**

   在 LLM 的基础上，把对应的原图行裁剪图一并传入，让模型同时参考“图片字形”和“OCR 文本坐标”。这能明显改善手写姓名、行对齐、相邻行串列、金额列误判等问题，是当前推荐的高准确率模式。

## 实测效果

传统 OCR 对手写姓名、混合数字字段、列边界判断比较敏感。LLM_VL 层会同时查看：

- 原始行裁剪图片
- OCR 文本和坐标
- 程序初步拆解出的字段

在一组手写台账图片测试中，三种处理方式的核心字段效果如下：

| 指标 | 规则拆解 OCR | LLM 拆解 OCR | LLM+VL 拆解 OCR |
|---|---:|---:|---:|
| 匹配行数 | 42 | 94 | 96 |
| 姓名正确率 | 59.52% | 61.7% | 75.0% |
| 身份证正确率 | 90.48% | 89.36% | 91.7% |
| 金额正确率 | 97.62% | 91.49% | 92.0% |

可以看到，LLM+VL 对姓名字段的提升最明显。金额字段本身更依赖列定位和数字边界判断，规则拆解在能成功匹配的行里表现较好，但整体匹配行数明显偏少；加入 LLM 和视觉复核后，覆盖行数和姓名识别能力更稳定。

## 处理流程

```text
图片
  -> 三次 OCR 识别与比对
  -> 行定位和字段初拆
  -> 可选 LLM 文本纠错
  -> 可选 LLM_VL 行图复核
  -> 逐行人工审核
  -> 按模板导出 Excel
```

## 配置

仓库中不保存任何真实 API Key。可以复制 `.env.example` 为 `.env`，或直接设置环境变量。

必填 OCR Token：

```powershell
$env:PADDLEOCR_TOKEN="your-ocr-token"
```

可选 LLM_VL 配置：

```powershell
$env:VISION_LLM_API_URL="https://your-openai-compatible-endpoint/v1/chat/completions"
$env:VISION_LLM_API_KEY="your-vision-llm-key"
$env:VISION_LLM_MODEL="your-vision-model"
```

可选默认模板路径：

```powershell
$env:LEDGER_TEMPLATE_PATH="C:\path\to\template.xls"
```

也可以在页面中为每次任务上传 `.xls` 模板。

## 浏览器模式运行

```powershell
python -m pip install -r requirements.txt
python app.py
```

打开：

```text
http://127.0.0.1:8765
```

## 桌面端运行

```powershell
.\run_desktop.ps1
```

## 打包桌面端

```powershell
.\build_desktop.ps1
```

构建结果位于：

```text
dist\LedgerOCRDesktop\LedgerOCRDesktop.exe
```

分发时需要移动整个 `dist\LedgerOCRDesktop` 文件夹。

## 注意

- 输出文件和审核数据保存在本地 `runs/` 目录。
- `runs/`、`dist/`、`build/`、日志和生成的表格文件不会提交到 Git。
- 公开部署或分享构建产物前，请确认 `.env`、日志、历史任务目录中没有敏感信息。
