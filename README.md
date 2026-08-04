# REPORT_GENERATOR

基于 `template/总量通用.docx` 的 OOXML 包级 Word 排版流水线。生成过程不新建空白文档；模板的节、页眉、页脚、主题、样式和法律声明页会被保留。

## 一键运行

```bash
/opt/miniconda3/bin/python scripts/run_pipeline.py "input/原始报告.docx"
```

默认输出：

```text
output/【广发策略】原始报告.docx
```

## 流水线

1. `parse_report.py`：按正文顺序解析标题、段落、图、图表和表格。
2. `asset_manager.py` v3：为原生 chart/image/table 建立稳定定位，并记录递归依赖。
3. `clean_structure.py`：合并题注并绑定所有 `asset_id`。
4. `generate_report.py`：替换模板封面、目录和正文示例区；保留后置模板区域。
5. `asset_copy.py`：复制图表、图片、表格及关系依赖，解决包内文件名冲突。

模板槽位配置位于 `config/template_layout.json`。生成器会在写出前检查 OOXML、内部关系目标和示例占位符残留；任一原生对象未匹配时直接失败，不会退化成截图或图片。

## 输入文档约定

- 标题层级按可见编号识别，不要求预先设置模板字体或 Word 标题样式：`一、` 为一级，`（一）`/`(一)` 为二级，`1.`/`（1）`/`(1)` 为三级。
- 编号可以是普通文本，也可以使用 Word 自动编号；正文的字体、字号和颜色不参与标题判定。
- 标题编号应位于段落开头，标题与正文分别成段；不要只靠加粗、字号或颜色表示层级。
- 图表题注应是紧邻图表的独立段落，并以 `图` 或 `表` 开头；没有题注时生成器不会从模板或其他报告猜测标题。
- 图片、Word 原生图表和 Word 表格均可直接放入正文；生成器按文档中的先后顺序复制。

分析师、研究团队和法律声明页只取自 `template/总量通用.docx`，运行时不依赖 `examples/`。
