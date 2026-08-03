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

