# python-word-
对于公司平台自动生成的大量漏洞报告，由于一个一个查看有点太浪费时间，于是使用自动化python脚本进行自动提取正则匹配到的漏洞情况，第一个脚本是提取具体漏洞名称，第二个脚本是提取具体的高低中危的漏洞数量。


#redate.py
# 批量 Word(.docx) 日期替换与统计脚本使用说明

脚本名称（建议）：`bulk_replace_phase1_multi.py`  
适用：批量扫描并替换多个固定日期字符串，同时按“期望出现次数”标记异常，生成完整统计报表。  
当前版本特性针对：  
- 将 `2025-08-19` 与 `2025-08-20`（各期望 4 次）替换为统一的新日期 `2025-07-31`  
- 出现次数 ≠ 4 仍然替换，但标记为异常（方便人工复核）  
- 支持正文、表格、页眉、页脚（不含文本框 / 形状 / 批注 / 脚注）

---

## 1. 适用场景
- 多个 Word 文档中重复出现固定日期，需要统一替换，同时希望保留统计证据。
- 有“期望出现次数”业务规则（如某字段应出现 4 次），需要标记偏差。
- 想要安全回滚：自动生成 `.bak` 备份文件。
- 需要按文件输出明细（CSV）和异常清单。

---

## 2. 功能清单
| 功能 | 说明 |
|------|------|
| 多规则日期替换 | 配置多个 target → replacement |
| 期望次数校验 | 每个规则可设 expected，不符标记异常 |
| DRY-RUN 预演 | 不改文件，仅输出统计 |
| 备份机制 | 修改前生成 `.docx.bak` |
| 递归遍历 | 支持子目录 |
| 结构遍历 | 正文 / 表格 / 页眉 / 页脚 |
| 报表输出 | `summary.txt` / `detail.csv` / `abnormal_list.txt` / `error_list.txt` |
| 异常判定 | 任一规则 mismatch → 文件状态 `abnormal` |
| 安全可扩展 | 可追加新规则、正则扫描、JSON 导出等（见扩展章节） |

---

## 3. 目录结构建议
```
your_root/
  ├─ docx批量文件...
  ├─ bulk_replace_phase1_multi.py
  ├─ README_bulk_replace_docx_dates.md
  └─ _phase1_report_multi/   # 每次运行生成/更新
```

---

## 4. 快速开始（核心流程）
1. 安装依赖：`pip install python-docx`
2. 放置脚本到目标根目录（或任意位置，但 ROOT_DIR 指向正确）。
3. 修改脚本配置区：
   - `ROOT_DIR` 指向包含全部 `.docx` 的根路径
   - `EXPECTED_COUNT = 4`
   - `NEW_DATE = "2025-07-31"`
   - `REPLACE_RULES` 中的日期与替换保持一致（已写好）
   - 首次运行保持：`DRY_RUN = True`
4. 运行：`python bulk_replace_phase1_multi.py`
5. 检查 `_phase1_report_multi/summary.txt`、`detail.csv`、`abnormal_list.txt`
6. 确认统计合理 → 改 `DRY_RUN = False` 再运行，执行真实替换（生成 `.bak`）。
7. 抽样打开几个“ok / ok_abnormal”文件验证内容。
8. 需要时清理或归档 `.bak`。

---

## 5. 配置项详解

| 配置项 | 类型 | 示例 | 说明 |
|--------|------|------|------|
| ROOT_DIR | str | r"D:\docs" | 扫描起始根路径（递归） |
| DRY_RUN | bool | True/False | True=只统计；False=实际替换 |
| MAKE_BACKUP | bool | True | 实际替换前生成 `.bak` |
| INCLUDE_HEADERS_FOOTERS | bool | True | 是否处理页眉/页脚 |
| EXPECTED_COUNT | int | 4 | 统一期望次数（用于本例两个日期） |
| NEW_DATE | str | "2025-07-31" | 统一替换后的日期 |
| REPLACE_RULES | list[dict] | 见脚本 | 可添加更多规则 |
| OUTPUT_DIR | str | "_phase1_report_multi" | 结果输出目录 |
| ANY_MISMATCH_MARKS_ABNORMAL | bool | True | 任一规则不符即异常 |

每条 `REPLACE_RULES` 元素：
```python
{
  "target": "2025-08-19",
  "replacement": "2025-07-31",
  "expected": 4  # 或 None 表示不校验出现次数
}
```

---

## 6. 运行输出说明

输出目录：`_phase1_report_multi/`  
文件列表：
| 文件 | 说明 |
|------|------|
| summary.txt | 全局汇总（总文件数、每规则统计、异常数等） |
| detail.csv | 每文件 × 每规则的出现/替换/期望/是否 mismatch |
| abnormal_list.txt | 仅列出异常文件（任一规则出现次数不等于 expected） |
| error_list.txt | 打开 / 备份 / 保存失败的文件（若有） |

---

## 7. detail.csv 字段解释
| 字段 | 含义 |
|------|------|
| file | 文件完整路径 |
| status | `none` / `dry_run_ok` / `dry_run_abnormal` / `ok` / `ok_abnormal` / `error` |
| note | 主要用于标记 mismatch 或错误信息 |
| found_<target> | 该 target 在此文件中原始出现次数 |
| replaced_<target> | 实际替换次数（DRY_RUN 为 0） |
| expected_<target> | 配置的期望次数（或空） |
| mismatch_<target> | 出现次数是否不等于期望（TRUE/FALSE） |

`ok_abnormal` = 有内容被替换，但至少一个 target 不符合期望次数。  
`dry_run_abnormal` = 同上，但处于 DRY_RUN 模式。

---

## 8. 异常判定逻辑
```
for 每个文件:
  对每条规则统计 found
  mismatch = (expected 不为 None 且 found != expected)
  如果 ANY_MISMATCH_MARKS_ABNORMAL = True 且任一 mismatch:
      文件整体标记 abnormal
```

---

## 9. 备份策略
- 条件：`DRY_RUN=False` 且该文件至少有一个 target 出现次数 > 0。
- 生成：`原文件名.docx.bak`
- 内容：原始未修改版本完整复制。
- 恢复：删除当前 `.docx`，将 `.docx.bak` 重命名为原名。

批量恢复（示例 PowerShell）：
```powershell
Get-ChildItem -Recurse -Filter *.bak | ForEach-Object {
  $orig = $_.FullName -replace '\.bak$',''
  if (Test-Path $orig) { Remove-Item $orig }
  Rename-Item $_.FullName $orig
}
```

删除全部 `.bak`（确认无回退需求后）：
```powershell
Get-ChildItem -Recurse -Filter *.bak | Remove-Item
```

---

## 10. 常见操作示例

### 10.1 新增一个日期规则
```python
REPLACE_RULES.append({
  "target": "2025-09-05",
  "replacement": "2025-07-31",
  "expected": 2  # 或 None
})
```

### 10.2 改所有规则的期望次数
改 `EXPECTED_COUNT`，并同步修改已有规则的 expected 字段；或在脚本顶部统一生成：
```python
EXPECTED_COUNT = 5
REPLACE_RULES = [
  {"target": "2025-08-19", "replacement": NEW_DATE, "expected": EXPECTED_COUNT},
  {"target": "2025-08-20", "replacement": NEW_DATE, "expected": EXPECTED_COUNT},
]
```

### 10.3 仅统计不替换
保持 `DRY_RUN = True`。

### 10.4 取消备份
设置：`MAKE_BACKUP = False`（不建议在验证阶段关闭）。

### 10.5 仅标记异常但不替换异常文件
当前逻辑是“异常仍替换”。如要“异常不替换”，可在处理时增加条件（伪代码）：
```python
if mismatch and SKIP_ABNORMAL:
    不执行真正 do_replace
```

### 10.6 输出 JSON（可扩展）
在脚本末尾追加：
```python
import json
with open(os.path.join(OUTPUT_DIR, "detail.json"), "w", encoding="utf-8") as jf:
    json.dump(results, jf, ensure_ascii=False, indent=2)
```

### 10.7 统计其它日期模式（例如所有 2025-08-01~31 但未替换）
增加一个正则扫描模块（示例）：
```python
import re
pattern = re.compile(r"\b2025-08-(0[1-9]|[12]\d|3[01])\b")

# 在 count_occurrences 基础上再额外扫描每段 p.text, 记录非规则 target 的匹配集合
```

---

## 11. 常见问题与排查

| 现象 | 原因 | 解决 |
|------|------|------|
| 运行后文件未改变 | 仍然 DRY_RUN=True | 改为 False 再跑 |
| 生成大量 .bak | MAKE_BACKUP=True 且文件有匹配项 | 正常；确认后再清理 |
| 发现目标日期仍存在 | 日期在文本框/形状/批注；脚本未覆盖 | 需用 `win32com` 版本（Word COM）扩展 |
| 格式(加粗/颜色)丢失 | 段落 run 被合并 | 若需保留，需 run 级替换算法（见扩展） |
| CSV 出现 UTF-8 编码异常 | 使用旧版 Excel | 已用 `utf-8-sig`，正常双击即可；否则用“数据导入” |
| 重复多次运行生成重复备份？ | `.bak` 已存在则脚本当前逻辑不会覆盖 | 如需覆盖，删除旧 `.bak` 后再运行 |

---

## 12. 扩展建议

| 需求 | 方案概要 |
|------|----------|
| 保留原 run 样式 | 遍历 runs，拼接索引表，定位 target 跨 run 分布，局部替换（需要更复杂逻辑） |
| 处理文本框 / 形状 / 批注 | 使用 `win32com.client.Dispatch("Word.Application")` 遍历 `StoryRanges`（包括 `wdTextFrameStory`） |
| 仅替换前 N 次 | 在 `do_replace` 里增加计数阈值 |
| 支持正则替换 | 将规则改为包含 `pattern` 和 `replacement`，使用正则重建段落文本 |
| 阶段合并 | Phase 2 追加更多 REPLACE_RULES，或把每阶段输出合并到一份主明细里 |
| 生成差异报告 | 保存修改前后文本快照（或 diff）到独立文件夹 |
| 并发加速 | 文档数较大时可多进程（I/O 为主，收益有限） |

---

## 13. 性能建议
- 431 份中等大小文档（<1MB），单线程足够（数秒~数十秒）。
- 不必重复运行真正替换；如需追加规则，可基于已替换后的最新文本继续。
- 若频繁迭代规则，保留一次性全量原始备份（压缩打包即可）。

---

## 14. 安全与合规提示
| 项目 | 建议 |
|------|------|
| 原始资料 | 先整体复制到“工作副本”再操作 |
| 备份保留 | 至少在验收通过前保留 `.bak` |
| 版本留痕 | 将 `summary.txt` 与脚本版本打包存档 |
| 复核 | 随机抽查正常文件与异常文件各至少 3 份 |

---

## 15. 快速核查清单（Checklist）
| 步骤 | 完成? |
|------|-------|
| ROOT_DIR 正确且指向副本目录 |
| DRY_RUN=True 预演结果符合预期 |
| detail.csv 中两个目标日期统计合理 |
| abnormal_list.txt 列表已人工确认 |
| 切换 DRY_RUN=False 前已备份原始整体目录（可选） |
| 正式替换后抽样打开验证内容正确 |
| 决定是否清理 `.bak` |
| 归档 summary.txt + detail.csv 作为留痕 |

---

## 16. 英文简要摘要（可选）
This script batch-processes multiple .docx files to replace predefined date strings, validates expected occurrence counts, marks anomalies, and produces detailed reports (summary, CSV, abnormal list, error list). It supports headers/footers, backups, dry-run mode, and is easily extensible with new rules or regex scanning.

---

## 17. 版本记录（建议手动维护）
| 日期 | 修改人 | 变更 |
|------|--------|------|
| 2025-08-XX | 初次生成 | Phase 1 两日期规则 |
| (后续补) |  |  |

---

## 18. 后续 Phase 2 提示
若后续还有“另外 2 处”不同类型日期或字段：
1. 确认新日期文本 / 替换目标 / 期望次数。
2. 追加至 `REPLACE_RULES`。
3. DRY_RUN 再次预演。
4. 必要时改为正则（如模式化日期群组）。

---

如需：  
- run 级保留格式版本  
- win32com 版本（支持文本框/形状）  
- 正则批量扫描  
请在后续迭代按需添加。

祝使用顺利。如需要进一步扩展或生成第二阶段脚本说明，请继续提出需求。
