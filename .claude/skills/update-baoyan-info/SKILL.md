---
name: update-baoyan-info
description: 根据用户指定的 Excel 数据文件更新保研信息，并生成对应的推文或文档 (docx) 内容，同时将新增数据填写到用户指定的微信共享表格中。当用户要求"更新保研院校信息共享表格"或"将院校信息更新表格转换为word"时触发此技能。
---

# 更新保研记录并同步微信表格 (Update Baoyan Info)

该技能用于处理本地抓取后的 Excel 数据文件，自动生成对应的 Word (.docx) 文稿或推文素材，并将 Excel 中的最新数据条目同步到用户指定的微信共享表格链接中。

## 使用场景
当用户提供包含最新保研信息的 Excel，并要求输出格式化文档和分享更新时，调用此技能。主要指令包括：
- "根据这个新整理出来的 excel 更新微信共享表格"
- "帮我跑一下 info_output.py 生成今天推文的 docx，顺便把表格更新过去"
- "把 2026院校信息_更新.xlsx 的数据更新到在线表格里去"

## 执行步骤
为了顺利和准确地完成任务，请遵循以下多步执行工作流：

### 1. 确认前置条件与诉求
- 识别工作区目录，检查工作区目录是否包含`2026院校信息_更新.xlsx`。若不存在，提醒用户更新文件不存在，需要先收集院校信息再进行更新。
- 确认用户输入的共享腾讯表格链接，若用户未给出链接，提示用户给出链接。若用户仍未提供链接，新建共享腾讯文档表格进行写入。

### 2. 生成 DOCX 文档
- 请直接在终端中运行内置在当前 Skill 目录下的 `scripts/info_output.py` 脚本，并确保将结果输出在用户给定的工作区目录下：
  ```bash
  python .claude/skills/update-baoyan-info/scripts/info_output.py
  ```
- 检查代码运行日志，向用户确认对应的 `.docx` 是否已成功保存至当前工作区。

### 3. 文件检查
确认当前项目中包含以下文件：
- `2026院校信息_更新.xlsx`：包含最新院校更新增量数据的 Excel 文件，按照sheet分类。
- `2026院校信息.xlsx`：未分类的数据文件，包含所有的保研信息
- `2026院校信息_分类.xlsx`：分类整理后的数据文件，包含按照专业大类分类的保研信息，用于与共享表格对比。Excel中包含四个sheet，分别为"理工农医"、"经管法"、"人文社科与艺术"及"单校"，每个sheet包含对应分类的院校信息。

### 4. 更新腾讯共享表格
使用内置脚本 `scripts/update_tencent_sheet.py` 将本地 Excel 数据追加写入腾讯共享表格。

**环境检查（首次使用或跨设备使用时必做）：**
```bash
python .claude/skills/update-baoyan-info/scripts/update_tencent_sheet.py --check-env
```

**正常使用：**
```bash
python .claude/skills/update-baoyan-info/scripts/update_tencent_sheet.py \
    --excel "2026院校信息_更新.xlsx" \
    --url "https://docs.qq.com/sheet/DWkpjRkZodUF3UG92"
```

**跨设备使用（指定路径）：**
```bash
# 如果自动检测失败，可以手动指定 mcporter 路径
python .claude/skills/update-baoyan-info/scripts/update_tencent_sheet.py \
    --excel "2026院校信息_更新.xlsx" \
    --url "https://docs.qq.com/sheet/DWkpjRkZodUF3UG92" \
    --mcporter /path/to/mcporter
```

**环境变量配置（可选）：**
```bash
# 设置环境变量后无需每次指定路径
export NODE_PATH=/path/to/node/bin        # Node.js bin 目录
export MCPORTER_PATH=/path/to/mcporter    # mcporter 可执行文件路径
```

**脚本功能说明：**
- 自动检测 Node.js 和 mcporter 路径（支持 nvm、volta、homebrew 等多种安装方式）
- 读取本地 Excel 文件的所有 sheet（理工农医、经管法、人文社科与艺术、单校）
- 获取腾讯在线表格的现有数据
- 通过"通知链接"字段进行去重对比，找出新增数据
- 将新增数据追加写入在线表格对应 sheet 的末尾
- **"申请倒计时"和"是否截止"字段使用 Excel 公式实时计算**：
  - 申请倒计时公式：`=IFERROR(IF(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{row},"年","-"),"月","-"),"日",""))<TODAY(),"已截止",DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{row},"年","-"),"月","-"),"日",""))-TODAY()&"天"),"未知")`
  - 是否截止公式：`=IFERROR(IF(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H{row},"年","-"),"月","-"),"日",""))<TODAY(),"是","否"),"未知")`

**依赖要求：**
- Python 3.x + openpyxl 库
- Node.js + mcporter CLI（自动检测或手动指定路径）
- 腾讯文档授权已完成

### 5. 总结与反馈给用户
完成后，向用户报告工作结果：
1. **文档生成结果**：生成的 Word (.docx) 文件名及其存储位置。
2. **表格更新结果**：汇报有多少条记录已被同步至微信共享表格（或已准备好供用户复制）。
3. 提示用户可以打开文件去核对生成的质量。

## 规则与约束
- **数据准确性**：从 Excel 读取与准备同步的数据必须保证不串行、不错列，严格对应学校、学院、专业名称等字段。
- **隐私与权限限制**：如果需要配置腾讯文档、微信开放平台的密钥，切勿将其明文打印在返回的日志或最终交付中。
- **容错处理**：若 `info_output.py` 在运行过程中报错，或遇到共享表格 API 读写被拒问题，停止操作并将关键报错日志返回，向用户提供修复建议。
- **公式字段**：新写入的数据必须使用 Excel 公式计算"申请倒计时"和"是否截止"字段，确保数据实时更新。