---
name: it-service-pm-training-ppt
description: >-
  基于培训大纲、基线PPT与多源课件池，按V6质量标准生成「IT服务管理与项目管理实战」类培训PPT。
  含叙事链设计、模板页插入、PREP备注200+、复合图例COM克隆与质检闭环。
  当用户制作IT服务管理/项目管理实战培训PPT、修订V4/V5/V6类课件、
  或提及「IT服务管理培训」「项目管理实战PPT」「PREP备注」「插入页模板」「PPT内容规格」时自动触发。
---

# IT服务管理与项目管理实战培训PPT生成技能（V6标准）

以 **V6 定稿** 为质量基准，通过 **六阶段流水线**（提取→规格→骨架→插入→备注→策展质检）产出约 250 页、高图例密度、PREP 备注 ≥200 字的 1 天培训课件。

> V6 质量 ≠ 一次脚本跑完。自动化约 40%，**复合图例粘贴 + 叙事策展** 约 60%——本 Skill 把两者编排成可重复流程。

## 触发场景

- 新建/修订「IT服务管理与项目管理实战」培训 PPT
- 从 V3/V4/V5 迭代到 V6 水平
- 从多源课件（实战篇、标准版、ITSM）组装 1 天课程
- 批量 PREP 备注、模板页插入、质量验收

## 前置：确认源文件

| 文件 | 用途 | 必需 |
|------|------|------|
| 培训大纲（`.docx` 或 `_outline_chapters.json`） | 章节、时长、知识点 | ✅ |
| 基线 PPT（如 `…V3.pptx`） | 主线骨架、章节过渡页 | ✅ |
| 空白插入模板页（标题含 `XXXX（插）`） | 统一版式 | ✅ |
| 源课件池（见下表） | 图例/案例来源 | ✅ |
| Gold Standard（如 `…V6.pptx`） | 质量参照 | 推荐 |

**源课件优先级**（匹配时按此顺序）：

1. `3-实战篇` — 案例、章程、变更、收尾
2. `4-如何做好项目管理`（`.ppt` 需先转 `.pptx`）— 框架图、PMBOK、需求
3. `6-IT服务管理` — ITIL、运维流程、SLA
4. `1-启航篇` / `2-工程篇` — 补充概念
5. `8-课程模板` — 仅版式，非内容

若路径缺失，询问用户或将文件放入工作区。

---

## 六阶段工作流

```
[1 提取索引] → [2 内容规格] → [3 骨架保留] → [4 分轨插入] → [5 PREP备注] → [6 策展定稿]
                                      ↓              ↓
                              保护页不动      A轨:脚本单图  B轨:COM整页克隆(人工)
```

**进度清单**（复制跟踪）：

```
- [ ] Phase 1: 提取大纲 + 索引源课件
- [ ] Phase 2: 产出 PPT内容规格_Vx.md 并获用户确认
- [ ] Phase 3: 复制基线 → 工作副本，标记保护页
- [ ] Phase 4A: 单图页脚本批量插入
- [ ] Phase 4B: 复合图例页 COM 克隆（或用户 PowerPoint 手工粘贴）
- [ ] Phase 5: PREP 备注 200+ 批量生成
- [ ] Phase 6: verify_quality → 策展删冗 → 去掉「（插）」→ 定稿
```

---

## Phase 1：提取与索引

```bash
pip install -r scripts/requirements.txt

# 提取大纲 + 基线PPT + 源课件池摘要
python scripts/extract_sources.py <大纲docx> <基线pptx> <工作目录>

# 索引源课件：标题、页码、图片数、形状数、是否复合图
python scripts/index_source_slides.py <工作目录> [--glob "*.pptx"]
```

**输出**（在工作目录）：

- `outline_extracted.txt`
- `baseline_extracted.txt`
- `source_slide_index.json` — 每页 `{file, page, title, pic_count, shape_count, composite}`

**复合图判定**：`shape_count >= 15` 或含 GROUP/LINE → `composite: true` → **走 Phase 4B**，不走脚本抽图。

**`.ppt` 转换**（Windows + 已装 PowerPoint）：

```bash
python scripts/convert_ppt_to_pptx.py <输入.ppt> <输出.pptx>
```

---

## Phase 2：PPT 内容规格（强制 Gate）

**未产出并确认规格前，禁止运行插入/修订脚本。**

使用 [spec-template.md](spec-template.md) 生成 `PPT内容规格_Vx.md`。每页必须包含：

| 字段 | 说明 |
|------|------|
| 页码/序号 | 规划页序 |
| 标题 | 定稿标题（无「（插）」） |
| 页面角色 | 见 [page-roles.md](page-roles.md) |
| 叙事链 | 所属链条，如「定位→价值→生命周期」 |
| 大纲章节 | 对应 `_outline_chapters.json` 条目 |
| 来源 | `基线保留` / `源:实战篇p16` / `新建` |
| 插入方式 | `keep` / `script` / `com-clone` / `manual` |
| 备注要点 | PREP 四段关键词 |

### V6 开篇叙事链（Gold Pattern，第 4–22 页）

新规格必须复现或超越此结构：

| 页段 | 主题 | 角色 |
|------|------|------|
| P4–5 | 信息化项目在组织中的定位 | 概念 + 图例 |
| P6–8 | 信息化项目在组织中的价值 | 概念 + 多页图例 |
| P9–10 | 价值交付的全生命周期 | 概念 + 图例 |
| P11 | 项目与运维的交付方式迭代演进 | 过渡深化 |
| P12 | 组织目标与 IT 目标一致性研讨 | 研讨页 |
| P13–14 | 价值交付的流水线 | 双页图例 |
| P15–18 | 如何做好项目管理? | 框架 + 能力模型 |
| P19–22 | 项目经理的能力评估 | 多页展开 |

**同一标题出现 3–9 次是 Feature**（概念页 + 图例页 + 案例页），不是错误。

规格产出后：**请用户确认**，再进入 Phase 3。

---

## Phase 3：骨架保留

1. 复制基线 PPT 为 `…Vx_work.pptx`（不直接改基线）
2. **保护页**（仅可改备注，不可改版式/正文）：
   - 封面、目录、讲师介绍
   - 各章过渡页（含 `Chapter` / `第X章`）
   - 用户明确指定的页（如 V4 的 2、3、53、54）
3. 删除规格中标记为「不保留」的冗余页（从后向前删）

---

## Phase 4：分轨插入

### 4A — 脚本插入（单图 / 纯文本页）

编辑 `insert_plan.json`（格式见 [reference.md](reference.md)），然后：

```bash
python scripts/revise_from_spec.py \
  --ppt <工作副本.pptx> \
  --plan insert_plan.json \
  --phase insert
```

脚本行为：

- 从 `XXXX（插）` 模板页 **duplicate_slide**（同文件复制，避免 rId 损坏）
- 填充标题（微软雅黑 30pt）、蓝色分割线、内容图/文
- 按 anchor 标题 **降序插入**，避免页码漂移
- 跳过已存在同标题页

**禁止**：跨文件 `deepcopy` 整张 slide XML（必损坏 rId）。

### 4B — COM 整页克隆（复合图例，V6 关键）

ITIL 生命周期、PMBOK 49 过程、组织框架等 **40+ 形状** 的页，脚本无法保真。

```bash
python scripts/clone_slides_com.py \
  --target <工作副本.pptx> \
  --source <源课件.pptx> \
  --pages 30,51,117 \
  --after-title "价值交付的全生命周期"
```

或在 PowerPoint 中：**源页全选复制 → 目标位置粘贴 → 仅改标题与备注**。

生成 `composite_pages_todo.md` 清单，逐项勾选。Phase 4B 完成前不进入 Phase 5。

---

## Phase 5：备注标准化

### 5A — PREP 结构（讲师手册风格）

```bash
python scripts/revise_from_spec.py \
  --ppt <工作副本.pptx> \
  --phase notes \
  --min-chars 200
```

### 5B — 口播稿（短视频/配音/逐页讲解风格）

将每页备注转为 **自然口语**，无 PREP 标签，适合剪映配音或现场照读：

```bash
python scripts/broadcast_notes.py \
  --input <源.pptx> \
  --output <输出版.pptx> \
  --min-chars 200 \
  --course-name "IT服务管理与项目管理实战" \
  --report broadcast_report.json
```

口播稿按页面角色自动选用话术：封面 / 章节过渡 / 概念 / 图例 / 案例 / 研讨 / Q&A / 封底。

**PREP 格式**（5A 专用，固定四段标签）：

```
【P-观点】…。
【R-理由】…。
【E-例证】…。
【P-重申】…。
```

插入页备注首行保留 `插入理由：…`（定稿时可删标题「（插）」，理由保留）。

按页面角色选用不同 PREP 话术 — 见 [page-roles.md](page-roles.md)。

---

## Phase 6：策展定稿

```bash
python scripts/verify_quality.py <工作副本.pptx> --gold <V6.pptx可选>
```

**V6 质量阈值**（1 天 / ~250 页）：

| 指标 | 目标 |
|------|------|
| 总页数 | 230–260 |
| 含图例页 | ≥90 |
| 备注 ≥200 字 | ≥85% 页面 |
| PREP 结构 | ≥80% 页面 |
| 空备注 | 0（封面/模板页除外） |
| 标题-内容错位 | 0 |
| PPT 打开 | 无 rId 损坏提示 |
| 外露「（插）」 | 0（定稿版） |

**策展动作**（人工，不可脚本替代）：

1. 删低价值重复页（V6 比 V5 少约 8 页）
2. 去掉标题「（插）」，融入主线
3. 检查图例与标题语义一致（V4 P8 曾错位）
4. 走查开篇 P4–22 叙事是否连贯
5. 另存为 `…Vx.pptx`

完整质检清单：[quality-gate.md](quality-gate.md)

---

## 版式契约（模板页）

| 元素 | 规范 |
|------|------|
| 标题字体 | Microsoft YaHei 30pt |
| 正文字号 | 17–24pt |
| 标题位置 | left≈1409700, top≈152400 EMU |
| 蓝色分割线 | 标题下 ~600000 EMU, 色 `#006BA6` |
| 单图内容区 | left≈838200, top≈1000000, max 9600000×5200000 |
| Logo | 右上角，插入时不删 |

---

## 互动确认（Phase 2 前询问）

1. **演讲时长**（默认 1 天 ≈ 250 页，~1.5 分钟/页）
2. **基线版本**（V3/V4/其他）
3. **保护页列表**（默认 + 用户追加）
4. **是否提供 V6 作 Gold Standard**
5. **定稿是否去掉「（插）」**（默认是，V6 风格）

---

## 参考文件

| 文件 | 内容 |
|------|------|
| [reference.md](reference.md) | insert_plan 格式、源映射、常见问题 |
| [page-roles.md](page-roles.md) | 页面角色与 PREP 话术 |
| [spec-template.md](spec-template.md) | 规格文档模板 |
| [quality-gate.md](quality-gate.md) | 定稿质检清单 |
| `scripts/` | 提取、索引、插入、备注、验证、COM 转换 |

## 关联技能

- `pptx-to-short-video` — 定稿后转短视频素材
- `business-case-generation` / `project-charter-generation` — 涉及交付物模板时可引用

## 禁止事项

1. 不臆造大纲章节
2. 不跨文件 deepcopy slide（防 rId 损坏）
3. 不改保护页版式与正文
4. 不跳过规格文档
5. 不把「重复标题」当 anomaly 删除
