# 参考：insert_plan 与源映射

## insert_plan.json 格式

```json
{
  "template_marker": "XXXX（插）",
  "inserts": [
    {
      "source_key": "file4",
      "source_page": 14,
      "anchor_title": "信息化项目在组织中的定位",
      "reason": "补充项目组织语境：项目与组织架构关系图",
      "mode": "script",
      "final_title": "信息化项目在组织中的定位"
    },
    {
      "source_key": "6-IT服务管理",
      "source_page": 51,
      "anchor_title": "价值交付的全生命周期",
      "reason": "ITIL4 服务价值系统 SVS 全景",
      "mode": "com-clone",
      "final_title": "价值交付的全生命周期"
    }
  ]
}
```

| 字段 | 说明 |
|------|------|
| `source_key` | 见 source_map 键名 |
| `source_page` | 1-based 页码 |
| `anchor_title` | 插入到该标题页**之后** |
| `mode` | `script`（4A）或 `com-clone`（4B，跳过脚本） |
| `final_title` | 定稿标题（不含「（插）」） |

## source_map 默认键

```json
{
  "3-实战篇": "3-实战篇",
  "6-IT服务管理": "6-IT服务管理",
  "1-启航篇": "1-启航篇",
  "2-工程篇": "2-工程篇",
  "file4": "_4_pm_training_temp.pptx"
}
```

`file4` 需要先运行 `convert_ppt_to_pptx.py` 转换 `4-如何做好项目管理*.ppt`。

## 保护页规则

默认保护（仅改备注）：

- 第 1 页封面
- 目录、讲师介绍
- 标题匹配 `第.*章|Chapter|方法论` 且 shape 数 ≤ 8 的过渡页
- 用户追加列表

## 常见问题

### PPT 打开报 rId 损坏

**原因**：跨文件复制 slide XML。  
**修复**：仅用同文件 `duplicate_slide` + 抽图/抽文，或 COM 克隆。

### 插入后无蓝色分割线

运行：`python scripts/revise_from_spec.py --ppt ... --phase fix-dividers`

### 页码漂移

插入按 anchor 索引**降序**执行；每批插入后重新 verify。

### 脚本图例模糊

复合图改 Phase 4B；单图检查源页 pic 尺寸，必要时手工粘贴。

## V6 量化基准

| 指标 | V6 |
|------|-----|
| 总页数 | 251 |
| 含图例页 | 96 |
| 备注均值 | 227 字 |
| ≥200 字 | 217/251 (86%) |
| PREP | 204/251 (81%) |
| 外露（插） | 0 |
