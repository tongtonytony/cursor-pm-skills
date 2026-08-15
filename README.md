# Cursor 项目管理 Agent Skills

一套用于 Cursor IDE 的 Agent Skills，覆盖项目管理文档生成、培训课件制作与 PPT 转短视频等场景。基于「AI赋能项目管理」课程方法论构建。

## Skills 列表

| 目录 | 名称 | 用途 |
|------|------|------|
| `project-charter-generation` | 项目章程生成 | 生成项目章程 / Project Charter |
| `business-case-generation` | 商业论证生成 | 生成商业论证 / Business Case |
| `requirement-specification-generation` | 需求规格说明书 | 生成 PRD、需求矩阵、用例图 |
| `outline-design-specification-generation` | 概要设计说明书 | 生成概要设计，含 PlantUML 架构图 |
| `ai-pm-training-ppt` | AI赋能项目经理培训PPT | 基于培训大纲生成 AI 项目管理培训课件 |
| `it-service-pm-training-ppt` | IT服务管理培训PPT | V6 标准，IT 服务管理与项目管理实战课件 |
| `pptx-to-short-video` | PPT转短视频 | 将 PPT 转为 Marp 播报文件，供剪映配音 |

## 安装方式

将本仓库克隆或下载后，把各 Skill 目录复制到 Cursor 个人 Skills 目录：

**Windows**
```
%USERPROFILE%\.cursor\skills\
```

**macOS / Linux**
```
~/.cursor/skills/
```

例如：
```
~/.cursor/skills/project-charter-generation/
~/.cursor/skills/business-case-generation/
...
```

每个 Skill 目录内应包含 `SKILL.md` 及配套的 `reference.md`、脚本等文件。

## 使用方式

在 Cursor 对话中提及对应关键词即可自动触发，例如：

- 「帮我写一份项目章程」→ `project-charter-generation`
- 「生成商业论证」→ `business-case-generation`
- 「编写 PRD」→ `requirement-specification-generation`
- 「制作 AI 赋能项目经理培训 PPT」→ `ai-pm-training-ppt`

## 依赖

部分 Skill 含 Python 脚本，使用前请安装对应 `scripts/requirements.txt` 中的依赖：

```bash
pip install -r ai-pm-training-ppt/scripts/requirements.txt
pip install -r it-service-pm-training-ppt/scripts/requirements.txt
pip install -r pptx-to-short-video/scripts/requirements.txt
```

## 目录结构

```
cursor-pm-skills/
├── README.md
├── project-charter-generation/
│   ├── SKILL.md
│   ├── reference.md
│   └── input-form.html
├── business-case-generation/
├── requirement-specification-generation/
├── outline-design-specification-generation/
├── ai-pm-training-ppt/
│   ├── SKILL.md
│   ├── scripts/
│   └── reference.md
├── it-service-pm-training-ppt/
│   └── scripts/
└── pptx-to-short-video/
    ├── scripts/
    └── themes/
```

## 作者

liutong2027@gmail.com

## License

MIT
