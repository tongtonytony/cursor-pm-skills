# 概要设计说明书结构参考

本文档描述默认输出的文档结构，参考「概要设计说明书模板样例」「基于RAG的AI智能客服项目交付样例（业务架构图+架构概览图+组件图+部署图）」及需求规格说明书。

---

## 一、引言

### 1.1 目的
说明本文档的编制目的及读者对象（如：开发人员、测试人员、项目经理）。

### 1.2 范围
说明本概要设计涵盖的系统范围，与需求规格说明书中的范围保持一致。

### 1.3 定义与缩写
列出文档中使用的术语、缩写及定义。

### 1.4 参考资料
列出需求规格说明书、相关法规、标准、上级文档等。

---

## 二、总体设计

### 2.1 设计目标
从需求规格说明书中提炼的设计目标（如：高可用、可扩展、易维护）。

### 2.2 设计原则
遵循的设计原则（如：模块化、松耦合、高内聚）。

### 2.3 技术选型
- 开发语言与框架
- 数据库与中间件
- 部署与运维技术

---

## 三、业务架构图

### 3.1 图说明
描述业务架构图展示的业务模块、角色、流程关系。

### 3.2 PlantUML 脚本示例（业务架构图）

```plantuml
@startuml 业务架构图
!theme plain
skinparam backgroundColor #FEFEFE
skinparam defaultFontName 微软雅黑

title 业务架构图 - 基于RAG的AI智能客服系统

actor "客服人员" as agent
actor "系统管理员" as admin
actor "外部用户" as user

rectangle "业务层" {
  package "智能问答" {
    [问题输入] as input
    [答案展示] as output
  }
  package "知识库管理" {
    [文档上传] as upload
    [知识检索] as search
  }
  package "会话管理" {
    [会话记录] as session
    [历史查询] as history
  }
}

rectangle "外部系统" {
  [企业微信] as wework
  [工单系统] as ticket
}

agent --> input : 使用
agent --> output : 查看
agent --> session : 查询
admin --> upload : 维护
admin --> search : 配置
user --> wework : 接入
input --> search : 调用
output --> session : 记录
@enduml
```

---

## 四、架构概览图

### 4.1 图说明
描述系统分层结构、模块划分、主要技术组件。

### 4.2 PlantUML 脚本示例（架构概览图）

```plantuml
@startuml 架构概览图
!theme plain
skinparam backgroundColor #FEFEFE
skinparam defaultFontName 微软雅黑

title 架构概览图 - 基于RAG的AI智能客服系统

together {
  rectangle "展示层" as UI #E8F4FD {
    [Web前端] as web
    [企业微信集成] as wechat
  }
}

together {
  rectangle "应用层" as APP #D4EDDA {
    [智能问答服务] as qa
    [知识库管理服务] as kb
    [会话管理服务] as sess
  }
}

together {
  rectangle "能力层" as CORE #FFF3CD {
    [RAG检索] as rag
    [大模型调用] as llm
    [向量数据库] as vecdb
  }
}

together {
  rectangle "数据层" as DATA #F8D7DA {
    [关系数据库] as rdb
    [文档存储] as doc
  }
}

UI --> APP : HTTP/API
APP --> CORE : 调用
CORE --> DATA : 读写
@enduml
```

---

## 五、组件图

### 5.1 图说明
描述核心组件及其依赖、接口关系。

### 5.2 PlantUML 脚本示例（组件图）

```plantuml
@startuml 组件图
!theme plain
skinparam backgroundColor #FEFEFE
skinparam defaultFontName 微软雅黑

title 组件图 - 基于RAG的AI智能客服系统

package "智能客服应用" {
  component [问答接口] as QA_API
  component [知识库接口] as KB_API
  component [会话接口] as Sess_API
  
  component [问答引擎] as QA_Engine
  component [知识库服务] as KB_Service
  component [会话服务] as Sess_Service
  
  component [RAG检索组件] as RAG
  component [LLM调用组件] as LLM
}

database "向量数据库" as VDB
database "关系数据库" as RDB

QA_API --> QA_Engine : 调用
KB_API --> KB_Service : 调用
Sess_API --> Sess_Service : 调用

QA_Engine --> RAG : 检索
QA_Engine --> LLM : 生成
KB_Service --> RAG : 写入
Sess_Service --> RDB : 读写
RAG --> VDB : 查询
@enduml
```

### 5.3 组件说明表

| 组件名称 | 职责 | 依赖 |
|----------|------|------|
| 问答接口 | 接收用户问题，返回答案 | 问答引擎 |
| 知识库接口 | 文档上传、知识管理 | 知识库服务 |
| 会话接口 | 会话记录、历史查询 | 会话服务 |
| 问答引擎 | 编排 RAG 与 LLM，生成回答 | RAG、LLM |
| RAG检索组件 | 向量检索、上下文组装 | 向量数据库 |

---

## 六、部署图

### 6.1 图说明
描述部署节点、运行环境、网络拓扑。

### 6.2 PlantUML 脚本示例（部署图）

```plantuml
@startuml 部署图
!theme plain
skinparam backgroundColor #FEFEFE
skinparam defaultFontName 微软雅黑

title 部署图 - 基于RAG的AI智能客服系统

node "应用服务器" as app_server {
  component [Web应用] as web
  component [API服务] as api
  component [RAG服务] as rag_svc
}

node "数据库服务器" as db_server {
  database [MySQL] as mysql
  database [向量库] as vecdb
}

node "外部服务" as ext {
  component [大模型API] as llm_api
}

cloud "用户端" {
  component [浏览器] as browser
  component [企业微信] as wechat
}

browser --> web : HTTPS
wechat --> api : HTTPS
web --> api : 内部
api --> rag_svc : 内部
rag_svc --> mysql : 3306
rag_svc --> vecdb : 默认端口
rag_svc --> llm_api : HTTPS
@enduml
```

### 6.3 部署说明表

| 节点 | 组件 | 环境要求 | 说明 |
|------|------|----------|------|
| 应用服务器 | Web应用、API服务、RAG服务 | 4核8G+、Linux | 可容器化部署 |
| 数据库服务器 | MySQL、向量库 | 2核4G+ | 可云托管 |
| 外部服务 | 大模型API | 网络可达 | 按需配置 |

---

## 七、接口设计

### 7.1 主要接口列表

| 接口名称 | 方法 | 路径 | 说明 |
|----------|------|------|------|
| 智能问答 | POST | /api/qa/ask | 接收问题，返回答案 |
| 知识库上传 | POST | /api/kb/upload | 上传文档到知识库 |
| 会话查询 | GET | /api/session/list | 查询历史会话 |

### 7.2 接口详细说明（示例）

**智能问答接口**
- 请求：`{ "question": "string", "session_id": "string" }`
- 响应：`{ "answer": "string", "sources": [] }`

---

## 八、非功能性设计

### 8.1 性能设计
- 响应时间目标
- 并发处理策略
- 缓存策略

### 8.2 安全设计
- 认证与授权
- 数据加密
- 审计日志

### 8.3 可扩展性设计
- 水平扩展方案
- 模块解耦设计

---

## 文档输出建议

- 若输出为 Markdown，可使用上述结构直接生成，PlantUML 代码块便于复制到 [PlantUML 在线](https://www.plantuml.com/plantuml) 或本地工具渲染。
- 若用户提供自定义 Word 模板，按模板章节结构填充，保持业务架构图、架构概览图、组件图、部署图四类 PlantUML 脚本完整。
- 所有内容使用中文，避免错别字与乱码。
