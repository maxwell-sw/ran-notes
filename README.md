<div align="center">
  <img src="assets/ran-banner.svg" width="100%" alt="研行记 RAN · Research Action Note" />

  <br />

  [English](README_EN.md) · 中文

  <p><strong>从多模态组会材料，到可追溯的研究行动。</strong></p>
  <p>一个面向科研团队的组会工作流 Agent：整合转写、PPT、文献与讨论内容，生成结构化纪要、可核查证据与跨组会行动项闭环。</p>

  <img src="https://img.shields.io/badge/Python-3.9%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.9+" />
  <img src="https://img.shields.io/badge/FastAPI-Backend-009688?style=flat-square&logo=fastapi&logoColor=white" alt="FastAPI" />
  <img src="https://img.shields.io/badge/LLM-OpenAI--compatible-5B5BD6?style=flat-square" alt="OpenAI-compatible LLM" />
  <img src="https://img.shields.io/badge/License-MIT-1f2937?style=flat-square" alt="MIT License" />
  <a href="https://ran-notes.onrender.com"><img src="https://img.shields.io/badge/在线演示-Render-46E3B7?style=flat-square&logo=render&logoColor=white" alt="在线演示" /></a>
</div>

---

## 在线体验

访问 [研行记在线演示](https://ran-notes.onrender.com)，无需安装即可体验完整工作流：

1. 点击"查看演示"，进入交互式 Demo 页面。
2. 载入内置"猪周期"研究组会样例，体验材料载入、AI 生成与归档链路。
3. 公共演示版使用服务端配置的模型额度；上传自己的数据时，请在 AI 设置中使用自己的 OpenAI 兼容服务 Key。

> 免费实例在无活动时会休眠，首次访问可能需要等待 30-50 秒冷启动。

## Why RAN

科研组会的价值往往沉没在语音转写、PPT、论文和零散讨论里。研行记（**R**esearch **A**ction **N**ote, RAN）将这些输入组织为一条可复核的研究工作流：明确来源、抽取讨论要点、生成待办，并将行动项带入后续组会。

> **Product thesis**：组会不是一次性的“总结文本”，而是可以被持续查询、追踪和推进的研究资产。

## What it can do

| 从输入到交付 | 具体能力 |
| :-- | :-- |
| **多模态材料接入** | 支持转写文本、DOCX、PPT/PPTX、PDF；可整理语音转写稿并保留人工核对环节。 |
| **结构化组会纪要** | 生成按汇报人/研究主题组织的摘要、导师反馈、风险与后续行动项。 |
| **证据可追溯** | 将关键判断与输入材料关联，避免只给“看似合理”的无来源总结。 |
| **行动项闭环** | 记录负责人、期限、优先级、状态与回应情况；支持后续编辑。 |
| **跨组会研究记忆** | 提供资料库、关键词搜索、日历热力图和未完成行动项视图。 |
| **一键演示案例** | 内置“猪周期”研究组会样例，可完整体验材料载入、生成与归档链路。 |

## Agent workflow

```text
PPT / PDF / DOCX / 转写文本
              ↓
      解析、清洗与人工确认
              ↓
   LLM 结构化推理与受约束输出
              ↓
纪要 · 导师反馈 · 风险 · 行动项 · 证据
              ↓
  本地资料库 · 跨组会追踪 · 导出交付
```

RAN 将 LLM 放在受控的工作流中：模型负责理解和归纳，应用负责材料解析、字段约束、文件归档、行动项状态与资料库检索。这样既保留自然语言交互的效率，也让结果能够被回看、编辑和继续推进。

## Quick start

### macOS：双击启动（推荐演示方式）

1. 在 `ran-backend/.env` 中配置模型服务，参考 `ran-backend/.env.example`。
2. 在 Finder 双击 [启动研行记.command](启动研行记.command)。
3. 首次运行会自动建立 Python 环境并安装依赖；浏览器随后自动打开主页。
4. 点击“立即试用”，载入内置案例后即可生成完整组会纪要。

> 启动终端保持打开；关闭该终端即停止本地服务。

### 手动启动

```bash
cd ran-backend
python3 -m venv .venv
./.venv/bin/python -m pip install -r requirements.txt
./.venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8003
```

在另一个终端运行：

```bash
cd "ran-page 3"
python3 -m http.server 8081
```

访问 `http://127.0.0.1:8081`。

## Model configuration

复制并填写 `ran-backend/.env`：

```dotenv
OPENAI_API_KEY=your_key
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4o
```

支持 OpenAI 兼容接口。模型选择建议见 [模型推荐.md](ran-backend/模型推荐.md)。

## Repository map

```text
ran-notes/
├── ran-page 3/             # 产品主页与交互式 Demo
├── ran-backend/
│   ├── main.py             # FastAPI、材料解析、LLM 编排、资料库接口
│   ├── trial_assets/       # 内置演示素材
│   ├── requirements.txt    # Python 依赖
│   └── .env.example        # 模型配置模板
├── 测试材料_三人组会/        # 可手动上传的测试包
└── 启动研行记.command       # macOS 一键启动入口
```

## Privacy & boundaries

- API Key 只从本地 `.env` 读取，已被 Git 忽略，不会提交到仓库。
- 组会生成记录与用户上传材料默认保存在本机资料库，已被 Git 忽略。
- 生成结果用于辅助梳理和推进研究，不应替代研究者对实验、数据、引用与结论的最终判断。
- 演示材料仅用于产品体验；使用自己的数据时，请确保拥有相应的处理和分享权限。

## Roadmap

- [x] 多来源材料解析与结构化纪要
- [x] 行动项闭环与本地资料库
- [x] 内置一键演示案例
- [ ] 更细粒度的引文定位与证据编辑
- [ ] 团队协作与权限管理
- [ ] 可部署的多用户版本

## License

Released under the [MIT License](LICENSE).

## Contact

Built by [@maxwell-sw](https://github.com/maxwell-sw). Issues and pull requests are welcome.
