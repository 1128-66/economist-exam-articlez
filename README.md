# economist-exam-article

**《经济学人》文风 · 考研英语一阅读练习生成器（WorkBuddy/CodeX Skill）**

输入约 100 个英文单词，一键生成一份排版规范的考研英语一阅读练习包（.docx）：
经济学人风格文章 + 中英对照译文 + 5 道阅读理解题 + 答案解析 + 默写词汇表。

Given a vocabulary list of ~100 English words, this WorkBuddy skill generates a complete
考研英语一 (Chinese postgraduate entrance exam English I) reading-practice package as a
formatted .docx file: an Economist-style article, bilingual translation, 5 exam-style
comprehension questions with answer explanations, and a dictation vocabulary table.

---

## 功能特性（Features）

| 部分 | 内容 |
|---|---|
| **Part 1 · 练习页** | 纯英文原文（无标注，真题排版，独立成页）+ 5 道阅读理解题（21-25，单独成页） |
| **Part 2 · 解析页** | 答案速查 + 每题「答案 + 原文定位 + 干扰项辨析」 |
| **Part 3 · 精读页** | 带注释英文原文（目标词**加粗** + 斜体括号中文释义）+ 逐段中文翻译 |
| **Part 4 · 默写页** | 词汇表（Word / 词性），释义列留空，供默写自测 |

- 文章难度对标考研英语一：450-600 词、4-6 段、每段含长难句
- 经济学人文风：开篇即观点、长短句交替、数据支撑、克制冷幽默
- 5 题覆盖细节 / 推断 / 主旨 / 词义 / 态度五类题型，干扰项按真题手法设计
- docx 排版：Times New Roman + 宋体、1.5 倍行距、字符间距 0.2pt、A4 分页

## 文件结构（Structure）

```
economist-exam-article/
├── SKILL.md                     # 技能主文件：工作流 + JSON 规范
├── references/
│   └── writing-guide.md         # 经济学人文风 · 考研难度 · 命题规范
└── scripts/
    └── build_docx.py            # docx 渲染脚本（python-docx）
```

## 安装（Installation）

### 方式一：直接复制（推荐）
将本仓库 `economist-exam-article/` 目录复制到 `~/.workbuddy/skills/`：

```bash
git clone https://github.com/<your-username>/economist-exam-article.git
cp -r economist-exam-article ~/.workbuddy/skills/
```

### 方式二：安装 zip 包
下载 release 中的 `economist-exam-article.zip`，导入 WorkBuddy 技能中心。

### 依赖
docx 生成脚本需要 `python-docx`（建议安装在隔离的虚拟环境中）：

```bash
python3 -m venv ~/.workbuddy/binaries/python/envs/default
~/.workbuddy/binaries/python/envs/default/bin/pip install python-docx
```

## 使用方法（Usage）

直接向 WorkBuddy 发送约 100 个单词（可附带指定板块/主题），技能会自动触发并输出 docx。
例如：

> 以下是我要背的 100 个单词：[……]，请生成一篇经济学人风格的文章，板块选商业。

脚本亦可独立使用：

```bash
python3 scripts/build_docx.py content.json -o output.docx
```

`content.json` 格式见 `SKILL.md` 中的「JSON 输入格式」一节。

## License
