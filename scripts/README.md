# scripts 使用说明

两个独立脚本，处理任意一本书时复制复用，**不修改本目录原件**。

## 依赖

```bash
pip install PyMuPDF python-docx
```

Windows 中文环境建议所有运行前加 `PYTHONUTF8=1`，并用 `python`（本机 `python3` 可能缺依赖）。

## 1. check_pdf.py — PDF 文本层检测

决定提取路线（有文本层）还是知识化路线（扫描版）。

```bash
PYTHONUTF8=1 python check_pdf.py "D:/大学/推免/语言学知识/某书.pdf"
# 文件名带特殊字符（括号/破折号/空格）时用 glob 通配：
PYTHONUTF8=1 python check_pdf.py "D:/大学/推免/语言学知识/某书*.pdf"
```

输出：页数、含文本层页数、全书字符数、三档判定（TEXT / SCAN / MIXED）、前 3 个有文本页面的样本。

**判定规则**：
- ≥80% 页有文本且全书 >5000 字符 → TEXT（逐章提取，证据索引可到页）
- 0 页有文本 → SCAN（知识化处理，README 声明）
- 其余 → MIXED（混合处理）

## 2. md_to_docx.py — Markdown → Word 转换

1. 复制到 `{书名}_详细梳理/` 文件夹
2. 修改文件顶部两个变量：

```python
BASE_DIR = r"D:/大学/推免/语言学知识/{书名}_详细梳理"
FILES = [
    ("README.md", "README.docx"),
    ("01_知识图谱/xxx.md", "01_知识图谱/xxx.docx"),
    ("02_结构化提取/xxx.md", "02_结构化提取/xxx.docx"),
    ("03_学习笔记/xxx.md", "03_学习笔记/xxx.docx"),
]
```

3. 运行：

```bash
cd "D:/大学/推免/语言学知识/{书名}_详细梳理" && PYTHONUTF8=1 python md_to_docx.py
```

### 已内置的渲染规则

| Markdown 元素 | Word 呈现 |
|---|---|
| `#`~`######` 标题 | Word 标题层级；h1 居中 |
| `\|` 表格 | Table Grid 实线表格，首行加粗 |
| `**粗体**` / `*斜体*` / `` `行内代码` `` | 对应 run 格式；行内代码 Consolas 9.5pt |
| `> 引用` | ▎前缀 + 灰色文字，左缩进 |
| ` ``` ` 代码块 | Consolas 9pt 深灰，左缩进 |
| `- ` / `1. ` 列表 | List Bullet / List Number 样式 |
| `---` 水平线 | 60 个 ─ 居中 |

### 已知注意点

- 转换前关闭 Word 中打开的对应 .docx（避免 `~$` 锁文件冲突）
- 表格单元内暂不支持行内格式（粗体等），标题行由 `bold=(row_idx==0)` 处理
- 字体统一微软雅黑；正文 11pt、表格 9pt
