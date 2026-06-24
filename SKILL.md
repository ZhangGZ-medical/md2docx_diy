# md2docx_diy

Markdown文档 → 中文DOCX格式转换工具。

## 功能说明

将Markdown文件转换为排版精良的A4 Word文档（支持横版/纵向），完整支持中文格式和各类Markdown语法。

## 支持的格式

| Markdown语法 | DOCX输出效果 |
|---|---|
| `#` ~ `######` 标题 | 1-6级标题，黑体，深蓝色，带底边线 |
| `**粗体**` | 黑体加粗 |
| `*斜体*` | 楷体斜体 |
| `***粗斜体***` | 黑体加粗斜体 |
| `` `代码` `` | Courier New蓝色 |
| `---` 分隔线 | 浅蓝横线 |
| `>` 引用块 | 浅蓝背景 + 蓝色左边框 |
| ` ```code``` ` 代码块 | 灰色背景代码框 |
| `- bullet` 列表 | 项目符号，蓝色圆点 |
| `\| table \|` 表格 | 深蓝表头 + 交替行背景色 + 单元格边距 |
| `[text](#anchor)` 链接 | 保留文本，去除链接标记，生成Word书签 |
| `<a name="xxx">` 锚点 | 生成Word书签（保留跳转功能） |

## 排版规格

- **页面方向**：A4（横版 29.7cm×21.0cm / 纵向 21.0cm×29.7cm，通过 `orientation` 参数控制）
- **正文字体**：宋体 10.5pt
- **标题字体**：黑体（各级递减 16pt → 9pt）
- **页眉**：文件名 + 团队名称（居中）
- **页脚**：文件名 + 页码
- **行距**：1.15倍
- **单元格边距**：上下0.15cm，左右0.2cm

## 使用方法

### 方法1：转换指定MD文件

在脚本中修改 `FILES` 列表：

```python
FILES = [
    (r'd:\path\to\input.md', r'd:\path\to\output.docx'),
]
python md2docx_diy.py
```

### 方法2：直接调用函数

```python
from md2docx_diy import md_to_docx
md_to_docx(r'd:\path\to\input.md', r'd:\path\to\output.docx')
# 纵向排版：
md_to_docx(r'd:\path\to\input.md', r'd:\path\to\output.docx', orientation='portrait')
```

## 注意事项

- 仅支持Windows（依赖SimHei/SimSun/楷体系统字体）
- Python依赖：`pip install python-docx`
- 输入MD文件必须为UTF-8编码
- 表格自动等宽分配列宽

## 文件结构

```
md2docx_diy/
├── SKILL.md          # 本说明文件
└── md2docx_diy.py    # 主转换脚本
```
