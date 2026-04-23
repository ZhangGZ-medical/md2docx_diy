# md2docx_diy

Markdown → 横版中文 Word 文档转换工具

## 功能

将 Markdown 文件转换为排版精良的横版 A4 Word 文档，完整支持中文格式和各类 Markdown 语法。

## 支持的格式

| Markdown 语法 | DOCX 输出效果 |
|--------------|---------------|
| `#` ~ `######` 标题 | 1-6级标题，黑体，深蓝色，带底边线 |
| `**粗体**` | 黑体加粗 |
| `*斜体*` | 楷体斜体 |
| `***粗斜体***` | 黑体加粗斜体 |
| `` `代码` `` | Courier New 蓝色 |
| `---` 分隔线 | 浅蓝横线 |
| `>` 引用块 | 浅蓝背景 + 蓝色左边框 |
| ` ```code``` ` 代码块 | 灰色背景代码框 |
| `- bullet` 列表 | 项目符号，蓝色圆点 |
| `\| table \|` 表格 | 深蓝表头 + 交替行背景色 |
| `[text](#anchor)` 链接 | 保留文本，生成 Word 书签 |

## 排版规格

- **页面方向**：横版 A4（29.7cm × 21.0cm）
- **正文字体**：宋体 10.5pt
- **标题字体**：黑体（各级递减 16pt → 9pt）
- **页眉**：文件名（居中）
- **页脚**：文件名 + 页码
- **行距**：1.15倍

## 文件结构

```
md2docx_diy/
├── SKILL.md          # 技能说明文件
└── md2docx_diy.py    # 主转换脚本
```

## 快速开始

```bash
# 1. 安装依赖
pip install python-docx

# 2. 转换单个文件
python md2docx_diy.py input.md -o output.docx

# 3. 批量转换
python md2docx_diy.py --all
```

## 注意事项

- 仅支持 Windows（依赖 SimHei/SimSun/楷体系统字体）
- 输入 MD 文件必须为 UTF-8 编码

---
Author: ZhangGZ-medical
