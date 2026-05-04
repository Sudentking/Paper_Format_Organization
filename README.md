# Word文档格式化工具

自动化修改Word文档格式的Python脚本，支持标题、正文、表格、图标题、目录、参考文献等多种元素的格式化。

## 功能特性

- ✅ 自动识别并格式化一、二、三级标题
- ✅ 格式化正文段落（字体、行距、对齐等）
- ✅ 自动应用三线表格式
- ✅ 格式化表标题和图标题
- ✅ 格式化目录条目
- ✅ 格式化参考文献
- ✅ 自动识别英文和数字并应用Times New Roman字体
- ✅ 完全可配置的格式参数

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 基本用法

```bash
python main.py input.docx
```

这将使用默认配置格式化`input.docx`，并生成`input_formatted.docx`。

### 指定输出文件

```bash
python main.py input.docx output.docx
```

### 使用自定义配置

```bash
python main.py input.docx output.docx config/custom_config.yaml
```

## 配置文件说明

配置文件位于`config/format_config.yaml`，包含以下主要部分：

### 1. 全局设置

```yaml
global:
  english_number_font: "Times New Roman"  # 所有英文和数字的字体
```

### 2. 标题格式

```yaml
headings:
  level1:  # 一级标题
    font_name_cn: "黑体"
    font_size: 14  # 四号
    bold: true
    alignment: "left"
```

支持三级标题，每级可单独配置字体、字号、加粗、对齐方式等。

### 3. 正文格式

```yaml
paragraph:
  font_name_cn: "宋体"
  font_size: 12  # 小四
  line_spacing: 1.5  # 1.5倍行距
  alignment: "justify"  # 两端对齐
  first_line_indent: 0  # 首行缩进字符数
```

### 4. 表格格式

```yaml
table:
  caption:  # 表标题（如"表1.1 表的名称"）
    font_size: 12
    bold: true
    alignment: "center"
  
  content:  # 表内文字
    font_size: 12
    alignment: "center"
    vertical_alignment: "center"
  
  style:
    border_style: "three_line"  # 三线表
    top_border_width: 1.5
    middle_border_width: 0.75
    bottom_border_width: 1.5
```

### 5. 图标题格式

```yaml
figure:
  caption:
    font_size: 12
    bold: true
    alignment: "center"
```

### 6. 参考文献格式

```yaml
reference:
  title:  # "参考文献"标题
    font_size: 10.5  # 五号
    bold: true
    alignment: "center"
  
  content:  # 参考文献条目
    font_size: 10.5
    line_spacing: 1.5
```

### 7. 目录格式

```yaml
toc:
  title:           # "目录"标题
    font_name_cn: "黑体"
    font_size: 16
    bold: true
    alignment: "center"

  level1:          # 一级目录条目
    font_name_cn: "宋体"
    font_size: 12
    bold: true
    indent: 0      # 缩进字符数

  level2:          # 二级目录条目
    font_name_cn: "宋体"
    font_size: 12
    bold: false
    indent: 2

  level3:          # 三级目录条目
    font_name_cn: "宋体"
    font_size: 12
    bold: false
    indent: 4

  line_spacing: 1.5
```

### 8. 识别规则

```yaml
recognition:
  heading_patterns:
    level1: '^[0-9]+\s+'  # 匹配"1 "、"2 "
    level2: '^[0-9]+\.[0-9]+\s+'  # 匹配"1.1 "
    level3: '^[0-9]+\.[0-9]+\.[0-9]+\s+'  # 匹配"1.1.1 "
  
  table_caption_pattern: '^表[0-9]+\.?[0-9]*\s+'
  figure_caption_pattern: '^图[0-9]+\.?[0-9]*\s+'
  reference_title_pattern: '^参考文献$'
```

## 字号对照表

| 字号 | 磅值 | 字号 | 磅值 |
|------|------|------|------|
| 初号 | 42 | 四号 | 14 |
| 小初 | 36 | 小四 | 12 |
| 一号 | 26 | 五号 | 10.5 |
| 小一 | 24 | 小五 | 9 |
| 二号 | 22 | 六号 | 7.5 |
| 小二 | 18 | 小六 | 6.5 |
| 三号 | 16 | 七号 | 5.5 |
| 小三 | 15 | 八号 | 5 |

## 自定义配置示例

如果你的文档有特殊要求，可以复制`format_config.yaml`并修改：

```bash
cp config/format_config.yaml config/my_config.yaml
# 编辑 my_config.yaml
python main.py input.docx output.docx config/my_config.yaml
```

### 示例：修改正文首行缩进2字符

```yaml
paragraph:
  first_line_indent: 2  # 首行缩进2个字符
```

### 示例：修改表格为全框线

```yaml
table:
  style:
    border_style: "full"  # 改为全框线
```

### 示例：修改标题编号模式

如果你的标题使用"第一章"、"第二章"格式：

```yaml
recognition:
  heading_patterns:
    level1: '^第[一二三四五六七八九十]+章\s+'
```

## 注意事项

1. **备份原文件**：建议先备份原始文档
2. **测试配置**：修改配置后先用测试文档验证效果
3. **复杂格式**：对于包含复杂嵌套表格或特殊格式的文档，可能需要手动调整
4. **字体安装**：确保系统已安装配置中指定的字体（黑体、宋体、Times New Roman）

## 项目结构

```
word-formatter/
├── main.py                        # 主程序入口
├── requirements.txt               # 依赖包
├── README.md                      # 说明文档
├── AGENT.md                       # 开发进度记录
├── config/
│   └── format_config.yaml         # 格式配置文件
├── formatters/                    # 格式化器模块
│   ├── __init__.py
│   ├── base_formatter.py
│   ├── heading_formatter.py
│   ├── paragraph_formatter.py
│   ├── table_formatter.py
│   ├── caption_formatter.py
│   ├── reference_formatter.py
│   └── toc_formatter.py
├── utils/                         # 工具函数
│   ├── __init__.py
│   └── docx_helper.py
└── tests/                         # 单元测试
    └── test_format.py
```

## 常见问题

### Q: 如何修改标题的段前段后间距？

A: 在配置文件中修改对应标题的`space_before`和`space_after`参数（单位：磅）。

### Q: 如何让正文首行缩进2字符？

A: 修改`paragraph`下的`first_line_indent: 2`。

### Q: 表格没有正确应用三线表格式？

A: 检查表格是否有合并单元格，复杂表格可能需要手动调整。

### Q: 英文和数字没有变成Times New Roman？

A: 确保配置文件中`global.english_number_font`设置正确，且系统已安装该字体。

## 技术支持

如有问题或建议，请联系开发团队。
