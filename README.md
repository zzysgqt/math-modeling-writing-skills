# 数模论文写作与排版 Skills

作者：[@zzysgqt](https://github.com/zzysgqt) · 许可：[MIT](LICENSE)

面向中文数学建模竞赛论文的两项原创技能：整理符号说明与公式解释，以及把 Markdown 初稿转换为按团队规范排版的 Word 文档。

## 包含的技能

| 技能 | 用途 | 入口 |
|---|---|---|
| `math-model-symbols` | 判断符号应进入总表还是在公式后解释，维护含义、单位、字体与一致性 | [SKILL.md](skills/math-model-symbols/SKILL.md) |
| `md-to-docx-layout` | Markdown → DOCX 转换流程，以及页面、标题、字体、三线表、图片和公式编号后处理 | [SKILL.md](skills/md-to-docx-layout/SKILL.md) |

## 适用范围

- 本仓库记录作者团队的写作经验和排版约定，并非任何赛事的官方标准，也不保证获奖。
- 请先核对当届竞赛的论文模板、匿名要求和提交规则；这些要求优先于技能默认值。
- 正大杯、统计建模等赛事应分别适配，不直接套用数学建模竞赛格式。
- 期刊论文和学位论文不应默认调用本仓库技能。
- 符号表长度、局部参数处理及英文字体属于可调整约定；核心无量纲变量也可能需要进入符号表。

## 安装

将 `skills/` 下所需的完整技能目录复制到支持 `SKILL.md` 的客户端技能目录。保留 `references/`、`scripts/` 和 `agents/`，不要只复制入口文件。

也可通过 Skills CLI 选装：

```bash
npx skills add zzysgqt/math-modeling-writing-skills --skill math-model-symbols --agent codex
npx skills add zzysgqt/math-modeling-writing-skills --skill md-to-docx-layout --agent codex
```

以上为项目级安装示例；如需用户级安装，按客户端说明选择对应作用域。安装技能文件不会自动安装 Pandoc、Python 或字体。

## 使用示例

### 整理符号说明

```text
使用 $math-model-symbols 检查这份数模论文的符号说明。
区分全局核心符号与局部参数，核对首次定义、单位和全文一致性。
以我提供的赛事模板为准，不补造论文中没有使用的变量。
```

### Markdown 转 Word

```text
使用 $md-to-docx-layout 将论文.md转换为论文.docx。
先核对赛事模板，保留原始文件，转换后检查公式、图表和分页。
```

## 本地运行排版脚本

配套 [国赛模板.docx](skills/md-to-docx-layout/assets/国赛模板.docx) 位于技能的 `assets/` 目录，可复制后填写。模板包含摘要、问题重述、模型假设、符号说明、问题建模求解、模型评价、AI 工具使用声明、参考文献和附录。它是用户提供的工作模板，不代表任何年份赛事的官方最新模板。AI 使用声明及支撑材料清单是示例，须按实际情况和当届规则修改，不得直接当作事实提交。

模板发布副本仅清空文档属性中的作者和最后修改者，正文、样式及页面结构保持原样。

如希望 Pandoc 借用模板中的同名 Word 样式，可使用：

```bash
pandoc paper.md --reference-doc=skills/md-to-docx-layout/assets/国赛模板.docx -o paper_raw.docx
```

`--reference-doc` 只借用样式和文档属性，不会把模板中的章节、表格或声明正文复制进输出。需要完整章节骨架时，应复制模板后直接填写。后处理脚本会重新设置页面和样式；若当届模板与团队默认规则不同，先调整脚本或跳过不适用的后处理。

依赖：Python 3.10+、`python-docx`、Pandoc 3.x。目标字体为宋体、黑体、Times New Roman 和 Cambria Math；字体不随仓库分发。建议在隔离的 Python 环境中安装依赖。

```bash
python -m pip install -r requirements.txt
pandoc paper.md -o paper_raw.docx
python skills/md-to-docx-layout/scripts/apply_docx_layout.py paper_raw.docx paper.docx
```

在 Markdown 所在目录运行 Pandoc，或显式配置图片资源路径。使用不同的输入、输出文件名；脚本省略输出参数时会覆盖输入文件。

标题约定：`#` 为论文标题，`##` 为一级标题，`###` 为二级标题，`####` 为三级标题。公式使用 `$...$` / `$$...$$`，独立公式不手写编号。

## 局限与检查

- 字体分类依赖规则和词表，不理解所有专业符号；请按论文调整脚本中的 `LABELS`、`VAR_NAMES` 和 `UNITS`。
- 两个技能中的字体约定并非完全一致：符号说明表的局部约定应明确指定，全文后处理后仍须复查该表。
- 公式编号、复杂表格、分页、字体替代及图片位置需要打开 Word 或导出 PDF 后目检；脚本运行成功不等于版面合格。
- 文档中的历史实测经验不代表所有系统、Word 或 Pandoc 版本都已验证。

## 来源与贡献

技能内容由作者整理，符号规范参考了既往竞赛论文和团队写作实践；具体参考说明保留在 [principles.md](skills/math-model-symbols/references/principles.md)。参考论文全文及商业字体不包含在本仓库，MIT 许可仅覆盖作者有权授权的本仓库内容。

欢迎通过 Issue 提供脱敏后的最小复现案例，或通过 Pull Request 改进规则与脚本。
