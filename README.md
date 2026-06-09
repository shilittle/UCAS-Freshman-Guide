# UCAS Freshman Guide

面向中国科学院大学 2026 级本科新生的新生指南协作仓库。

本仓库把几份早期资料整理为可维护的 LaTeX 项目：其中《来到国科大的第一课》和《如何选课》逆向自试管撰写的旧版 PDF，并在章节中保留原作者署名；企业号流程、军训作息和 2025 年度本科生学业导师名单作为参考资料归档和索引。

## 目录结构

| 路径 | 用途 |
| --- | --- |
| `source/main.tex` | 手册主入口 |
| `source/chapters/` | 正文章节 |
| `source/appendices/` | 附录 |
| `references/originals/` | 原始 PDF/DOCX/XLSX 资料 |
| `data/mentors-2025/` | 由导师名单 XLSX 导出的 CSV |
| `scripts/bootstrap_sources.py` | 从原始资料生成初版 LaTeX/CSV |
| `scripts/build.ps1` | 本地 XeLaTeX 构建脚本 |

## 本地构建

需要 TeX Live 或 MacTeX，并确保 `latexmk` 和 `xelatex` 在 PATH 中。

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build.ps1
```

生成的 PDF 位于 `build/main.pdf`。

也可以手动执行：

```powershell
latexmk -xelatex -interaction=nonstopmode -halt-on-error "-outdir=build" source/main.tex
```

## 协作约定

优先编辑 `source/**/*.tex`。不要直接修改 PDF、DOCX、XLSX 原件；如果原始资料有新版，请放入 `references/originals/` 并在 README 或对应章节注明来源。

首次逆向稿主要用于建立结构，仍需要人工校对断句、错字、标题层级和 2026 级政策变化。企业号章节含大量截图和二维码，当前只保留文字抽取结果，后续应按 2026 级实际流程重做截图或改为文字说明。

## 版权说明

《来到国科大的第一课》和《如何选课》原作者为试管，中国科学院大学本科部 23 级学生。仓库暂未声明统一开源许可证；正式发布或二次分发前，请维护者确认原作者授权和资料适用范围。
