"""排版演示：无需 API Key，直接渲染一份样例文档，查看 Word 排版效果。

运行：python demo.py  ->  output/示例_排版演示.docx
"""
from __future__ import annotations

from pathlib import Path

from agent.renderer import render_markdown_to_docx

SAMPLE = """## 项目概述

本演示文档用于展示 WordAgent 的**排版能力**，包括标题层级、列表、表格、代码块与引用等元素。

## 核心要点

1. 自动识别需求并生成**结构化大纲**
2. 逐节生成正文，保持上下文一致
3. 一键输出排版精美的 `.docx` 文件

### 功能清单

- 需求解析：指令 -> 大纲
- 内容生成：大纲 -> 正文（Markdown）
- 排版输出：Markdown -> docx
- 上下文记忆：历史指令与结果持久化

### 数据示例

| 指标 | 数值 | 说明 |
| --- | --- | --- |
| 文档生成速度 | 约 1 分钟 | 视章节数量而定 |
| 排版准确率 | 高 | 标题/表格/代码块自动处理 |
| 依赖 | 3 个 | requests / python-docx / dotenv |

## 代码示例

```python
from agent import run_pipeline

path = run_pipeline("写一份季度汇报", config, memory)
print(f"已生成: {path}")
```

> WordAgent 支持中文排版，正文使用微软雅黑，一级标题自动添加分隔线，页脚自动生成页码。

## 总结

输入一句话需求，即可得到一份结构完整、排版规范的 Word 文档。
"""


def main() -> None:
    out = Path("output") / "示例_排版演示.docx"
    render_markdown_to_docx(
        SAMPLE, out, title="WordAgent 排版演示文档", style="business", toc=True
    )
    print(f"✔ 演示文档已生成：{out.resolve()}")


if __name__ == "__main__":
    main()
