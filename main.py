"""WordAgent 控制台入口。

用法：
  python main.py                          # 交互式会话
  python main.py --once "写一份产品需求文档"
  python main.py --file 报告.docx --once "把第二段改成……"     # 编辑模式
  python main.py --file 报告.docx --once "……" --preview       # 只预览修改计划
  python main.py --file 报告.docx --once "……" --overwrite     # 覆盖原文件（自动备份）
  python main.py --file 报告.docx --once "……" --no-verify     # 跳过 AI 复核
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from agent import (
    AmbiguousTargetError, Config, LLMClient, LLMError, Memory, edit_document,
    fill_template, run_pipeline,
)

BANNER = r"""
  __        __   _     _____             _             _
  \ \      / /__| |__ |  __ \           | |           | |
   \ \ /\ / / _ \ '_ \| |  | | __ _  ___| |_ __ _ _ __| |_
    \ V  V /  __/ | | | |  | |/ _` |/ _ \ __/ _` | '__| __|
     \_/\_/ \___|_| |_| |_| |_|\__,_|\___/\__\__,_|_|  \__|
     AI Word 文档 Agent  v1.4.0（生成 + 编辑）
"""

HELP_TEXT = f"""{BANNER}
【生成模式】（默认）
  输入需求即可自动生成新文档（生成前会先预览大纲，确认后才写正文，不满意可直接取消）：
  「写一份《智能办公软件》产品需求文档，商务风格，6 个章节」

【编辑模式】修改已有文档
  /file 文件路径.docx        指定要编辑的文件
  /new                       退出编辑模式，回到生成模式
  指定文件后，直接输入修改要求即可，例如：
  「把第二段改得更正式」「标题改成……」「删除最后一段」「新增一节：风险分析」

【通用命令】
  /help               显示本帮助
  /memory             查看上下文记忆
  /clear              清空上下文记忆
  /output             查看输出文件夹
  /set-output <路径>   修改输出文件夹（本次会话生效）
  /exit 或 /quit      退出程序

【参考文件】生成时可用 --refs 附带 图片/文本/csv/docx 作为写作参考：
  python main.py --once "写一份实验报告" --refs 数据.csv 截图.png
"""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="WordAgent - AI Word 文档生成/编辑 Agent（控制台版）")
    parser.add_argument("--api-key", help="DeepSeek API Key（也可用 DEEPSEEK_API_KEY 环境变量或 .env）")
    parser.add_argument("--model", help="模型名称（默认读取 DEEPSEEK_MODEL 或 deepseek-v4-flash）")
    parser.add_argument("--standard", metavar="标准", default=None,
                        help="排版标准：gb_9704（公文）/paper（论文）/experiment（实验）/meeting（纪要）/weekly（周报）/business（商务）/general（通用）")
    parser.add_argument("--output-dir", help="输出/备份文件夹（默认 output/）")
    parser.add_argument("--once", metavar="指令", help="执行单条指令后退出")
    parser.add_argument("--file", "--edit", dest="file", metavar="路径", help="编辑模式：指定要修改的 .docx 文件")
    parser.add_argument("--preview", action="store_true", help="编辑模式下只预览修改计划，不写盘")
    parser.add_argument("--overwrite", action="store_true", help="编辑后覆盖原文件（自动备份到 output/backups/）")
    parser.add_argument("--no-verify", action="store_true", help="跳过 AI 复核（默认修改后会复核是否满足要求）")
    parser.add_argument("--auto-fix", action="store_true", help="自动应用 AI 复核建议的补充操作（默认仅报告不自动执行）")
    parser.add_argument("--yes", action="store_true", help="编辑前不询问确认（配合 --once 使用）")
    parser.add_argument("--refs", nargs="+", metavar="路径", default=None,
                        help="参考文件（图片/文本/csv/docx），内容会注入生成提示")
    parser.add_argument("--fill-template", metavar="模板.docx", default=None,
                        help="模板填写模式：在给定模板 docx 中直接填写内容（保留模板表格/格式），需配合 --once 指令")
    parser.add_argument("--reset-memory", action="store_true", help="启动时清空上下文记忆文件")
    return parser


def ensure_api_key(config: Config) -> bool:
    if config.api_key:
        return True
    print("✖ 未配置 DeepSeek API Key。请任选一种方式：")
    print("  1. 创建 .env 文件（参考 .env.example），填入 DEEPSEEK_API_KEY=sk-xxx")
    print("  2. 设置环境变量 DEEPSEEK_API_KEY=sk-xxx")
    print("  3. 启动时传参：python main.py --api-key sk-xxx")
    return False


def console_confirm_plan(plan: dict) -> bool:
    print("\n========== 文档大纲预览 ==========")
    print(f"标题：{plan.get('title', '')}｜风格：{plan.get('style', '')}｜类型：{plan.get('doc_type', '')}")
    for sec in plan.get("sections", []):
        indent = "　" * (max(1, int(sec.get("level", 1))) - 1)
        print(f"   {indent}· {sec.get('heading', '')}")
    answer = input("\n按此大纲生成正文？(y/N，直接回车=取消): ").strip().lower()
    return answer in ("y", "yes", "是")


def console_confirm(plan: dict) -> bool:
    print("\n========== 修改计划预览 ==========")
    print(f"说明：{plan.get('summary', '')}")
    for op in plan["operations"]:
        extra = op.get("new_text", op.get("style", ""))
        print(f"  [{op['op']}] 目标「{op['target'][:40]}」" + (f"\n        → {str(extra)[:60]}" if extra else ""))
    answer = input("\n确认执行以上修改？(y/N): ").strip().lower()
    return answer in ("y", "yes", "是")


def console_resolve_ambiguous(op: dict, candidates: list[str]) -> str | None:
    print(f"\n⚠ 目标「{str(op.get('target', ''))[:40]}」匹配到多个段落，请选择：")
    shown = candidates[:5]
    for i, candidate in enumerate(shown, 1):
        print(f"   {i}. {candidate[:60]}")
    answer = input("输入序号（1-5），直接回车取消: ").strip()
    if answer.isdigit() and 1 <= int(answer) <= len(shown):
        return shown[int(answer) - 1]
    return None


def console_confirm_warnings(issues: list[str]) -> bool:
    print("\n⚠ 结构校验发现以下风险：")
    for issue in issues:
        print(f"  · {issue}")
    answer = input("\n仍要继续保存吗？(y/N): ").strip().lower()
    return answer in ("y", "yes", "是")


def run_edit(args, config: Config, memory: Memory, instruction: str, file_path: Path) -> int:
    try:
        llm = LLMClient(config)
        edit_document(
            file_path,
            instruction,
            config,
            memory,
            llm=llm,
            output_dir=config.output_dir,
            confirm=None if args.yes else console_confirm,
            save_as_new=not args.overwrite,
            dry_run=args.preview,
            verify=not args.no_verify,
            auto_fix=args.auto_fix,
            resolve_ambiguous=console_resolve_ambiguous if sys.stdin.isatty() else None,
            on_warnings=None if args.yes else console_confirm_warnings,
            reference_files=args.refs,
        )
        if hasattr(llm, "usage_text"):
            print(f"   ⚙ {llm.usage_text()}")
    except (LLMError, Exception) as exc:  # noqa: BLE001
        print(f"✖ 编辑失败：{exc}")
        return 1
    return 0


def main() -> int:
    args = build_parser().parse_args()
    config = Config.from_env()
    if args.api_key:
        config.api_key = args.api_key.strip()
    if args.model:
        config.model = args.model.strip()
    standard_override = args.standard.strip() if args.standard else None
    if args.output_dir:
        config.output_dir = Path(args.output_dir)

    if args.reset_memory:
        config.memory_file.unlink(missing_ok=True)
        print("已清空上下文记忆文件。")

    memory = Memory(config.memory_file)
    if not ensure_api_key(config):
        return 1

    edit_file: Path | None = Path(args.file) if args.file else None

    # 单次执行
    if args.once:
        if args.fill_template:
            try:
                llm = LLMClient(config)
                fill_template(Path(args.fill_template), args.once, config, memory, llm=llm)
            except (LLMError, Exception) as exc:  # noqa: BLE001
                print(f"✖ 模板填写失败：{exc}")
                return 1
            return 0
        if edit_file is not None:
            return run_edit(args, config, memory, args.once, edit_file)
        try:
            refs = [Path(r) for r in args.refs] if args.refs else None
            llm = LLMClient(config)
            run_pipeline(args.once, config, memory, reference_files=refs, llm=llm,
                         standard_override=standard_override)
            print(f"   ⚙ {llm.usage_text()}")
        except (LLMError, Exception) as exc:  # noqa: BLE001
            print(f"✖ 生成失败：{exc}")
            return 1
        return 0

    print(HELP_TEXT)
    while True:
        try:
            command = input("\n📝 请输入 > ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break
        if not command:
            continue
        low = command.lower()
        if low in ("/exit", "/quit", "退出"):
            print("再见！")
            break
        if low == "/help":
            print(HELP_TEXT)
            continue
        if low == "/memory":
            print(memory.history_text())
            continue
        if low == "/clear":
            memory.clear()
            print("已清空记忆。")
            continue
        if low == "/output":
            print(f"输出/备份文件夹：{config.output_dir.resolve()}")
            continue
        if low.startswith("/set-output "):
            target = command[len("/set-output "):].strip()
            config.output_dir = Path(target)
            memory.set_setting("output_dir", str(Path(target).resolve()))
            print(f"输出文件夹已改为：{config.output_dir.resolve()}")
            continue
        if low.startswith("/file "):
            target = command[len("/file "):].strip()
            if not Path(target).exists():
                print(f"✖ 文件不存在：{target}")
                continue
            edit_file = Path(target)
            print(f"✎ 已进入编辑模式，当前文件：{edit_file}（输入修改要求即可；/new 退出）")
            continue
        if low == "/new":
            edit_file = None
            print("已退出编辑模式，回到生成模式。")
            continue

        if edit_file is not None:
            if run_edit(args, config, memory, command, edit_file) != 0:
                print("可输入 /new 退出编辑模式。")
        else:
            try:
                refs = [Path(r) for r in args.refs] if args.refs else None
                llm = LLMClient(config)
                run_pipeline(command, config, memory, reference_files=refs, llm=llm,
                             confirm_plan=console_confirm_plan,
                             standard_override=standard_override)
                print(f"   ⚙ {llm.usage_text()}")
            except (LLMError, Exception) as exc:  # noqa: BLE001
                print(f"✖ 生成失败：{exc}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
