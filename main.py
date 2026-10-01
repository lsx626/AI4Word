"""AI4Word 主程序：流式生成 + 打字机写入 + 可视化 AI 排版。

阶段一：SSE 流式接收 AI 生成的 Markdown，边收边按块解析、应用格式并以
        打字机动画写入活动的 Word 文档（真正的流式输出）。按 ESC 可中断，
        选择回滚本次生成、或保留已写内容并追加补充。
阶段二：AI 智能排版模式。AI 生成的 pywin32 代码被逐语句可视化执行：
        终端实时显示每条语句与执行结果，Word 里受影响区域滚动进视野并闪烁。
        编辑已写入的内容请优先使用注入的块编辑原语（replace_block 等），
        它们基于每块登记的动态 Range，定位精确、不会破坏文档其余部分。
        块模型可随文档保存（save_blocks），重开同一文档自动恢复索引。

CLI 与 GUI（app 包，PySide6 悬浮窗）共用 app.agent 里的代码生成提示词、
沙箱执行环境与 Word 连接逻辑，两端行为保持一致。
"""
import os

from dotenv import load_dotenv

from ai_client import ai_stream
from app.agent import (WRITER_SYSTEM, build_exec_globals, fix_code, gen_code,
                       get_word)
from doc_model import DocModel
from format_runner import run_code
from session import Session
from streaming_writer import StreamingWriter

try:
    import msvcrt
    _HAS_MSVCRT = True
except ImportError:
    msvcrt = None
    _HAS_MSVCRT = False


def _esc_pressed():
    """Windows 控制台下检测 ESC 键（流式生成期间按 ESC 中断）。"""
    if not _HAS_MSVCRT:
        return False
    try:
        if msvcrt.kbhit():
            return msvcrt.getch() == b"\x1b"
    except Exception:
        pass
    return False


def main():
    dotenv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(dotenv_path=dotenv_path)
    api_key = os.getenv("ATRIA_API_KEY")
    if not api_key:
        print("错误：未能加载有效的 ATRIA_API_KEY（请在 .env 中设置）。")
        return

    app = get_word()
    if not app:
        print("错误：无法启动或连接到 Word。")
        return
    doc = app.Documents.Add() if app.Documents.Count == 0 else app.ActiveDocument
    sel = app.Selection
    print("\nWord 已准备就绪。")

    writer = StreamingWriter(app, doc, sel, model=None, char_delay=0.01)
    model = DocModel(app, doc, sel, writer)
    writer.model = model
    session = Session()

    restored = model.load_blocks()  # 重开同一文档时恢复上次的块模型
    if restored:
        print(f"已从文档恢复 {restored} 个块（上次会话的块索引继续可用）。")
    else:
        imported = model.import_document()  # 首次接入没有存档的非空文档：读取现有内容
        if imported:
            print(f"已读取现有文档 {imported} 个块（现有内容的块索引可用）。")

    # --- 阶段一：流式生成 + 动画写入 ---
    prompt = input("请输入您的写作需求：")
    if prompt:
        system_prompt = WRITER_SYSTEM
        print("\n正在流式生成并以动画效果写入 Word（按 ESC 可中断）...")
        writer.set_speed("auto")
        snap_before = model.snapshot()
        received = 0
        interrupted = False
        try:
            for piece in ai_stream(prompt, api_key, system_prompt):
                if _esc_pressed():
                    interrupted = True
                    break
                if not piece:
                    continue
                received += len(piece)
                print(piece, end="", flush=True)  # 终端同步显示流式原文
                writer.feed(piece)
        except Exception as e:
            print(f"\n流式生成出错：{e}")
        writer.flush()
        print()  # 换行结束终端原文回显
        if received:
            print("写入完成。")
        elif not interrupted:
            print("没有收到任何内容（AI 返回为空）。")
        if interrupted:
            print("生成已被 ESC 中断。")
            choice = input("输入 r 回滚本次生成，输入 w 保留并追加补充（回车=保留现状）：").strip().lower()
            if choice == "r":
                try:
                    model.restore_snapshot(snap_before)
                    print("已回滚到生成前的文档状态。")
                except Exception as e:
                    print(f"回滚失败：{e}")
            elif choice == "w":
                extra = input("补充内容描述：").strip()
                if extra:
                    writer.set_speed("auto")
                    print("\n正在流式追加补充内容...")
                    try:
                        for piece in ai_stream(extra, api_key, system_prompt):
                            if _esc_pressed():
                                break
                            print(piece, end="", flush=True)
                            writer.feed(piece)
                    except Exception as e:
                        print(f"\n追加生成出错：{e}")
                    writer.flush()
                    print()
        try:
            model.save_blocks()  # 块模型随文档存档，下次打开可恢复
        except Exception:
            pass

    # --- 阶段二：可视化 AI 智能排版 ---
    print("\n--- 进入 AI 智能排版模式（逐语句可视化执行 + 自我修复）---")
    print("编辑内容请用块索引，例如「改写第 2 块：……」「在第 0 块后插入一段……」；")
    print("格式指令例如「将所有一级标题居中并改为蓝色」；整篇风格可用 apply_preset。输入「退出」结束。")

    exec_globals = build_exec_globals(app, doc, sel, model, writer, session)

    while True:
        model.watch.poll()  # 空闲点：检测用户在停顿期间的手动改动
        notices = model.drain_events()
        for ev in notices:
            print(f"  [注意] {ev}")
        fmt = input("\n您想如何调整？> ").strip()
        if not fmt:
            continue
        if fmt.lower() in ("退出", "exit", "quit"):
            print("程序已结束。")
            break

        prompt = fmt
        if notices:
            # 用户在生成期间动过文档：把漂移警告带给代码生成模型，避免按旧结构猜索引
            prompt += "\n\n注意：用户在我生成期间手动编辑了文档，块索引可能已偏移，"
            prompt += "请先重新调用 block_map() 确认当前结构。"
        try:
            code = gen_code(prompt, api_key, session, block_map_fn=model.block_map)
        except Exception as e:
            print(f"请求 AI 生成代码失败：{e}")
            continue
        if not code:
            print("无法理解您的指令或 AI 未返回有效代码。")
            continue

        ok, err = run_code(code, exec_globals)
        if not ok and model.is_txn_active():
            try:
                model.rollback_txn()
                print("已回滚未完成的事务。")
            except Exception as e:
                print(f"回滚未完成事务失败：{e}")
        if ok:
            print("代码执行完毕，已应用。")
            session.record_turn(fmt, True)
            try:
                model.save_blocks()
            except Exception:
                pass
            continue

        try:
            corrected = fix_code(fmt, code, err, api_key, session,
                                 block_map_fn=model.block_map)
        except Exception as e:
            print(f"请求 AI 自我修复失败：{e}")
            continue
        if not corrected:
            print("AI 未能生成有效的修正代码。")
            continue
        ok2, err2 = run_code(corrected, exec_globals)
        if not ok2 and model.is_txn_active():
            try:
                model.rollback_txn()
                print("已回滚未完成的事务。")
            except Exception as e:
                print(f"回滚未完成事务失败：{e}")
        if ok2:
            print("修正代码执行完毕，已应用。")
            session.record_turn(fmt, True)
            try:
                model.save_blocks()
            except Exception:
                pass
        else:
            print(f"自我修复仍然失败（{err2}），请尝试更简单的指令。")
            session.record_turn(fmt, False)


if __name__ == "__main__":
    main()
