import json
import sys
from datetime import datetime, timezone, timedelta

KST = timezone(timedelta(hours=9))


def render(path: str) -> None:
    print(f"\n{'='*70}\nSESSION FILE: {path}\n{'='*70}\n")
    title = None
    with open(path, "r", encoding="utf-8") as f:
        for line_num, raw in enumerate(f, 1):
            try:
                rec = json.loads(raw)
            except json.JSONDecodeError:
                continue

            t = rec.get("type")
            if t == "ai-title":
                title = rec.get("aiTitle")
                continue
            if t in ("queue-operation", "file-history-snapshot", "last-prompt", "attachment"):
                continue

            ts_raw = rec.get("timestamp")
            ts_str = ""
            if ts_raw:
                try:
                    dt = datetime.fromisoformat(ts_raw.replace("Z", "+00:00")).astimezone(KST)
                    ts_str = dt.strftime("%H:%M:%S")
                except Exception:
                    pass

            msg = rec.get("message") or {}
            role = msg.get("role")
            content = msg.get("content", [])

            if t == "user":
                text_parts = []
                for c in content if isinstance(content, list) else []:
                    if c.get("type") == "text":
                        text_parts.append(c.get("text", ""))
                    elif c.get("type") == "tool_result":
                        continue
                if text_parts:
                    print(f"[{ts_str}] USER: {' '.join(text_parts).strip()[:2000]}\n")

            elif t == "assistant" and role == "assistant":
                text_parts = []
                tool_uses = []
                for c in content if isinstance(content, list) else []:
                    if c.get("type") == "text":
                        text_parts.append(c.get("text", ""))
                    elif c.get("type") == "tool_use":
                        name = c.get("name", "?")
                        inp = c.get("input", {})
                        if name == "Bash":
                            tool_uses.append(f"Bash: {inp.get('command','')[:200]}")
                        elif name == "PowerShell":
                            tool_uses.append(f"PowerShell: {inp.get('command','')[:200]}")
                        elif name == "Read":
                            tool_uses.append(f"Read {inp.get('file_path','')}")
                        elif name == "Edit":
                            tool_uses.append(f"Edit {inp.get('file_path','')}")
                        elif name == "Write":
                            tool_uses.append(f"Write {inp.get('file_path','')}")
                        elif name == "Agent":
                            tool_uses.append(f"Agent({inp.get('subagent_type','')}): {inp.get('description','')}")
                        elif name == "Skill":
                            tool_uses.append(f"Skill: {inp.get('skill','')}")
                        elif name == "TodoWrite":
                            tool_uses.append("TodoWrite")
                        else:
                            tool_uses.append(name)
                txt = "\n".join(p for p in text_parts if p).strip()
                if txt:
                    print(f"[{ts_str}] CLAUDE: {txt[:3000]}\n")
                for tu in tool_uses:
                    print(f"        -> tool: {tu}")
                if tool_uses:
                    print()

    if title:
        print(f"(session title: {title})\n")


if __name__ == "__main__":
    for p in sys.argv[1:]:
        render(p)
