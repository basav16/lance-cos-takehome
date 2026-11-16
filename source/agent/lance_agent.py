"""Lance CoS agent: runs the Chief of Staff's daily loop on the workbook with Claude. Drafts only, never sends.

    python lance_agent.py --workbook Lance_CoS_Workbook_Basav_Sharma.xlsx --task morning
    python lance_agent.py --workbook ... --task eod
    python lance_agent.py --workbook ... --task investor
    python lance_agent.py --workbook ... --ask "Move the insurance pack to Tuesday and tell the COO"
    python lance_agent.py --sheet <google-sheet-id> --credentials sa.json --task morning   # the live Google Sheet

Every write lands as a Draft or a calendar block, and is logged to the Agent Log tab.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
import uuid
from pathlib import Path

from tools import ExcelStore, SheetsStore, Toolbox, ToolError

HERE = Path(__file__).resolve().parent
DEFAULT_MODEL = "claude-opus-5-5"
TASKS = {
    "morning": "Run the morning loop.",
    "eod": "Run the end-of-day loop.",
    "investor": "Draft this month's investor update for my review.",
}


def load_config(store) -> dict:
    """Settings from the Agent tab (key / value). Missing keys fall back to safe defaults."""
    cfg = {"model": DEFAULT_MODEL, "effort": "high", "timezone": "America/Los_Angeles", "as_of": "",
           "max_drafts_per_run": 30, "max_turns": 40, "mode": "draft_only"}
    if "Agent" in store.tabs():
        for r in store.rows("Agent"):
            k = str(r.get("key", "")).strip()
            if k:
                cfg[k] = r.get("value", "")
    if cfg.get("mode") != "draft_only":
        raise SystemExit("Agent tab: mode must be draft_only. This agent has no way to send, by design.")
    return cfg


def as_of_time(cfg: dict) -> dt.datetime:
    v = str(cfg.get("as_of") or "").strip()
    if v:
        return dt.datetime.fromisoformat(v)
    try:
        from zoneinfo import ZoneInfo

        return dt.datetime.now(ZoneInfo(cfg.get("timezone") or "America/Los_Angeles")).replace(tzinfo=None)
    except Exception:
        return dt.datetime.now()


def make_client():
    import anthropic  # imported here so the offline tests run without the SDK

    return anthropic.Anthropic(max_retries=4)


def run(store, task: str, ask: str | None = None, client=None, model: str | None = None, verbose: bool = True) -> dict:
    cfg = load_config(store)
    model = model or os.environ.get("LANCE_AGENT_MODEL") or cfg["model"] or DEFAULT_MODEL
    now = as_of_time(cfg)
    run_id = now.strftime("%m%d-%H%M-") + uuid.uuid4().hex[:4]
    box = Toolbox(store=store, run_id=run_id, task=task, as_of=now, max_drafts=int(cfg.get("max_drafts_per_run") or 30))
    system_text = (HERE / "system_prompt.md").read_text(encoding="utf-8")
    tools = Toolbox.definitions()
    first = f"Task: {task}. Now: {now.strftime('%A %Y-%m-%d %H:%M')} ({cfg.get('timezone')}). {TASKS.get(task, '')}"
    if ask:
        first += "\n\nMy request (from Basav):\n" + ask
    messages = [{"role": "user", "content": first}]
    client = client or make_client()
    usage = {"input": 0, "output": 0, "cache_read": 0, "cache_write": 0}
    box.log("run_start", detail=f"model {model}, task {task}" + (f", ask: {ask}" if ask else ""))
    final_text, stop = "", "max_turns"

    for turn in range(int(cfg.get("max_turns") or 40)):
        resp = _call(client, model=model, system_text=system_text, tools=tools, messages=messages, effort=str(cfg.get("effort") or "high"))
        u = getattr(resp, "usage", None)
        if u is not None:
            usage["input"] += getattr(u, "input_tokens", 0) or 0
            usage["output"] += getattr(u, "output_tokens", 0) or 0
            usage["cache_read"] += getattr(u, "cache_read_input_tokens", 0) or 0
            usage["cache_write"] += getattr(u, "cache_creation_input_tokens", 0) or 0
        # Send the assistant turn back exactly as it came (thinking blocks included): the history is append-only.
        messages.append({"role": "assistant", "content": resp.content})

        if resp.stop_reason == "refusal":
            stop = "refusal"
            box.log("stopped", detail="The model declined this request. Nothing else was written.")
            break
        if resp.stop_reason == "pause_turn":
            continue
        if resp.stop_reason == "max_tokens":
            stop = "max_tokens"
            box.log("stopped", detail="Hit the output limit mid-turn; rerun the task.")
            break
        calls = [b for b in resp.content if getattr(b, "type", None) == "tool_use"]
        if resp.stop_reason == "end_turn" or not calls:
            final_text = "\n".join(getattr(b, "text", "") for b in resp.content if getattr(b, "type", None) == "text").strip()
            stop = "done"
            break
        results = []
        for c in calls:  # all results go back in ONE user message
            try:
                out = box.run(c.name, c.input if isinstance(c.input, dict) else json.loads(c.input))
                results.append({"type": "tool_result", "tool_use_id": c.id, "content": json.dumps(out, default=str)})
                if verbose:
                    print(f"  · {c.name}", file=sys.stderr)
            except (ToolError, TypeError, ValueError) as e:
                results.append({"type": "tool_result", "tool_use_id": c.id, "content": f"Error: {e}", "is_error": True})
                box.log("refused", detail=f"{c.name}: {e}")
                if verbose:
                    print(f"  ✗ {c.name}: {e}", file=sys.stderr)
        messages.append({"role": "user", "content": results})
        store.save()  # keep progress if a later turn fails

    box.log("run_end", detail=f"{stop}; {box.drafts} drafts waiting for approval; tokens in {usage['input']} (cached {usage['cache_read']}), out {usage['output']}")
    if final_text:
        box.log("summary", detail=final_text)
    store.save()
    return {"run_id": run_id, "stop": stop, "summary": final_text, "drafts": box.drafts, "writes": box.writes, "usage": usage, "turns": turn + 1}


def _call(client, *, model, system_text, tools, messages, effort):
    """One Messages API call. Tools and system prompt are fixed per run, so the prefix caches across turns."""
    params = dict(
        model=model,
        max_tokens=16000,
        system=[{"type": "text", "text": system_text}],
        tools=tools,
        messages=messages,
        thinking={"type": "adaptive"},
        output_config={"effort": effort},
        cache_control={"type": "ephemeral"},  # auto-caches the growing conversation prefix
    )
    if model.startswith(("claude-opus-5", "claude-fable-5", "claude-sonnet-5-5")):
        # If the model declines on a safety classifier, the API retries the same request on a fallback model.
        params.update(betas=["server-side-fallback-2026-07-01"], fallbacks="default")
    # The client retries 408/409/429/5xx and dropped connections itself (max_retries=4); anything else should stop the run.
    return client.beta.messages.create(**params)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Lance CoS agent: drafts only, never sends.")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--workbook", help="Path to Lance_CoS_Workbook_Basav_Sharma.xlsx")
    src.add_argument("--sheet", help="Google Sheet id (live mode)")
    ap.add_argument("--credentials", help="Service account JSON for --sheet")
    ap.add_argument("--out", help="Write to a copy instead of the workbook itself (a dry run)")
    ap.add_argument("--task", choices=["morning", "eod", "investor", "ask"], default="morning")
    ap.add_argument("--ask", help="Free-text request (sets --task ask)")
    ap.add_argument("--model", help=f"Override the model (default from the Agent tab, else {DEFAULT_MODEL})")
    a = ap.parse_args(argv)
    task = "ask" if a.ask else a.task
    if task == "ask" and not a.ask:
        ap.error("--task ask needs --ask")
    import anthropic

    store = SheetsStore(a.sheet, a.credentials) if a.sheet else ExcelStore(a.workbook, out=a.out)
    try:
        res = run(store, task, ask=a.ask, model=a.model)
    except anthropic.AuthenticationError:
        print("No valid Claude API key. Set ANTHROPIC_API_KEY (see .env.example) or run `ant auth login`.", file=sys.stderr)
        return 2
    except anthropic.NotFoundError:
        print("Model not found. Check the model on the Agent tab.", file=sys.stderr)
        return 2
    except anthropic.RateLimitError:
        print("Rate limited after retries. Try again in a minute; the work so far is saved.", file=sys.stderr)
        return 3
    except anthropic.APIStatusError as e:
        print(f"Claude API error {e.status_code}: {e.message}. The work so far is saved.", file=sys.stderr)
        return 3
    except anthropic.APIConnectionError:
        print("Could not reach the Claude API. Check the network; the work so far is saved.", file=sys.stderr)
        return 3
    print(res["summary"] or f"(stopped: {res['stop']})")
    print(f"\n{res['drafts']} drafts waiting for your approval · run {res['run_id']} · "
          f"tokens in {res['usage']['input']} (cached {res['usage']['cache_read']}), out {res['usage']['output']}", file=sys.stderr)
    return 0 if res["stop"] == "done" else 1


if __name__ == "__main__":
    sys.exit(main())
