"""Offline tests: no API key, no network. A scripted fake Claude drives the real harness and the real tools
against a copy of the workbook.  Run:  python test_agent.py  [path/to/workbook.xlsx]   (or: pytest test_agent.py)"""
from __future__ import annotations

import datetime as dt
import os
import shutil
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace as NS

from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import lance_agent  # noqa: E402
from tools import ExcelStore, SheetsStore, Toolbox, ToolError  # noqa: E402

WB = Path(os.environ.get("LANCE_WORKBOOK") or (sys.argv[1] if len(sys.argv) > 1 and sys.argv[1].endswith(".xlsx") else HERE.parent / "Lance_CoS_Workbook_Basav_Sharma.xlsx"))
TMP = Path(tempfile.mkdtemp(prefix="lance-agent-"))


def fixture(name="wb.xlsx", new_item=True):
    """A copy of the workbook with one new, unsorted inbox item (#8), as Monday 9 AM would see it."""
    p = TMP / name
    shutil.copyfile(WB, p)
    if new_item:
        wb = load_workbook(p)
        ws = wb["Inbox"]
        r = ws.max_row + 1
        for j, v in enumerate([8, "Hotel group wants our SOC 2 letter before signing", "", "", "", "", None], start=1):
            ws.cell(row=r, column=j, value=v)
        wb.save(p)
    return p


def tool_use(i, name, **inp):
    return NS(type="tool_use", id=f"toolu_{i}", name=name, input=inp)


def msg(stop, *blocks):
    return NS(content=list(blocks), stop_reason=stop,
              usage=NS(input_tokens=1000, output_tokens=200, cache_read_input_tokens=800, cache_creation_input_tokens=0))


class FakeClient:
    """Stands in for anthropic.Anthropic(). Returns scripted turns and records every request."""

    def __init__(self, turns):
        self.turns, self.calls = list(turns), []
        self.beta = NS(messages=NS(create=self._create))

    def _create(self, **kw):
        # copy the list so later appends don't rewrite what this call saw
        self.calls.append({**kw, "messages": list(kw["messages"])})
        return self.turns.pop(0)


THINK = NS(type="thinking", thinking="", signature="sig-abc")
MORNING = [
    msg("tool_use", THINK, tool_use(1, "get_briefing")),
    msg("tool_use",
        tool_use(2, "triage_item", id=8, priority="P1", rule="outside_party_date_or_cost",
                 plan="Send the SOC 2 Type I letter + security page by Tue noon", due="Tue noon"),
        tool_use(3, "add_block", date="2026-11-17", start="11:30", end="11:45", workstream="Inbox fire", inbox_ids="8",
                 title="SOC 2 letter to hotel group", steps=["Pull the letter from the compliance folder", "Email it with the security page"],
                 thinking="Outside party with a deal at stake, so it goes before noon.")),
    msg("tool_use",
        tool_use(4, "draft_message", block_id="AG01", channel="Email", to="Hotel group IT lead",
                 subject="Lance SOC 2 letter", text="Hi [name], attached is our SOC 2 Type I letter and our security page. Happy to set up a call with our CTO. Basav"),
        tool_use(5, "draft_message", block_id="B99", channel="Slack DM", to="CEO", subject="", text="FYI"),
        tool_use(6, "add_block", date="2026-11-16", start="9:00", end="9:30", workstream="Team", inbox_ids="", title="Clash", steps=[], thinking="x")),
    msg("end_turn", NS(type="text", text="Sorted #8 (P1). Added 1 block Tue 11:30. 1 draft waits for you.")),
]


# ------------------------------------------------------------------------------------------- end to end
def test_morning_run_end_to_end():
    p = fixture()
    out = TMP / "after_morning.xlsx"
    fake = FakeClient(MORNING)
    res = lance_agent.run(ExcelStore(p, out=out), "morning", client=fake, verbose=False)
    assert res["stop"] == "done" and res["drafts"] == 1 and res["turns"] == 4, res
    wb = load_workbook(out)
    inbox = {r[0]: r for r in wb["Inbox"].iter_rows(min_row=2, values_only=True) if r[0] is not None}
    assert inbox[8][2] == "P1" and str(inbox[8][4]).startswith("Plan: Send the SOC 2") and inbox[8][5] == "Open" and inbox[8][3] == "Tue noon"
    blocks = [r for r in wb["Blocks"].iter_rows(min_row=2, values_only=True) if r[0]]
    ag = [b for b in blocks if b[0] == "AG01"]
    assert len(ag) == 1 and ag[0][2:4] == ("11:30", "11:45") and ag[0][1].date() == dt.date(2026, 11, 17)
    assert not any(b[0] == "AG02" for b in blocks), "the overlapping block was refused"
    msgs = [r for r in wb["Messages"].iter_rows(min_row=2, values_only=True) if r[0]]
    assert len(msgs) == 25 and msgs[-1][0] == "AG01" and msgs[-1][6] == "Draft" and msgs[-1][4] == "Lance SOC 2 letter"
    assert all(m[6] == "Draft" for m in msgs), "nothing is ever marked Ready or Sent by the agent"
    log = [r[3] for r in wb["Agent Log"].iter_rows(min_row=2, values_only=True) if r[3]]
    for a in ("run_start", "triage_item", "add_block", "draft_message", "refused", "run_end", "summary"):
        assert a in log, (a, log)
    assert log.count("refused") == 2
    # the original workbook is untouched (dry run wrote to --out)
    assert load_workbook(p)["Messages"].max_row == load_workbook(WB)["Messages"].max_row


def test_request_shape():
    fake = FakeClient(MORNING)
    lance_agent.run(ExcelStore(fixture("shape.xlsx"), out=TMP / "shape_out.xlsx"), "morning", client=fake, verbose=False)
    kw = fake.calls[0]
    assert kw["model"] == "claude-opus-5-5" and kw["max_tokens"] == 16000
    assert kw["thinking"] == {"type": "adaptive"} and kw["output_config"] == {"effort": "high"}
    assert kw["cache_control"] == {"type": "ephemeral"}
    assert kw["fallbacks"] == "default" and "server-side-fallback-2026-07-01" in kw["betas"]
    assert not any(k in kw for k in ("temperature", "top_p", "top_k", "tool_choice")), "sampling params and forced tool use are rejected on this model"
    for t in kw["tools"]:
        assert t["strict"] is True and t["input_schema"]["additionalProperties"] is False, t["name"]
    assert "Now:" not in kw["system"][0]["text"], "the system prompt has no timestamp, so it caches"
    assert all(c["system"] == kw["system"] and c["tools"] == kw["tools"] for c in fake.calls), "same prefix every turn"


def test_history_is_append_only_and_results_batched():
    fake = FakeClient(MORNING)
    lance_agent.run(ExcelStore(fixture("hist.xlsx"), out=TMP / "hist_out.xlsx"), "morning", client=fake, verbose=False)
    for a, b in zip(fake.calls, fake.calls[1:]):
        assert b["messages"][: len(a["messages"])] == a["messages"], "earlier turns are never edited"
    second = fake.calls[1]["messages"]
    assert second[1]["content"][0] is THINK, "the thinking block goes back unchanged"
    third = fake.calls[2]["messages"]
    assert len(third[-1]["content"]) == 2 and all(r["type"] == "tool_result" for r in third[-1]["content"]), "parallel calls answered in one message"
    fourth = fake.calls[3]["messages"][-1]["content"]
    errs = [r for r in fourth if r.get("is_error")]
    assert len(errs) == 2 and "no block 'B99'" in errs[0]["content"] and "overlaps B01" in errs[1]["content"]


def test_refusal_and_max_tokens_stop_cleanly():
    for stop in ("refusal", "max_tokens"):
        fake = FakeClient([msg(stop, NS(type="text", text=""))])
        res = lance_agent.run(ExcelStore(fixture(f"{stop}.xlsx"), out=TMP / f"{stop}_out.xlsx"), "morning", client=fake, verbose=False)
        assert res["stop"] == stop and res["drafts"] == 0


def test_real_sdk_request_round_trip():
    """The real Anthropic SDK builds the requests (no network: a mock transport answers). Checks the wire format,
    including passing the SDK's own response blocks (thinking + tool_use) back on the next turn."""
    try:
        import anthropic
        try:
            import httpx2 as httpx  # anthropic SDK 1.x+
        except ImportError:
            import httpx
    except ImportError:
        print("    (skipped: anthropic SDK not installed)")
        return
    import json as _json

    bodies = []
    replies = [
        {"content": [{"type": "thinking", "thinking": "", "signature": "sig1"},
                     {"type": "tool_use", "id": "toolu_a", "name": "triage_item",
                      "input": {"id": 8, "priority": "P2", "rule": "batch_and_nudge", "plan": "Nudge Wed"}}], "stop_reason": "tool_use"},
        {"content": [{"type": "text", "text": "Sorted #8."}], "stop_reason": "end_turn"},
    ]

    def handler(request):
        bodies.append((dict(request.headers), _json.loads(request.content)))
        r = replies[len(bodies) - 1]
        return httpx.Response(200, json={"id": f"msg_{len(bodies)}", "type": "message", "role": "assistant", "model": "claude-opus-5-5",
                                         "content": r["content"], "stop_reason": r["stop_reason"], "stop_sequence": None,
                                         "usage": {"input_tokens": 10, "output_tokens": 5, "cache_read_input_tokens": 0, "cache_creation_input_tokens": 0}})

    client = anthropic.Anthropic(api_key="test", http_client=httpx.Client(transport=httpx.MockTransport(handler)))
    res = lance_agent.run(ExcelStore(fixture("sdk.xlsx"), out=TMP / "sdk_out.xlsx"), "morning", client=client, verbose=False)
    assert res["stop"] == "done" and res["summary"] == "Sorted #8."
    h0, b0 = bodies[0]
    assert "server-side-fallback-2026-07-01" in h0.get("anthropic-beta", "") and b0["fallbacks"] == "default"
    assert b0["thinking"] == {"type": "adaptive"} and b0["cache_control"] == {"type": "ephemeral"} and b0["tools"][2]["strict"] is True
    sent_back = bodies[1][1]["messages"][1]["content"]
    assert sent_back[0] == {"type": "thinking", "thinking": "", "signature": "sig1"} and sent_back[1]["type"] == "tool_use"
    assert bodies[1][1]["messages"][2]["content"][0]["type"] == "tool_result"


# ------------------------------------------------------------------------------------------- guardrails
def box(name="g.xlsx", **kw):
    st = ExcelStore(fixture(name), out=TMP / ("o_" + name))
    return Toolbox(store=st, run_id="t", task="test", as_of=dt.datetime(2026, 11, 16, 9, 0), **kw)


def raises(fn, text):
    try:
        fn()
    except ToolError as e:
        assert text in str(e), (text, str(e))
        return
    raise AssertionError("expected ToolError: " + text)


def test_no_send_tool_exists():
    names = [t["name"] for t in Toolbox.definitions()]
    assert names == ["get_briefing", "read_tab", "triage_item", "add_block", "draft_message", "onboarding_status", "investor_metrics", "log_note"]
    assert not any("send" in n or "delete" in n for n in names)


def test_draft_cap():
    b = box("cap.xlsx", max_drafts=2)
    for _ in range(2):
        b.t_draft_message("B02", "Slack DM", "CEO", "", "hello")
    raises(lambda: b.t_draft_message("B02", "Slack DM", "CEO", "", "third"), "draft limit reached")


def test_secret_and_email_checks():
    b = box("sec.xlsx")
    raises(lambda: b.t_draft_message("B02", "Slack DM", "Eng admin", "", "token xoxb-1234567890-abcdef"), "key, token or card")
    raises(lambda: b.t_draft_message("B02", "Email", "Vendor", "", "Paid today"), "emails need a subject")
    raises(lambda: b.t_draft_message("B02", "Fax", "Vendor", "", "x"), "channel must be")


def test_calendar_checks():
    b = box("cal.xlsx")
    raises(lambda: b.t_add_block("2026-11-21", "10:00", "11:00", "Team", "", "Sat", [], ""), "weekends")
    raises(lambda: b.t_add_block("2026-11-17", "11:45", "11:30", "Team", "", "x", [], ""), "end must be after start")
    raises(lambda: b.t_add_block("2026-11-17", "6:00", "6:30", "Team", "", "x", [], ""), "between 7:00 and 20:00")
    raises(lambda: b.t_add_block("2026-11-17", "9:5", "10:00", "Team", "", "x", [], ""), "must look like")
    assert b.t_add_block("2026-11-17", "11:30", "12:00", "Team", "", "Free slot", ["a"], "b")["block_id"] == "AG01"
    assert b.t_add_block("2026-11-18", "8:00", "8:30", "Team", "", "Early", [], "")["block_id"] == "AG02"


def test_triage_checks():
    b = box("tri.xlsx")
    raises(lambda: b.t_triage_item(3, "P0", "outside_party_date_or_cost", "Pay it"), "already Done")
    raises(lambda: b.t_triage_item(99, "P0", "outside_party_date_or_cost", "x"), "no inbox item")
    raises(lambda: b.t_triage_item(8, "P0", "vibes", "x"), "rule must be")
    assert b.t_triage_item(8, "P2", "batch_and_nudge", "Nudge Wed")["ok"]


def test_investor_math_matches_the_update():
    m = box("inv.xlsx").t_investor_metrics()
    assert m["arr_oct"] == 64 * 2734 * 12 == 2099712 and m["arr_jul"] == 58 * 2299 * 12 == 1600104 and m["arr_growth_vs_jul_pct"] == 31.2
    assert m["hotels_growth_vs_jul_pct"] == 10.3 and m["arr_of_waiting"] == 22 * 2734 * 12 == 721776
    assert m["jan_goal_hotels"] == 86 and m["jan_goal_arr"] == 86 * 2810 * 12 == 2899920
    assert m["runway_months_now"] == 16.3 and m["runway_months_after_hiring"] == 10.0


def test_onboarding_status():
    b = box("onb.xlsx")
    s = b.t_onboarding_status()
    ready = {h["hire_id"]: h["ready_for_day1"] for h in s["hires"]}
    assert ready["A"] and ready["B"] and not ready["C"] and s["overdue_by_owner"] == {}
    a = next(h for h in s["hires"] if h["hire_id"] == "A")
    assert a["working_days_to_start"] == 1 and a["done_pct"] == 71
    b.as_of = dt.datetime(2026, 11, 25, 9, 0)
    late = b.t_onboarding_status()
    assert sum(late["overdue_by_owner"].values()) > 0, "a week later, open tasks show as overdue by owner"


def test_briefing_finds_the_new_item():
    br = box("brief.xlsx").t_get_briefing()
    assert [r["id"] for r in br["inbox_unsorted"]] == [8]
    assert len(br["blocks_today"]) == 12 and [h["hire_id"] for h in br["hires_starting_soon"]] == ["A", "B"]
    assert br["investor"]["arr_oct"] == 2099712


def test_config_refuses_send_mode():
    p = fixture("mode.xlsx")
    wb = load_workbook(p)
    ws = wb["Agent"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(row=r, column=1).value == "mode":
            ws.cell(row=r, column=2, value="send")
    wb.save(p)
    try:
        lance_agent.load_config(ExcelStore(p, out=TMP / "mode_out.xlsx"))
    except SystemExit as e:
        assert "draft_only" in str(e)
        return
    raise AssertionError("send mode must be refused")


def test_skill_cli_uses_the_same_rules():
    import io
    import json as _json
    from contextlib import redirect_stdout

    import tools_cli

    p = fixture("cli.xlsx")
    buf = io.StringIO()
    with redirect_stdout(buf):
        assert tools_cli.main([str(p), "add_block", _json.dumps({"date": "2026-11-17", "start": "11:30", "end": "11:45", "workstream": "Inbox fire",
                                                                  "inbox_ids": "8", "title": "From the skill", "steps": ["x"], "thinking": "y"})]) == 0
        assert tools_cli.main([str(p), "draft_message", _json.dumps({"block_id": "AG01", "channel": "Slack DM", "to": "CEO", "subject": "", "text": "Hi"})]) == 0
        assert tools_cli.main([str(p), "add_block", _json.dumps({"date": "2026-11-16", "start": "9:00", "end": "9:30", "workstream": "Team",
                                                                  "inbox_ids": "", "title": "Clash", "steps": [], "thinking": ""})]) == 1
    assert "overlaps B01" in buf.getvalue()
    msgs = [r for r in load_workbook(p)["Messages"].iter_rows(min_row=2, values_only=True) if r[0]]
    assert msgs[-1][0] == "AG01" and msgs[-1][6] == "Draft"


# ------------------------------------------------------------------------------------------- Google Sheets backend
class FakeWS:
    def __init__(self, title, grid):
        self.title, self.grid = title, [list(r) for r in grid]

    def get_all_values(self, **kw):
        return [list(r) for r in self.grid]

    def row_values(self, i):
        return list(self.grid[i - 1]) if len(self.grid) >= i else []

    def append_row(self, values, value_input_option=None):
        self.grid.append(list(values))

    def update_cell(self, r, c, v):
        while len(self.grid) < r:
            self.grid.append([])
        row = self.grid[r - 1]
        while len(row) < c:
            row.append("")
        row[c - 1] = v


class FakeBook:
    def __init__(self, sheets):
        self.sheets = {k: FakeWS(k, v) for k, v in sheets.items()}

    def worksheet(self, t):
        return self.sheets[t]

    def worksheets(self):
        return list(self.sheets.values())


def test_google_sheets_backend():
    src = load_workbook(WB, data_only=False)
    sheets = {}
    for ws in src.worksheets:
        g = []
        for row in ws.iter_rows(values_only=True):
            g.append(["" if v is None else (v.strftime("%Y-%m-%d") if isinstance(v, (dt.date, dt.datetime)) else (f"${v:,}" if ws.title == "Investor" and isinstance(v, int) and v > 1000 else v)) for v in row])
        sheets[ws.title] = g
    client = NS(open_by_key=lambda k: FakeBook(sheets))
    st = SheetsStore("sheet-id", client=client)
    b = Toolbox(store=st, run_id="g", task="test", as_of=dt.datetime(2026, 11, 16, 9, 0))
    assert b.t_investor_metrics()["arr_oct"] == 2099712, "reads '$2,734'-style cells from Sheets"
    blk = b.t_add_block("2026-11-17", "11:30", "11:45", "Inbox fire", "", "From Sheets", ["x"], "y")["block_id"]
    b.t_draft_message(blk, "Slack DM", "CEO", "", "Live-sheet draft")
    last = st.rows("Messages")[-1]
    assert last["block_id"] == "AG01" and last["status"] == "Draft"
    assert st.rows("Agent Log")[-1]["action"] == "draft_message"


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items(), key=lambda kv: kv[1].__code__.co_firstlineno if callable(kv[1]) and hasattr(kv[1], "__code__") else 0) if n.startswith("test_") and callable(f)]
    for n, f in tests:
        f()
        print("  ✓", n[5:].replace("_", " "))
    print(f"\n{len(tests)} agent tests passed (offline, no API key)")
    if os.environ.get("KEEP_OUTPUT"):
        shutil.copyfile(TMP / "after_morning.xlsx", os.environ["KEEP_OUTPUT"])
