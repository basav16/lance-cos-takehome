"""Tools the Lance CoS agent can use, and the rules they enforce.

The agent never sends anything. It reads the workbook (or the live Google Sheet), sorts new inbox items,
plans calendar blocks, and writes messages as Draft rows. A person marks a row Ready; the Apps Script
in the plug-in kit sends it. Every write is logged to the Agent Log tab.
"""
from __future__ import annotations

import datetime as dt
import math
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

WEEK_TABS = ["Days", "Inbox", "Blocks", "Messages", "Background", "Deadlines"]
ONB_TABS = ["Settings", "Hires", "Checklist", "Tasks", "Owners", "Access", "Week1"]
READABLE_TABS = WEEK_TABS + ONB_TABS + ["Agent", "Agent Log"]
WORKSTREAMS = ["Inbox fire", "Onboarding project", "Investor update", "CEO + closing loops", "Team"]
CHANNELS = ["Slack DM", "Slack #general", "Slack #deployment", "Slack #leads", "Email", "Email (from CEO)"]
RULES = {
    "outside_party_date_or_cost": "Outside party, hard date or real cost: done before noon",
    "needs_someone_elses_time": "Needs someone else's time: ask in the first hour",
    "under_5_minutes": "Under 5 minutes: do it now",
    "batch_and_nudge": "Low cost, not urgent: batch it, nudge, then escalate",
}
# Things that must never end up in a draft: API keys, Slack tokens, card numbers.
SECRET_PATTERNS = [r"sk-ant-[A-Za-z0-9_-]{8,}", r"xox[abprs]-[A-Za-z0-9-]{8,}", r"\b(?:\d[ -]?){13,16}\b", r"-----BEGIN [A-Z ]*PRIVATE KEY-----"]
LOG_COLS = ["time", "run_id", "task", "action", "tab", "row", "detail"]


class ToolError(Exception):
    """A tool refused the request. The message goes back to the model as an error result."""


# ----------------------------------------------------------------------------------------------- storage
def _cell_out(v):
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute) == (0, 0) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return "" if v is None else v


class ExcelStore:
    """The workbook on disk. Formulas on Summary and Investor are left alone and recalculate when opened."""

    def __init__(self, path: str | Path, out: str | Path | None = None, backup: bool = True):
        from openpyxl import load_workbook

        self.path, self.out = Path(path), Path(out) if out else Path(path)
        if backup and self.out == self.path:
            shutil.copyfile(self.path, self.path.with_suffix(".before-agent.xlsx"))
        self.wb = load_workbook(self.path)  # formulas kept as formulas
        self.dirty = False

    def tabs(self):
        return list(self.wb.sheetnames)

    def header(self, tab):
        ws = self.wb[tab]
        return [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    def rows(self, tab):
        ws, head = self.wb[tab], self.header(tab)
        out = []
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if row is None or all(v in (None, "") for v in row):
                continue
            rec = {h: _cell_out(v) for h, v in zip(head, row) if h}
            rec["_row"] = i
            out.append(rec)
        return out

    def grid(self, tab):
        return [list(r) for r in self.wb[tab].iter_rows(values_only=True)]

    def _last_row(self, tab):
        ws = self.wb[tab]
        last = 1
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if any(v not in (None, "") for v in row):
                last = i
        return last

    def append(self, tab, record: dict) -> int:
        ws, head = self.wb[tab], self.header(tab)
        r = self._last_row(tab) + 1
        for j, h in enumerate(head, start=1):
            if h in record:
                v = record[h]
                if h in ("date", "due_date", "start_date", "closed_on") and isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                    v = dt.date.fromisoformat(v)
                ws.cell(row=r, column=j, value=v)
        self.dirty = True
        return r

    def set(self, tab, row: int, column: str, value):
        head = self.header(tab)
        if column not in head:
            raise ToolError(f"{tab} has no column {column!r}")
        self.wb[tab].cell(row=row, column=head.index(column) + 1, value=value)
        self.dirty = True

    def save(self):
        if self.dirty:
            self.wb.calculation.fullCalcOnLoad = True  # Excel recomputes the Summary on open
            self.wb.save(self.out)
            self.dirty = False


class SheetsStore:
    """The live Google Sheet (same tabs). Needs `pip install gspread` and a service account the Sheet is shared with."""

    def __init__(self, sheet_id: str, credentials_file: str | None = None, client=None):
        if client is None:
            import gspread  # optional dependency

            client = gspread.service_account(filename=credentials_file) if credentials_file else gspread.service_account()
        self.book = client.open_by_key(sheet_id)
        self._cache = {}

    def _ws(self, tab):
        if tab not in self._cache:
            self._cache[tab] = self.book.worksheet(tab)
        return self._cache[tab]

    def tabs(self):
        return [ws.title for ws in self.book.worksheets()]

    def header(self, tab):
        return [str(h).strip() for h in self._ws(tab).row_values(1)]

    def rows(self, tab):
        vals = self._ws(tab).get_all_values()
        if not vals:
            return []
        head = [str(h).strip() for h in vals[0]]
        out = []
        for i, row in enumerate(vals[1:], start=2):
            if not any(str(v).strip() for v in row):
                continue
            rec = {h: (int(v) if isinstance(v, str) and v.isdigit() and h in ("id", "lane", "day", "minutes", "days_from_start") else v)
                   for h, v in zip(head, row) if h}
            rec["_row"] = i
            out.append(rec)
        return out

    def grid(self, tab):
        return self._ws(tab).get_all_values(value_render_option="UNFORMATTED_VALUE")

    def append(self, tab, record):
        head = self.header(tab)
        ws = self._ws(tab)
        ws.append_row([record.get(h, "") for h in head], value_input_option="USER_ENTERED")
        return len(ws.get_all_values())

    def set(self, tab, row, column, value):
        head = self.header(tab)
        if column not in head:
            raise ToolError(f"{tab} has no column {column!r}")
        self._ws(tab).update_cell(row, head.index(column) + 1, value)

    def save(self):
        pass  # every write is already live


# ----------------------------------------------------------------------------------------------- helpers
def _hm(s: str) -> int:
    m = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", str(s))
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
        raise ToolError(f"time {s!r} must look like 9:15 or 14:30")
    return int(m.group(1)) * 60 + int(m.group(2))


def _date(s: str) -> dt.date:
    try:
        return dt.date.fromisoformat(str(s)[:10])
    except ValueError:
        raise ToolError(f"date {s!r} must look like 2026-11-16")


def _clip(s, n=400):
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def _obj(props: dict, required: list[str]):
    return {"type": "object", "properties": props, "required": required, "additionalProperties": False}


# ----------------------------------------------------------------------------------------------- the tools
@dataclass
class Toolbox:
    store: object
    run_id: str
    task: str
    as_of: dt.datetime
    max_drafts: int = 30
    drafts: int = 0
    writes: list = field(default_factory=list)

    # ---- definitions sent to Claude (fixed order and text, so the prompt prefix caches)
    @staticmethod
    def definitions():
        S = {"type": "string"}
        return [
            {"name": "get_briefing", "strict": True,
             "description": "Start here. One call returns the morning picture: the as-of time, inbox items still open or never sorted, "
                            "today's and tomorrow's blocks, message counts by status, hires starting in the next 3 working days with "
                            "anything still open, overdue onboarding tasks by owner, and the investor numbers.",
             "input_schema": _obj({}, [])},
            {"name": "read_tab", "strict": True,
             "description": "Read rows from one tab of the workbook. Use a filter to keep it short. Long text is clipped to 400 characters.",
             "input_schema": _obj({"tab": {"type": "string", "enum": READABLE_TABS},
                                   "column": {"type": "string", "description": "Optional column to filter on, e.g. status"},
                                   "equals": {"type": "string", "description": "Value the column must equal (case-insensitive)"},
                                   "limit": {"type": "integer", "description": "Max rows, default 50"}}, ["tab"])},
            {"name": "triage_item", "strict": True,
             "description": "Sort one inbox item with my rules: set its priority, the rule that decided it, and a one-line plan "
                            "(what happens and by when). Writes Inbox.priority, Inbox.due (if blank) and Inbox.resolution. "
                            "Refuses items that are already Done.",
             "input_schema": _obj({"id": {"type": "integer"}, "priority": {"type": "string", "enum": ["P0", "P1", "P2"]},
                                   "rule": {"type": "string", "enum": list(RULES)},
                                   "plan": {"type": "string", "description": "One line, under 120 characters"},
                                   "due": {"type": "string", "description": "Optional, e.g. 'Tue 9:30' or 'today'"}},
                                  ["id", "priority", "rule", "plan"])},
            {"name": "add_block", "strict": True,
             "description": "Put a block on my calendar (Blocks tab). It must not overlap another block that day. "
                            "Steps are short lines of what I do. Thinking is one or two sentences on why this, why now. "
                            "Returns the new block_id to attach drafts to.",
             "input_schema": _obj({"date": {"type": "string", "description": "YYYY-MM-DD"}, "start": {"type": "string", "description": "H:MM, 24-hour"},
                                   "end": {"type": "string", "description": "H:MM, 24-hour"}, "workstream": {"type": "string", "enum": WORKSTREAMS},
                                   "inbox_ids": {"type": "string", "description": "Inbox ids this block serves, e.g. '3' or '2, 7'; empty if none"},
                                   "title": S, "steps": {"type": "array", "items": S}, "thinking": S},
                                  ["date", "start", "end", "workstream", "inbox_ids", "title", "steps", "thinking"])},
            {"name": "draft_message", "strict": True,
             "description": "Write one email or Slack message as a Draft row under a block (Messages tab). It is NOT sent: "
                            "I review it in the calendar app and mark it Ready. Emails need a subject. Plain, short, specific: "
                            "who, what, by when.",
             "input_schema": _obj({"block_id": S, "channel": {"type": "string", "enum": CHANNELS},
                                   "to": {"type": "string", "description": "Person or group, e.g. 'CEO', 'Eng admin', 'Vendor AR'"},
                                   "subject": {"type": "string", "description": "Required for email, empty for Slack"},
                                   "text": S}, ["block_id", "channel", "to", "subject", "text"])},
            {"name": "onboarding_status", "strict": True,
             "description": "Per hire: start date, share of tasks done, tasks still open before Day 1, overdue tasks with owner. "
                            "Plus overdue counts by owner, for nudges.",
             "input_schema": _obj({}, [])},
            {"name": "investor_metrics", "strict": True,
             "description": "The investor numbers computed from the Investor tab: hotels, ARR, growth, waiting hotels, runway. "
                            "Use these numbers only; never invent one.",
             "input_schema": _obj({}, [])},
            {"name": "log_note", "strict": True,
             "description": "Leave a note for me in the Agent Log, e.g. something you could not do or a decision only I or the CEO can make.",
             "input_schema": _obj({"text": S}, ["text"])},
        ]

    # ---- dispatch
    def run(self, name: str, args: dict):
        fn = getattr(self, "t_" + name, None)
        if fn is None:
            raise ToolError(f"unknown tool {name!r}")
        if not isinstance(args, dict):
            raise ToolError("tool input must be an object")
        return fn(**args)

    def log(self, action, tab="", row="", detail=""):
        self.store.append("Agent Log", {"time": dt.datetime.now().strftime("%Y-%m-%d %H:%M"), "run_id": self.run_id,
                                        "task": self.task, "action": action, "tab": tab, "row": row, "detail": _clip(detail, 900)})

    # ---- read tools
    def t_read_tab(self, tab, column=None, equals=None, limit=50):
        if tab not in READABLE_TABS:
            raise ToolError(f"tab must be one of {READABLE_TABS}")
        rows = self.store.rows(tab)
        if column:
            rows = [r for r in rows if str(r.get(column, "")).strip().lower() == str(equals or "").strip().lower()]
        limit = max(1, min(int(limit or 50), 200))
        return {"tab": tab, "count": len(rows), "rows": [{k: _clip(v) for k, v in r.items()} for r in rows[:limit]]}

    def t_get_briefing(self):
        today = self.as_of.date()
        tomorrow = today + dt.timedelta(days=1)
        inbox = self.store.rows("Inbox")
        blocks = self.store.rows("Blocks")
        msgs = self.store.rows("Messages")
        counts = {}
        for m in msgs:
            counts[m.get("status") or "Draft"] = counts.get(m.get("status") or "Draft", 0) + 1
        onb = self.t_onboarding_status()
        soon = [h for h in onb["hires"] if 0 <= h["working_days_to_start"] <= 3]
        return {
            "as_of": self.as_of.strftime("%a %Y-%m-%d %H:%M"),
            "inbox_unsorted": [r for r in inbox if not str(r.get("priority", "")).strip()],
            "inbox_open": [r for r in inbox if str(r.get("status", "")).strip() not in ("Done",) and str(r.get("priority", "")).strip()],
            "blocks_today": [{k: b.get(k) for k in ("block_id", "start", "end", "workstream", "title")} for b in blocks if str(b.get("date")) == today.isoformat()],
            "blocks_tomorrow": [{k: b.get(k) for k in ("block_id", "start", "end", "workstream", "title")} for b in blocks if str(b.get("date")) == tomorrow.isoformat()],
            "messages_by_status": counts,
            "hires_starting_soon": soon,
            "overdue_by_owner": onb["overdue_by_owner"],
            "investor": self.t_investor_metrics(),
        }

    def t_onboarding_status(self):
        today = self.as_of.date()
        hires, tasks = self.store.rows("Hires"), self.store.rows("Tasks")
        names = {str(c["task_id"]): (c.get("task"), c.get("owner")) for c in self.store.rows("Checklist")}
        out, by_owner = [], {}
        for h in hires:
            hid, start = str(h["hire_id"]), _date(h["start_date"])
            mine = [t for t in tasks if str(t["hire_id"]) == hid]
            done = sum(1 for t in mine if t.get("status") == "Done")
            open_ = [t for t in mine if t.get("status") != "Done"]
            pre = [t for t in open_ if _date(t["due_date"]) < start]
            overdue = [t for t in open_ if _date(t["due_date"]) < today]
            for t in overdue:
                owner = names.get(str(t["task_id"]), ("", "?"))[1]
                by_owner[owner] = by_owner.get(owner, 0) + 1
            wd = sum(1 for k in range((start - today).days) if (today + dt.timedelta(days=k + 1)).weekday() < 5) if start >= today else -1
            out.append({"hire_id": hid, "name": h.get("name"), "start_date": start.isoformat(), "working_days_to_start": wd,
                        "done_pct": round(100 * done / len(mine)) if mine else 0,
                        "open_before_day1": len(pre), "ready_for_day1": len(pre) == 0,
                        "overdue": [{"task": names.get(str(t["task_id"]), (t["task_id"], ""))[0], "owner": names.get(str(t["task_id"]), ("", ""))[1],
                                     "due": str(t["due_date"])} for t in overdue]})
        return {"as_of": today.isoformat(), "hires": out, "overdue_by_owner": by_owner}

    def t_investor_metrics(self):
        def _r1(x):  # half up, the way Excel shows 0.0
            return math.floor(x * 10 + 0.5) / 10

        """Same math as the Investor tab formulas, computed here so it works on a file nobody has opened in Excel."""
        if "Investor" not in self.store.tabs():
            raise ToolError("no Investor tab in this workbook")
        cells = self._investor_cells()
        hotels = {"jul": cells["hotels_jul"], "oct": cells["hotels_oct"]}
        arpu = {"jul": cells["arpu_jul"], "oct": cells["arpu_oct"]}
        arr = {k: hotels[k] * arpu[k] * 12 for k in hotels}
        waiting = cells["waiting"]
        return {
            "hotels_live_oct": hotels["oct"], "hotels_live_jul": hotels["jul"],
            "hotels_growth_vs_jul_pct": round(100 * (hotels["oct"] / hotels["jul"] - 1), 1),
            "revenue_per_hotel_month_oct": arpu["oct"],
            "arr_oct": arr["oct"], "arr_jul": arr["jul"], "arr_growth_vs_jul_pct": round(100 * (arr["oct"] / arr["jul"] - 1), 1),
            "signed_waiting": waiting, "arr_of_waiting": waiting * arpu["oct"] * 12,
            "jan_goal_hotels": hotels["oct"] + waiting, "jan_goal_arr": (hotels["oct"] + waiting) * cells["arpu_goal"] * 12,
            "cash": cells["cash"], "burn_now": cells["burn_now"], "burn_after_hiring": cells["burn_after"],
            "runway_months_now": _r1(cells["cash"] / cells["burn_now"]), "runway_months_after_hiring": _r1(cells["cash"] / cells["burn_after"]),
            "note": "Estimates for the take-home assignment; the CEO confirms ARR, cash and runway before anything goes out.",
        }

    def _investor_cells(self):
        def num(v):
            if isinstance(v, (int, float)):
                return v
            t = str(v or "").replace("$", "").replace(",", "").strip()
            try:
                return float(t)
            except ValueError:
                raise ToolError(f"Investor tab: {v!r} is not a number")

        find = {}
        for row in self.store.grid("Investor")[:60]:
            if len(row) > 1 and isinstance(row[1], str) and row[1].strip():
                find[row[1].strip()] = list(row[2:6]) + [None] * 4
        try:
            hotels, arpu = find["Hotels live (count)"], find["Revenue per hotel per month ($)"]
            return {"hotels_jul": num(hotels[1]), "hotels_oct": num(hotels[2]), "arpu_jul": num(arpu[1]), "arpu_oct": num(arpu[2]),
                    "arpu_goal": num(arpu[3]) if arpu[3] not in (None, "") else num(arpu[2]),
                    "waiting": num(find["Signed hotels waiting to go live"][1]), "cash": num(find["Cash in bank ($)"][0]),
                    "burn_now": num(find["Net burn per month today ($)"][0]), "burn_after": num(find["Net burn per month after hiring ($)"][0])}
        except KeyError as e:
            raise ToolError(f"Investor tab is missing the row {e}")

    # ---- write tools (the only ways the agent changes anything)
    def t_triage_item(self, id, priority, rule, plan, due=None):
        if rule not in RULES:
            raise ToolError(f"rule must be one of {list(RULES)}")
        if priority not in ("P0", "P1", "P2"):
            raise ToolError("priority must be P0, P1 or P2")
        plan = str(plan).strip()
        if not plan or len(plan) > 160:
            raise ToolError("plan must be one line under 160 characters")
        row = next((r for r in self.store.rows("Inbox") if str(r.get("id")) == str(id)), None)
        if row is None:
            raise ToolError(f"no inbox item with id {id}")
        if str(row.get("status")) == "Done":
            raise ToolError(f"inbox item {id} is already Done; leave it")
        self.store.set("Inbox", row["_row"], "priority", priority)
        self.store.set("Inbox", row["_row"], "resolution", "Plan: " + plan)
        if due and not str(row.get("due", "")).strip():
            self.store.set("Inbox", row["_row"], "due", str(due))
        if not str(row.get("status", "")).strip():
            self.store.set("Inbox", row["_row"], "status", "Open")
        self.log("triage_item", "Inbox", row["_row"], f"#{id} {row.get('item')}: {priority}, {RULES[rule]}. {plan}")
        self.writes.append(("Inbox", row["_row"]))
        return {"ok": True, "id": id, "priority": priority, "rule": RULES[rule]}

    def t_add_block(self, date, start, end, workstream, inbox_ids, title, steps, thinking):
        d, s, e = _date(date), _hm(start), _hm(end)
        if e <= s:
            raise ToolError("end must be after start")
        if s < 7 * 60 or e > 20 * 60:
            raise ToolError("blocks stay between 7:00 and 20:00")
        if d.weekday() >= 5:
            raise ToolError("no blocks on weekends")
        if workstream not in WORKSTREAMS:
            raise ToolError(f"workstream must be one of {WORKSTREAMS}")
        if not str(title).strip():
            raise ToolError("title is required")
        blocks = self.store.rows("Blocks")
        for b in blocks:
            if str(b.get("date"))[:10] == d.isoformat():
                bs, be = _hm(b["start"]), _hm(b["end"])
                if s < be and bs < e:
                    raise ToolError(f"overlaps {b['block_id']} {b['start']}-{b['end']} '{b.get('title')}'. Pick a free slot or log a note.")
        n = 1 + sum(1 for b in blocks if str(b.get("block_id", "")).startswith("AG"))
        bid = f"AG{n:02d}"
        steps_txt = "\n".join(str(x).strip() for x in (steps or []) if str(x).strip())
        r = self.store.append("Blocks", {"block_id": bid, "date": d.isoformat(), "start": f"{s // 60}:{s % 60:02d}", "end": f"{e // 60}:{e % 60:02d}",
                                         "workstream": workstream, "inbox_ids": str(inbox_ids or ""), "title": str(title).strip(),
                                         "steps": steps_txt, "thinking": str(thinking).strip(), "short_note": "Added by the agent"})
        self.log("add_block", "Blocks", r, f"{bid} {d.isoformat()} {start}-{end} {title}")
        self.writes.append(("Blocks", r))
        return {"ok": True, "block_id": bid, "row": r}

    def t_draft_message(self, block_id, channel, to, subject, text):
        if self.drafts >= self.max_drafts:
            raise ToolError(f"draft limit reached ({self.max_drafts} per run). Log a note instead.")
        if channel not in CHANNELS:
            raise ToolError(f"channel must be one of {CHANNELS}")
        if not any(str(b.get("block_id")) == str(block_id) for b in self.store.rows("Blocks")):
            raise ToolError(f"no block {block_id!r}; add the block first")
        text, subject = str(text).strip(), str(subject or "").strip()
        if not text:
            raise ToolError("text is required")
        if channel.startswith("Email") and not subject:
            raise ToolError("emails need a subject")
        for p in SECRET_PATTERNS:
            if re.search(p, text) or re.search(p, subject):
                raise ToolError("that draft looks like it contains a key, token or card number; remove it")
        r = self.store.append("Messages", {"block_id": str(block_id), "channel": channel, "to": str(to).strip(), "send_to": "",
                                           "subject": subject, "text": text, "status": "Draft", "sent_at": ""})
        self.drafts += 1
        self.log("draft_message", "Messages", r, f"{channel} to {to}: {_clip(text, 140)} [Draft, waits for approval]")
        self.writes.append(("Messages", r))
        return {"ok": True, "row": r, "status": "Draft", "note": "Saved as Draft. Not sent; a person marks it Ready."}

    def t_log_note(self, text):
        self.log("note", "", "", str(text).strip())
        return {"ok": True}
