"""Call one agent tool from the command line, with the same rules the agent has (drafts only, no overlaps, logged).
Used by the Claude skill (SKILL.md), where Claude itself is the model and runs these commands.

    python tools_cli.py WORKBOOK get_briefing
    python tools_cli.py WORKBOOK triage_item '{"id": 8, "priority": "P1", "rule": "outside_party_date_or_cost", "plan": "Send the letter by Tue noon"}'
    python tools_cli.py WORKBOOK draft_message '{"block_id": "AG01", "channel": "Email", "to": "Vendor AR", "subject": "Payment", "text": "..."}'
    python tools_cli.py WORKBOOK tools          # list the tools and their inputs
"""
import datetime as dt
import json
import sys

from tools import ExcelStore, Toolbox, ToolError


def main(argv):
    if len(argv) < 2:
        print(__doc__)
        return 2
    path, name = argv[0], argv[1]
    if name == "tools":
        print(json.dumps([{"name": t["name"], "description": t["description"], "input": t["input_schema"]["properties"]} for t in Toolbox.definitions()], indent=1))
        return 0
    args = json.loads(argv[2]) if len(argv) > 2 else {}
    store = ExcelStore(path, backup=False)
    as_of = next((str(r.get("value")) for r in store.rows("Agent") if r.get("key") == "as_of" and r.get("value")), "") if "Agent" in store.tabs() else ""
    box = Toolbox(store=store, run_id="skill-" + dt.datetime.now().strftime("%m%d-%H%M"), task="skill",
                  as_of=dt.datetime.fromisoformat(as_of) if as_of else dt.datetime.now())
    try:
        out = box.run(name, args)
    except (ToolError, TypeError) as e:
        print(json.dumps({"error": str(e)}))
        return 1
    store.save()
    print(json.dumps(out, indent=1, default=str))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
