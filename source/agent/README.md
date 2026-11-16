# Lance CoS agent

The agent runs my daily loop on the CoS workbook with Claude. It sorts new inbox items with my three rules, puts the work on my calendar, writes every email and Slack as a **Draft**, checks that new hires are ready for Day 1, and drafts the investor update. It never sends anything. I approve drafts in the week calendar app; the Apps Script in the plug-in kit sends only rows I mark Ready.

```
Inbox, Hires, Tasks, Investor  ──►  Claude agent (8 AM, 5:15 PM)  ──►  Draft rows + calendar blocks + Agent Log
   (the workbook / Google Sheet)       sorts, plans, drafts              │
                                                                         ▼
                                     I review in the calendar app, mark Ready  ──►  Apps Script sends
```

## What's in this folder

| File | What it is |
|---|---|
| `system_prompt.md` | The agent's instructions: my rules, the 4 tasks, writing style, safety. Also the **Claude Project instructions** (paste it in). |
| `SKILL.md` | The same job as a **Claude skill**, for Claude Code or Claude with code execution. Claude runs the tools itself. |
| `lance_agent.py` | The harness: the Claude API loop, run limits, logging, error handling. |
| `tools.py` | The 8 tools and the rules they enforce. Works on the Excel workbook or the live Google Sheet. |
| `tools_cli.py` | One tool from the command line (used by the skill). |
| `test_agent.py` | 16 offline tests: a scripted fake Claude, plus the real SDK against a mock server. No API key needed. |
| `requirements.txt`, `.env.example` | Setup. |

## The 8 tools

`get_briefing` (the morning picture in one call) · `read_tab` · `triage_item` · `add_block` · `draft_message` · `onboarding_status` · `investor_metrics` · `log_note`

There is no send tool and no delete tool. The rules live in code, not just in the prompt:

- every message is written with status Draft; the agent can't set Ready or Sent
- at most 30 drafts and 40 model calls per run (Agent tab)
- no double-booked blocks, no weekend blocks, times between 7:00 and 20:00
- items already Done can't be re-sorted
- drafts with something that looks like an API key, Slack token or card number are refused
- `mode` on the Agent tab must say `draft_only` or the agent won't start
- every write, refusal and summary goes to the Agent Log tab

## Run it (5 minutes)

```bash
pip install -r requirements.txt
cp .env.example .env            # add ANTHROPIC_API_KEY, or run `ant auth login`
python test_agent.py            # 16 offline checks, no key needed

# a dry run on a copy (the workbook itself is untouched)
python lance_agent.py --workbook ../Lance_CoS_Workbook_Basav_Sharma.xlsx --task morning --out monday_try.xlsx

python lance_agent.py --workbook ../Lance_CoS_Workbook_Basav_Sharma.xlsx --task morning    # 8 AM
python lance_agent.py --workbook ../Lance_CoS_Workbook_Basav_Sharma.xlsx --task eod        # 5:15 PM note to the CEO
python lance_agent.py --workbook ../Lance_CoS_Workbook_Basav_Sharma.xlsx --task investor   # monthly
python lance_agent.py --workbook ../Lance_CoS_Workbook_Basav_Sharma.xlsx --ask "Move the insurance pack to Tuesday and tell the COO"
```

Then open the week calendar app, press Connect data, pick the workbook: the agent's blocks and drafts are on the calendar, ready for review. A run in place keeps a backup next to the file (`*.before-agent.xlsx`).

## Go live on Google

1. Put the workbook's tabs in Google Sheets (the plug-in kit's template is the same layout).
2. Make a Google Cloud service account, share the Sheet with its email, download its key as `sa.json`.
3. `pip install gspread`, then `python lance_agent.py --sheet <sheet id> --credentials sa.json --task morning`.
4. Schedule it: cron on any small server (`0 8 * * 1-5` and `15 17 * * 1-5`, Pacific time), or a scheduled GitHub Action.
5. New inbox items come in as rows with a blank priority: a Slack Workflow form ("Add to CoS inbox") or a Gmail filter + Zapier can add them. The agent picks them up at the next run.

The keys stay in environment variables, never in the Sheet.

## Use it inside Claude instead

- **Claude Project:** paste `system_prompt.md` as the project instructions and add the workbook as a file. Ask "run the morning loop"; Claude answers with the plan and the drafts for me to copy.
- **Claude skill:** zip this folder (it has `SKILL.md` at the top) and add it as a skill. In Claude Code or Claude with code execution, Claude runs `tools_cli.py` on the workbook, with the same rules.

## Model and cost

Default model: `claude-opus-5-5` with adaptive thinking at high effort (set on the Agent tab; `claude-sonnet-5-5` is the cheaper option for routine days). The system prompt and tool list are fixed, so they are cached across the turns of a run. If the model declines a request on a safety check, the API retries it on a fallback model (`fallbacks: "default"`); a full refusal stops the run and is logged.

## Limits

- The offline tests prove the loop, the rules and the workbook writes. The quality of the sorting and the drafts needs a real run with an API key; review the first week of drafts before trusting it.
- The agent sees only what is in the workbook. It doesn't read Gmail or Slack directly; that's deliberate for a first version.
