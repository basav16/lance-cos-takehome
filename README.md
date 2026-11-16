# Lance Chief of Staff take-home · Basav Sharma

**Live site: https://basav16.github.io/lance-cos-takehome/**

Start with the thinking deck. It links to every app and file below.

| # | Piece | Part | Live link | File in this repo |
|---|---|---|---|---|
| 1 | Thinking deck (18 slides) | All parts | [PDF](https://basav16.github.io/lance-cos-takehome/decks/Lance_CoS_Thinking_Deck_Basav_Sharma.pdf) | `decks/Lance_CoS_Thinking_Deck_Basav_Sharma.pdf` |
| 2 | Q3 2026 investor update (13 slides) | 2B | [PDF](https://basav16.github.io/lance-cos-takehome/decks/Lance_Q3_2026_Investor_Update.pdf) | `decks/Lance_Q3_2026_Investor_Update.pdf` |
| 3 | Week calendar app | 1 | [Open](https://basav16.github.io/lance-cos-takehome/week-calendar/) | `week-calendar/index.html` |
| 4 | Onboarding Hub | 2A | [Open](https://basav16.github.io/lance-cos-takehome/onboarding-hub/) | `onboarding-hub/index.html` |
| 5 | Investor dashboard | 2B | [Open](https://basav16.github.io/lance-cos-takehome/investor-dashboard/) | `investor-dashboard/` |
| 6 | CoS workbook | 1 and 2 | [Download](https://basav16.github.io/lance-cos-takehome/files/Lance_CoS_Workbook_Basav_Sharma.xlsx) | `files/Lance_CoS_Workbook_Basav_Sharma.xlsx` |
| 7 | Extra: CoS agent and plug-in kits | Optional | [Download](https://basav16.github.io/lance-cos-takehome/files/Lance_CoS_Agent_and_Kits_Basav_Sharma.zip) | `files/…zip`, unpacked in `source/` |

## How the pieces fit

- **The workbook is the one source of data.** Both apps load it. In Google Sheets it becomes the live database for the team.
- **The apps are single HTML files.** They open in any browser with no install and run on sample data. Each has a short quick tour.
- **The plug-in kits** (`source/plug-in-kits/`) hold the Google Apps Script and template for each app. With them, ticks and approvals write to the Sheet so everyone sees them live. In these demos, ticks save in your browser.
- **The agent** (`source/agent/`) is optional. It runs my daily loop on the workbook with the Claude API and writes drafts only. It has no send or delete tool. The API key stays in an environment variable.

## Notes

- Lance’s internal figures aren’t public. Investor numbers are my estimates, marked *, until finance confirms.
- Customer quotes are word for word from Lance’s published customer stories.
- At Lance, each tool would sit on Lance’s own domain, behind work-email sign-in. This repo is a public demo with sample data only.
