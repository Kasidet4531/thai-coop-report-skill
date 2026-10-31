# Thai Co-op Report Skill

An agent skill that builds your **รายงานสหกิจศึกษาฉบับสมบูรณ์** (Thai co-op education final report) as a Word file that already matches the department format, so you don't spend an evening fighting margins, Thai fonts and page numbers.

Tell your agent what you have: a few cover details, a messy draft, or notes from your internship. You get back a `.docx` with the covers, both abstracts, the acknowledgement, an automatic table of contents, chapters 1–5, references, appendices and the author biography, all laid out correctly. Anything you haven't provided shows up as a red `[placeholder]`. The skill never makes up results or names.

Works with Claude Code, Claude Desktop / Claude.ai, Codex, Antigravity, and any agent that reads `SKILL.md` skills.

## Installation (1-minute setup)

### 1. Install Python packages (once)

```bash
pip install python-docx pythainlp
```

Also install the **TH SarabunPSK** font (free from SIPA or f0nt.com). Without it, Word shows a substitute font on screen.

### 2. Get the skill

<details>
<summary><strong>Claude Code</strong></summary>

```
/plugin marketplace add Kasidet4531/thai-coop-report-skill
/plugin install coop-report-docx@thai-coop-report
```

To update later: `/plugin marketplace update thai-coop-report`.

</details>

<details>
<summary><strong>Codex / Antigravity</strong> (one-line installer)</summary>

Windows (PowerShell):

```powershell
& ([scriptblock]::Create((irm https://raw.githubusercontent.com/Kasidet4531/thai-coop-report-skill/main/install.ps1))) -Target codex,antigravity
```

macOS / Linux:

```bash
curl -fsSL https://raw.githubusercontent.com/Kasidet4531/thai-coop-report-skill/main/install.sh | bash -s -- codex antigravity
```

The installer downloads the skill into the folder each agent scans and installs the Python packages for you. Choose any combination of targets:

| Target | Agent | Folder |
|---|---|---|
| `codex` | Codex (CLI / IDE / app) | `~/.agents/skills/` |
| `antigravity` | Antigravity IDE and 2.0 | `~/.gemini/config/skills/` |
| `antigravity-cli` | Antigravity CLI (`agy`) | `~/.gemini/antigravity-cli/skills/` |
| `claude` | Claude Code, without the plugin | `~/.claude/skills/` |
| `project` | Current project only (Codex and Antigravity) | `./.agents/skills/` |
| `all` (default) | claude + codex + antigravity | |

Restart the app afterwards. Re-run the same command to update. Older Codex versions read `~/.codex/skills/`; if Codex doesn't see the skill, copy the folder there too.

</details>

<details>
<summary><strong>Claude Desktop / Claude.ai</strong></summary>

1. Click **Code → Download ZIP** on this page and unzip it.
2. Zip the `skills/coop-report-docx` folder on its own.
3. Go to **Settings → Capabilities → Skills → Upload skill** and pick that zip.

Claude.ai on the web runs skills in a sandbox that may not have `pythainlp`. Without it, body text falls back to left alignment instead of Thai Distributed.

</details>

<details>
<summary><strong>Any other agent</strong></summary>

Copy `skills/coop-report-docx` into the folder your agent loads skills from.

</details>

### 3. Ask for your report

In Thai or English, for example:

> ช่วยทำเทมเพลตเล่มสหกิจให้หน่อย ชื่อ นายสมชาย ใจดี ภาควิชาวิศวกรรมไฟฟ้า ฝึกที่บริษัท ABC ปีการศึกษา 2568

> ไฟล์ draft.docx เป็นร่างเล่มสหกิจ ช่วยจัดให้ถูก format หน่อย เนื้อหาห้ามเปลี่ยน

> Here are my internship notes. Draft my co-op report and leave anything I haven't measured as placeholders.

When you open the file, click **Yes** when Word asks to update fields. That fills in the table of contents.

## Why This Skill Exists

We ran the same three requests through an agent with and without the skill, two rounds each, and checked every output file with a script. With the skill, every report passed every check. Without it, reports passed about 62% of checks, and they failed in the same ways each time.

### #1: The Agent Doesn't Know Your Department's Format

**The problem.** Without the format, an agent falls back on generic Thai thesis conventions. In our tests those defaults broke these department rules:

| Rule | Without the skill | Department format |
|---|---|---|
| Margins | 1.5" top and left | 3 cm left; 2.54 cm top, right, bottom |
| Chapter titles | 20 pt | 18 pt bold, centred |
| Front-matter page numbers | ก ข ค or i ii iii | I II III |
| Page number position | top right | bottom centre |
| Chapter first page | numbered | no number |
| Font | TH Sarabun New | TH SarabunPSK |

**The fix.** The layout comes from a script (`build_report.py`), not from the model's memory, so it is the same every time. A second script (`check_format.py`) checks the result before you get it.

### #2: Thai Text Looks Broken in Word

**The problem.** Thai has no spaces between words, and many Word installations can't find Thai word boundaries. They wrap lines only at spaces, so Thai Distributed alignment spreads the letters across the line: "ว ิ ธ ี ก า ร".

**The fix.** The script splits Thai text into words with [PyThaiNLP](https://github.com/PyThaiNLP/pythainlp) and marks each boundary with an invisible zero-width space. Word then wraps between words, and the lines fill properly.

### #3: The Agent Fills Gaps With Plausible Fiction

**The problem.** A co-op report has results, names and dates the agent can't know. A made-up satisfaction score in a submitted report is much worse than a visible gap.

**The fix.** The skill tells the agent to leave missing facts as red `[placeholders]`, and the checker lists every placeholder still to fill.

## What's Inside

```
skills/coop-report-docx/
├── SKILL.md                    # the workflow the agent follows
├── scripts/
│   ├── build_report.py         # content.json → formatted .docx
│   ├── check_format.py         # verifies the format, lists placeholders
│   └── finalize_with_word.ps1  # Windows + Word: refresh the TOC, export a PDF
├── references/
│   ├── format-rules.md         # the department format, plus conventions from a reviewed report
│   └── content-schema.md       # every field and option in content.json
└── evals/                      # test cases (claude plugin eval)
```

### Use the scripts without an agent

```bash
# blank template
python skills/coop-report-docx/scripts/build_report.py --template report.docx

# from a content file (see references/content-schema.md)
python skills/coop-report-docx/scripts/build_report.py content.json report.docx

# check the format
python skills/coop-report-docx/scripts/check_format.py report.docx
```

## Adapting It to Another University

The defaults follow the Faculty of Engineering at KMITL (สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง).

- **Institute and faculty names:** tell the agent where you study, or set `institute` and `faculty` in `content.json`.
- **Layout options:**
  - `reference_style`: `numbered` [1] or `apa`
  - `chapter_first_page_number`: show page numbers on chapter first pages
  - `justify`: `thaiDistribute`, `both` or `left`
  - toggles for the outer cover and the flyleaf
- **Margins, sizes and spacing:** edit the constants at the top of `build_report.py`. The rules are documented in `references/format-rules.md`.

## License

MIT
