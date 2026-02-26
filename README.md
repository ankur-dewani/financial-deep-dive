# Deep Dive 1: Finance and Accounting

**Candidate**: Ankur Dewani | **Role**: VP of Operations | **Date**: February 2026

A root cause analysis of why a $79.2M business unit operates at 8.86% margin versus the 70% target, focused on Finance and Accounting as the deep dive function. The analysis uses the Central Finance Roles provided by the evaluator to model a specific fix with quantified savings.

---

## Project Structure

```
dd1/
├── scripts/
│   ├── cost_model.py              # Central Finance cost modeling (current vs target)
│   ├── analyze_pl.py              # P&L analysis engine (reads 8 sheets, writes 4 analysis sheets)
│   └── generate_dd1.py            # DD1 document generator (fills template docx)
├── data/
│   ├── input/
│   │   ├── Deep Dive 1 - Template.docx
│   │   ├── Central Finance Roles.xlsx
│   │   └── Operational Leadership Real Work - Input P&L.xlsx
│   └── output/
│       ├── DD1 - Ankur Dewani.docx
│       └── Input P&L - Ankur Dewani.xlsx
└── README.md
```

## Prerequisites

| Requirement | Version | Purpose |
|---|---|---|
| Python | 3.8+ | Runtime |
| openpyxl | 3.1+ | Excel workbook read/write |
| python-docx | 1.0+ | Word document generation |

```bash
pip3 install openpyxl python-docx
```

## Usage

Generate both output files:

```bash
cd dd1/
python3 scripts/analyze_pl.py     # Produces Input P&L with 4 analysis sheets
python3 scripts/generate_dd1.py   # Produces DD1 document
```

---

## The Problem

This business unit generates $79.2M in annual revenue but retains only $7M (8.86% margin). The target is 70% margin with total expenses at 30% of revenue. Actual expenses are 91.1%.

| Benchmark Category | Actual % | Target % | Gap |
|---|---|---|---|
| Shared Services | 12.0% | 4.5% | +7.5% |
| Executive Team | 7.9% | 4.5% | +3.4% |
| Sales | 14.4% | 5.0% | +9.4% |
| Marketing | 3.2% | 1.0% | +2.2% |
| Technical Support | 5.3% | 2.0% | +3.3% |
| Hosting | 9.3% | 1.0% | +8.3% |
| Product | 17.2% | 2.0% | +15.2% |
| Engineering | 21.8% | 10.0% | +11.8% |
| **Total** | **91.1%** | **30.0%** | **+61.1%** |

Every category is over model. The acquisition was never operationally transformed.

---

## Deep Dive: Finance and Accounting

### Why F&A

The evaluator provided a Central Finance Roles spreadsheet with 5 specific roles and hourly rates. This signals where the deep dive should focus. F&A also has the clearest single function story:

- F&A costs $3.82M (4.8% of revenue)
- That alone exceeds the entire Shared Services benchmark (4.5%)
- 51% of F&A spend flows to external providers at unmanaged rates
- The Central Finance Roles file provides exact target costs for the fix

### 5 Why Root Cause Chain

| Why | Question | Answer |
|---|---|---|
| 1 | Why is this BU not in model? | Still running pre acquisition cost structure. No function was restructured for Central Factory. All 8 categories over. |
| 2 | Why start with Shared Services when Engineering has a larger gap? | Engineering requires inventing a role structure. F&A has a predefined Central Finance model with exact target costs, enabling a specific fix in 16 weeks. F&A alone ($3.82M, 4.8%) exceeds the entire SS benchmark. |
| 3 | Why does F&A cost $3.82M? | 51% ($1.96M) goes to external providers. 18 staff coordinate vendor handoffs instead of executing standardized processes. |
| 4 | Why is 51% going to external providers? | Each entity kept its own audit firm, tax advisor, and close process. Duplicate work across entities, no consolidation. |
| 5 | Why do these fragmented processes persist? | No one was tasked with integration. A leadership gap, not a technical one. The Central Finance Roles define the target state. |

### Root Cause

No one was tasked with integrating Finance and Accounting into the Central Finance model. The function runs as a fragmented multi entity operation where 51% of cost flows to external providers doing duplicated work, and 18 internal staff coordinate vendor handoffs rather than executing standardized processes. The fix exists. The gap is execution.

---

## The Fix: Central Finance Model

Replace fragmented F&A with a Central Finance team using standardized roles:

| Role | Count | Annual Cost | Total |
|---|---|---|---|
| VP of Finance | 1 | $200,000 | $200,000 |
| Finance Manager | 2 | $100,000 | $200,000 |
| Senior Accountant | 5 | $60,000 | $300,000 |
| Accountant | 10 | $30,000 | $300,000 |
| Statutory Audit (retained) | | | $200,000 |
| **Total** | **18** | | **$1,200,000** |

| Metric | Value |
|---|---|
| Current F&A cost | $3,822,076 |
| Target in model cost | $1,200,000 |
| Annual savings | $2,622,076 |
| Reduction | 69% |

Implementation: 16 weeks. Simplify first: document all F&A processes, map duplicates, write standardized close checklist. Then migrate: extend Central finance platform to this BU, unify chart of accounts, consolidate audit firms. Run first close cycle on new model by week 16.

---

## AI Opportunities

**Area 1: LLM Driven Financial Close**
I configure an LLM reconciliation agent on the unified GL. It auto matches 100% of intercompany transactions, generates journal entries, and routes only true exceptions for human review. Close cycle drops from 15 to 20 days to under 5 days. Reduces Senior Accountant need from 5 to 3. Additional savings: $120K/year.

**Area 2: AI Spend Intelligence**
I build a Claude Code pipeline on existing AP data. It classifies vendor spend daily, detects duplicates, flags benchmark non compliance, and produces a weekly list of vendors to cut, renegotiate, or consolidate. Prevents post transformation cost drift. Additional savings: $200K to $300K/year.

---

## Analysis Sheets Added to P&L Workbook

The output P&L workbook preserves all 8 original sheets and adds 5 analysis sheets. All 8 source sheets are programmatically read: Benchmarks and P&L Summary for targets and totals, OPEX/COGS/Empl for expense line items, and all 3 Revenue sheets for revenue mix analysis.

| Sheet | Source Sheets Used | Purpose |
|---|---|---|
| Revenue Analysis | P&L Summary, RecurringRevenue, PSORevenue, PerpetualRevenue | Revenue mix breakdown (87% recurring, 12% PSO, 2% perpetual) with P&L summary |
| Benchmark Mapping | Benchmarks, Empl, OPEX-NEmpl, COGS-NEmpl | Every expense mapped to 8 benchmark categories with actual vs target % |
| SS Breakdown | Empl, OPEX-NEmpl | Shared Services broken into 11 G&A sub departments with HC and non-HC costs |
| FA Deep Dive | Empl, OPEX-NEmpl, Central Finance Roles | F&A cost by component, Central Finance target model, savings calculation |
| FA Employee Analysis | Empl, Central Finance Roles | 18 current employees mapped to Central Finance role tiers by salary band |

---

## Tools and Approach

| Tool | Role |
|---|---|
| **Claude Code CLI** | Primary analysis engine: P&L parsing, benchmark mapping, cost modeling, document generation |
| **Python 3 + openpyxl** | Data processing, pivot table generation, XLSX output |
| **Python 3 + python-docx** | DD1 template filling with formatted text |

---

## Reflection

### How I Structured the Analysis

I started by reading all 8 sheets of the P&L workbook to understand the full cost structure before choosing a deep dive target. The Benchmark Mapping showed every category over model, with absolute gaps ranging from $900K (Marketing) to $12M (Product). Engineering has the largest dollar gap at $9.4M.

But the question is not "what is the biggest gap?" The question is "where can I deliver a fix that works, with exact dollar savings?" The Central Finance Roles spreadsheet provided exact roles, hourly rates, and annual salaries for a standardized finance function. That is not a suggestion. It is a target model. F&A can be fixed to that target in 16 weeks with specific dollar outcomes. Engineering cannot, because no target role structure exists for it yet. F&A is the proof of concept. Engineering and Sales follow using the same methodology once F&A proves it works.

### Where AI Was Used

Claude Code drove every step of this analysis. It parsed 2,092 OPEX rows and 1,000 COGS rows to build the benchmark mapping. It identified the F&A cost anomaly by aggregating employee and non employee data simultaneously, something that would have taken multiple manual pivot tables. It built the cost model by reading the Central Finance Roles spreadsheet and mapping current salaries to role tiers by salary band. And it generated both output files (the annotated P&L workbook and the DD1 document) programmatically.

The important thing is not that I used AI. The important thing is that AI made the analysis repeatable. If I receive another P&L workbook next week for a different acquisition, the same scripts produce the same structured output in minutes. That matters when the operating model targets one acquisition per week.

### Key Judgment Decisions

Three decisions required human judgment, not just data processing:

1. **Choosing F&A over Engineering.** Engineering has a $9.4M gap but no predefined target model. F&A has the Central Finance Roles file with exact roles and rates. The right move is to fix what you can specify precisely, prove the methodology, then apply it to larger functions. Fixing F&A first also means simplifying the function before moving it to Central, not dumping a fragmented process onto a shared service.

2. **Setting the target team at 18 people.** I kept total headcount at 18 but restructured the role mix to Central Finance tiers. The savings come from eliminating $1.96M in duplicated outsourced work, not from layoffs. Consolidating 7+ audit engagements into 1 to 2, unifying the chart of accounts, and standardizing the close process removes the duplication that made external providers necessary.

3. **Keeping $200K for statutory audit.** External audit is legally required and cannot be brought in house. Rather than claim 100% savings, I retained $200K for statutory compliance. This makes the model credible rather than aspirational.

### What I Would Do Next

1. **Apply this methodology to Engineering and Sales.** Engineering has a $9.4M gap. Sales has a $7.4M gap. With the F&A proof of concept complete, I apply the same process: benchmark mapping, 5 Why root cause, Central Factory target model, phased migration.

2. **Vendor level breakout.** The $1.43M in outsourced services is an aggregate. With individual vendor data, I build a contract by contract termination schedule with specific dates and notice periods.

3. **Conservative scenarios.** The current model assumes all 18 employees fit Central Finance roles. In practice, some will not pass skills assessments. I would run 70% and 85% scenarios alongside the 100% target to give the CEO a realistic range.

4. **Post migration KPIs.** Quarterly tracking: close cycle under 5 days, F&A cost under $1.2M annualized, outsourced spend under $250K, zero duplicate vendor payments. Without these metrics, cost structures drift back.

The Central Factory methodology turns "this feels expensive" into exact dollar gaps with a specific fix. "F&A costs 4.8% of revenue versus a 4.5% Shared Services benchmark, and 51% of that cost flows to unmanaged external providers." That precision is what I use to run a 16 week transformation with confidence in the outcome.
