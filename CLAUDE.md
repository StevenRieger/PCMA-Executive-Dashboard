# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose and Context

Steven Rieger is Director of Enterprise Systems, Technology and Data at PCMA (Professional Convention Management Association), a global professional association for the business events industry, headquartered in Chicago. PCMA has approximately 69 staff and a fragmented technology stack of 35+ applications. Steven's core mandate spans enterprise systems modernization, AI governance and adoption, data infrastructure, and organizational discovery as a relatively recent hire.

Key colleagues: Sherrif Karamat (CEO), Ori Klein (CTO), Kimberly Maggio (CFO), Barbara Palmer (Convene magazine). Steven is the policy/systems owner on most initiatives, with Ori as technical implementation lead and Sherrif as executive sponsor/approver.

PCMA's primary platforms include Fonteva (OEM Salesforce AMS - targeted for replacement), RainFocus (event execution), Microsoft 365 (SharePoint, Power Automate, Teams), Business Central (ERP), Pardot (marketing automation), WordPress (chapter sites), Stripe (e-commerce), and LearnUp (LMS).

## What This Is

A single-page executive dashboard for PCMA's 2026 Enterprise Goals. It is a self-contained `index.html` (HTML + CSS + JS, no build step) that displays KPI progress across five weighted sections with live data polling. Hosted on GitHub Pages at stevenrieger.github.io/PCMA-Executive-Dashboard.

## Running

Open `index.html` directly in a browser, or serve it with any static server (e.g. `npx serve .`). Live data polling requires the file to be served over HTTP so `fetch()` works.

## Architecture

**Single file:** All markup, styles, and logic live in `index.html` (~809 lines). There is no build system, bundler, or framework.

**Data flow:**
1. Hardcoded fallback data lives in the `DATA` constant (line ~252) with five sections: `financial`, `membership`, `endowment`, `education`, `internal`.
2. On load, `showLoadingState()` runs, then `tryLiveData()` fetches `dashboard_data.json` (generated externally by Power Automate from a SharePoint Excel list).
3. `parseExcelRows()` (line ~639) maps SharePoint row names to section/metric indices via `NAME_MAP` - this is the critical data-binding layer. When SharePoint column names change, `NAME_MAP` fragments must be updated.
4. `buildDashboard()` and `buildComposite()` re-render the entire UI from the `DATA` object.
5. Polls every 30 seconds via `setInterval(tryLiveData, 30000)`.

**Key functions:**
- `parseExcelRows()` - Converts raw SharePoint rows into section updates. Uses name-fragment matching, not positional indexing.
- `buildDashboard()` - Generates all section blocks, cubes, detail panels, and chart canvases.
- `buildDetailCharts()` - Creates Chart.js bar/progress charts when a section is expanded.
- `secAvg()` / `compositeScore()` - Calculate weighted progress scores displayed in the header and thermometers.
- `pColor()` / `chipClass()` - Determine RAG (red/amber/green) status colors based on progress percentage.

**External dependency:** Chart.js loaded via CDN (`<script>` tag, line ~242).

**Metric properties:** Each metric object can have: `goal`, `ytd`, `gl` (goal label), `yl` (YTD label), `fmt` (currency/number/percent), `src` (data source name), `inverse` (lower-is-better, e.g. turnover), `threshold` (target-based), `exceeds` (exceeded goal flag), `status` (text status like "Quiet Phase"), `checklist` (array of {label, done} for the Data & Technology initiatives), `budget`, `variance`.

## Key Conventions

- Section weights (in the `DATA` config) must sum to 100 - they drive the composite score calculation.
- The `NAME_MAP` in `parseExcelRows()` uses substring matching (`includes()`) against the SharePoint "Financial Results" column. Order matters: first match wins.
- Financial metrics use a special pattern: the goal row maps to the annual goal, but YTD/budget/variance come from separate named rows (e.g. "PCMA: Consolidated February 2026 YTD Net Operating Income Actual").
- Internal metrics with decimal YTD values < 1 are auto-converted to percentages (multiplied by 100).
- The `dashboard_data.json` file is machine-generated - do not hand-edit it. To change displayed data, modify the `DATA` constant in `index.html` or update the upstream SharePoint list.
- CORS is a hard architectural constraint for dashboards hosted on external domains - JSON must be served from the same origin (GitHub Pages) rather than SharePoint to avoid cross-origin fetch failures.
- Excel dashboard data should be parsed by name, not row position, so adding rows does not break the pipeline.

## Current State and Active Workstreams

- **AMS replacement** - RFP in progress to replace Fonteva on Salesforce. Leading candidates are iMIS (primary) and Protech on Dynamics 365. Key risk: Momentive Software now controls most enterprise AMS products after acquiring Personify in January 2026.
- **AI governance** - Formal policy completed with five deliverables: Word policy document, executive PowerPoint, SharePoint HTML living document, SharePoint list schema, and Power Automate automation flows. Four-tier data classification and three-tier approved tools list (Enterprise: Claude M365 add-in and Microsoft 365 Copilot; Conditional: paid personal accounts; Prohibited: free-tier tools).
- **Executive dashboards** - Two live dashboards hosted on GitHub Pages pulling from SharePoint Excel via Power Automate. CEO Enterprise Goals Dashboard (this repo) and Business Unit Revenue Dashboard. Auto-refresh every 30 seconds via GitHub API JSON pipeline. Repo: github.com/StevenRieger/PCMA-Executive-Dashboard.
- **Insights and Consulting product launch** - Outlook 2026 Report suite (seven PDFs) targeting launch. Access architecture under evaluation: SharePoint-native guest OTP approach (used by EIC) is the leading candidate.
- **Copier lease return** - Equipment rental with American Capital Financial Services ends October 2026. Formal return notice targeting June 1 send date (90-day notice required). Ori Klein to sign. MOE (moetrans.com) is top recommended deinstall/shipping vendor.
- **Annual goals** - Steven's personal goals and team goals drafted for annual review cycle, covering SOP library, AI governance/adoption, AMS RFP, executive reporting infrastructure, and Insights product launch.

## On the Horizon

- Salesforce evaluation - Assessing whether Salesforce's native event management product could replace Fonteva, framed around a "Data 360" concept introduced by Kim Maggio.
- Org discovery completion - Onboarding interviews with approximately 69 staff, with personalized questionnaire documents and tracker spreadsheet in use.
- Interactive org chart tool - Planned web-based tool where clicking a person's node reveals interview response summaries.
- Power Automate premium license - 90-day trial accepted for HTTP connector use in dashboard pipeline; long-term licensing decision pending.
- LiteSpeed ADC cache exclusion (Fix 2) - Recommended as a precaution after Cloudflare Rocket Loader fix resolved the Gravity Forms Save and Continue issue on pcmainstitute.org.

## Key Learnings and Principles

- Fonteva's "Shippable" checkbox on product records triggers shipping address requirements in payment flows - a non-obvious configuration cause.
- SharePoint's native OTP guest access (used by EIC) is a viable no-additional-cost alternative to third-party DRM tools for gated PDF delivery.
- OneDrive Request Files requires OneDriveRequestFilesLinkEnabled set to True at the tenant level via PowerShell - it is off by default.
- Cloudflare Rocket Loader defers all JavaScript including Gravity Forms dependencies, causing race conditions on resumed Save and Continue forms - disabling it per page resolves the issue.
- Power Automate Microsoft Forms trigger avoids the need for a premium HTTP license when collecting external user emails.
- For SharePoint Graph API lookups, the correct hostname is pcma2.sharepoint.com (not pcma.sharepoint.com), and folder paths within a drive should not include the drive name as a prefix.
- Excel 365 native checkboxes cannot be reliably generated programmatically via openpyxl - deliver empty cells with a clear instruction banner directing the user to Insert > Checkbox in the ribbon.
- PCMA's brand presence in the events industry is meaningful negotiating leverage with AMS and technology vendors.

## Approach and Patterns

- Prefers clarifying questions before major builds - wants complete understanding established before implementation begins.
- Wants ready-to-use output (copy, documents, code) rather than general guidance or frameworks to adapt himself.
- Iterates quickly - provides direct corrective feedback mid-session and moves on without extensive back-and-forth once a fix is confirmed.
- Board and executive materials favor brevity, active voice, and concrete status framing (what has been done, what is in progress, what is next).
- Documents are polished enough for executive presentation but detailed enough for technical vendor or implementation use.
- Flags unknowns explicitly on diagrams and documents rather than omitting uncertain information.
- Thinks end-to-end about solutions - has corrected Claude for suggesting steps without considering downstream implications (CORS, file naming, data URL targets).
- Prefers Word documents using PCMA brand styles (Heading 1/2/3, List Paragraph, Normal) for deliverables intended for SharePoint or executive distribution.

## Brand and Style

Reference document: `../pcma brand definitions.pdf` (Brand Guidelines v1.2, June 2025). Consult this file for logo usage, graphic language, photography style, and application examples.

**Tagline:** "Leading minds. Leading change."

**Primary palette:**
- Dark Mauve: #590044 (R89 G0 B68) - primary background color, emphasis
- Mid Mauve: #6F004F (R111 G0 B79)
- Light Mauve: #85165A (R133 G22 B90)
- Light Teal: #60E5CF (R96 G229 B207) - accent color
- Mid Neutral: #D2C3B9 (R210 G195 B185)
- Light Neutral: #F1E7E3 (R241 G231 B227) - page background
- White: #FFFFFF

**Secondary palette (for charts and graphs, not covers):**
- Light Blue: #83E7FF, Mid Blue: #009FDA, Dark Blue: #164C76
- Light Lilac: #C3B4FF, Mid Lilac: #8D7DE4, Dark Lilac: #4D3971
- Light Coral: #FF9999, Mid Coral: #F25D7A, Dark Coral: #A5254A
- Light Orange: #FFCB57, Mid Orange: #D18309, Dark Orange: #984929
- Light Green: #ADEC73, Mid Green: #44A753, Dark Green: #205840
- Mid Teal: #00A4A4, Dark Teal: #1D5F6F

**Chart color order (for maximum contrast):** Dark Mauve, Light Teal, Light Coral, Light Blue, Light Orange, Light Lilac, Light Green, Light Mauve, Dark Teal, Dark Coral, Mid Blue, Mid Orange, Dark Lilac, Dark Green, Mid Teal, Mid Coral, Dark Blue, Dark Orange, Mid Lilac, Mid Green. Pie charts use donut style with white rules between segments.

**Typography:**
- Brand typeface: Gabarito (Google Font, Bold for headlines, Regular for standfirsts)
- System font (Microsoft apps, web): Aptos (ExtraBold for headlines/sub-headings, SemiBold for standfirsts, Light for body copy, Regular for small copy)
- Fallback: Segoe UI

**Color usage principles:**
- Parent brand: emphasis on Dark Mauve, supported by neutrals and white space, Light Teal as accent. Secondary palette only in charts/graphs, never on covers.
- Sub brands: flip emphasis so neutrals, Light Teal, or White lead, with mauve in support.
- On dark backgrounds: use White, Light Neutral, Mid Neutral, or Light Teal for text.
- On light backgrounds: use Dark Mauve, Mid Mauve, or Light Mauve for text.

**Writing style:**
- Never use em dashes, en dashes, or any non-ASCII or hidden characters in responses. Use commas or regular hyphens (-) only if punctuation is needed.

## Tools and Resources

- Microsoft 365: SharePoint (pcma2.sharepoint.com), Power Automate (make.powerautomate.com), Outlook, Teams, Excel, Word, PowerPoint, Entra/Azure AD
- Salesforce with Fonteva (AMS, targeted for replacement), AdVendio (ad management)
- GitHub Pages (stevenrieger.github.io/PCMA-Executive-Dashboard) for dashboard hosting
- Claude M365 add-in (org-wide deployment via M365 Admin Center)
- RainFocus (event execution, two-way Salesforce sync), Stripe (e-commerce), Business Central (ERP), Pardot (marketing), WordPress (chapter sites), LearnUp (LMS)
- PowerShell with Microsoft Graph API (TenantId: 758ab235-b480-4c59-8e56-30144f4893ce) for SharePoint and Entra administration
- Python with openpyxl for Excel file generation; PptxGenJS for PowerPoint generation
