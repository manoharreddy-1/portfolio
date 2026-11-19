import os
import re
import html
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

base_dir = r'c:\Users\manoh\OneDrive\Desktop\manohar\portfolio'
assets_dir = os.path.join(base_dir, 'assets', 'case-studies')
fig_dir = os.path.join(assets_dir, 'figures')
docx_out = os.path.join(assets_dir, 'Product-Management-AI-Case-Studies.docx')
html_out = os.path.join(base_dir, 'case-studies.html')

fig_map = {
    'Figure 1.1': 'fitsure-flow.png',
    'Figure 1.2': 'fitsure-architecture.png',
    'Figure 2.1': 'learnpath-flow.png',
    'Figure 2.2': 'learnpath-architecture.png',
    'Figure 3.1': 'resolveai-flow.png',
    'Figure 3.2': 'resolveai-architecture.png',
    'Figure 4.1': 'ruralroute-flow.png',
    'Figure 4.2': 'ruralroute-architecture.png'
}

with open(os.path.join(base_dir, 'scripts', 'new_prompt_text.txt'), 'r', encoding='utf-8') as f:
    raw_text = f.read()

# Fix the truncated Senior PM Review table in raw_text if needed
senior_pm_full = '''Senior PM Review of the Portfolio
This is my own review, written to the standard an interviewer or leadership panel would apply. I haven’t given myself top marks. The biggest limit is that the portfolio has no primary user research and no company data, so its numbers are structured assumptions, not findings.
| Dimension   | Score / 10   | Reason   |
| Problem understanding   | 8   | Root-cause tables and explicit “problem ≠ solution”; limited by lack of primary research.   |
| User empathy   | 6   | Segments and JTBD are sensible, but every need is a hypothesis; no real interviews.   |
| Product strategy   | 8   | Gap analysis against existing products; sequencing; explicit non-goals.   |
| AI / technical understanding   | 8   | Right tool per job (rules, trees, hybrid retrieval, constrained LLM); detail is design-level, not validated by prototype.   |
| Business thinking   | 7   | Transparent ROI models and break-evens; all inputs illustrative; no real cost data.   |
| Metrics   | 8   | Gameable metrics replaced (verified resolution, mastery); guardrails defined; baselines still to be measured.   |
| Experimentation   | 8   | Unit of randomisation chosen per feature; power calculations; stop rules; return-lag and clustering handled.   |
| Cost awareness   | 8   | Batch over per-request LLMs; integrate existing optimiser; buy-vs-build threshold; token prices are assumptions.   |
| Responsible AI   | 8   | Concrete controls, “never” lists, DPDP, prompt-injection and liability evidence; legal conclusions deferred to counsel.   |
| Prioritisation   | 7   | Numeric RICE with dependency caveats; inputs are my estimates.   |
| Real-world feasibility   | 7   | Timelines lengthened for lags; depends on data access (seller data, partner operator, Google-internal data).   |
| Overall PM quality   | 7.5   | Strong judgement and honesty about evidence; would rise to 8.5+ with real interviews and one prototype backtest.   |

Top 10 weaknesses in the original document
Targets without baselines or evidence (10–15% fewer returns, +15% goal completion, 30–40% autonomous resolution).
Presented existing capabilities as new (YouTube Courses, Ask and quizzes; Google offline maps, two-wheeler mode, landmark descriptors, route optimisation).
Gameable or undefined metrics: “autonomous resolution”, “learning-goal completion”, “drift”.
RICE shown as High/Medium labels with no numbers and no dependency logic.
MVPs too broad (Case 3 combined search, RAG, copilot and actions; Case 4 built five features at once).
Timelines ignored lags: return windows, shadow mode, monsoon seasonality.
Thin technical depth: no label design, evaluation sets, thresholds, retrieval detail or tenant isolation.
Responsible AI as short bullet lists; missing DPDP for minors, prompt injection, liability, location-privacy risks.
No ROI or cost model, and no break-even logic.
Evidence quality: a vendor blog as support; unsourced statements; experiments without unit, power or decision rules.

Top 10 improvements made
Added evidence from official and research sources and labelled every assumption, target and hypothesis.
Re-scoped each case to the real gap after checking what exists today.
Replaced gameable metrics with verified ones, and added guardrails.
Added weighted decision matrices comparing six options each, with explicit reasons for rejection.
Converted RICE to numbers and stated its limits.
Designed experiments with randomisation unit, power, duration, outcomes and failure plans.
Added accuracy strategies: label audits, calibration, abstention, repeated-trial reliability, shadow mode.
Built transparent ROI models with sensitivity and break-even.
Strengthened responsible AI with “never” lists and regulatory, security and safety controls.
Added architecture and flow diagrams plus cost-optimisation reasoning.

Top 10 strongest PM decisions in the final version
FitSure: data first (measurements, flags), ML second, computer vision and AR deliberately deferred.
FitSure: lowering the target to match the best public benchmark rather than the original’s optimism.
LearnPath: define success as verified mastery, with an encouragement design that handles self-selection.
LearnPath: no AI answer-giving and no under-18 personalisation until learning-harm and legal risks are resolved.
ResolveAI: verified resolution instead of deflection, and a copilot-first autonomy ladder.
ResolveAI: an explicit volume threshold below which to buy rather than build.
RuralRoute: integrate the existing optimiser and DIGIPIN rather than rebuild them.
RuralRoute: choose the randomisation unit per feature (per order for cards, geo-staggered for shared road state).
Every case: explicit continue-and-stop evidence, so the idea can fail cleanly.
Every case: cost levers that keep expensive models offline, batched or optional.

Top 5 areas an interviewer may challenge, and strong answers
| Challenge   | Strong answer   |
| “Your targets and ROI are made up.”   | They are labelled assumptions, built to show decision logic: baseline first, sensitivity tables, and break-even thresholds (for example FitSure needs about 4.6% for a 24-month payback). I anchored targets to the best public benchmark (Zalando’s 8%) and would replace every input with company data before committing.   |
| “Doesn’t this already exist?”   | Yes in part, which is why each case starts with a gap analysis. YouTube has Courses and Ask; Google has offline maps, descriptors and an optimiser. The proposals add only what is missing: cross-creator mastery-verified paths; rural drop-point learning; and so on.   |
| “How do you measure learning or ‘resolved’?”   | With delayed retrieval checks and an encouragement design for learning; with a 7-day no-recontact rule plus audit sampling for resolution. Both avoid counting activity or deflection as success.   |
| “Why not go fully autonomous, or add vision and AR?”   | Because risk and cost rise faster than evidence. The autonomy ladder, reliability evidence, liability and the break-even maths show copilot-first pays and lowers risk; body scanning adds privacy and cost without proven incremental value.   |
| “Where is your user research, and what about privacy?”   | There is none yet, and the document says so: needs are hypotheses with a concrete research plan. Privacy is designed in: optional inputs, task-window access, no minors’ profiling, DPDP readiness, and “never” lists. I would run the interviews before any build decision.   |
'''

# Check if Senior PM Review needs full replacement
if '| Prioritisation   | 7   | Numeric RICE with dependency caveats' in raw_text:
    idx_p = raw_text.find('| Prioritisation   | 7   | Numeric RICE with dependency caveats')
    idx_end = raw_text.find('Appendices', idx_p)
    raw_text = raw_text[:idx_p] + senior_pm_full + '\n' + raw_text[idx_end:]

lines = raw_text.splitlines()

# Parse into blocks: ('h1', text), ('h2', text), ('h3', text), ('p', text), ('table', [rows]), ('figure', (caption, img_file)), ('code', [lines])
blocks = []
i = 0
n = len(lines)

while i < n:
    line = lines[i].strip()
    if not line:
        i += 1
        continue
    
    # Check for Markdown table
    if line.startswith('|') and line.endswith('|'):
        table_rows = []
        while i < n and lines[i].strip().startswith('|') and lines[i].strip().endswith('|'):
            row_str = lines[i].strip()
            # Split cells
            cells = [c.strip() for c in row_str.strip('|').split('|')]
            # Skip separator row if any (e.g. | --- | --- |)
            if not all(re.match(r'^:?-+:?$', c) for c in cells):
                table_rows.append(cells)
            i += 1
        blocks.append(('table', table_rows))
        continue
    
    # Check for Figure
    if line.startswith('Figure ') and any(f in line for f in fig_map):
        matched_fig = None
        for fkey in fig_map:
            if fkey in line:
                matched_fig = fkey
                break
        img_name = fig_map[matched_fig] if matched_fig else None
        blocks.append(('figure', (line, img_name)))
        i += 1
        continue
    
    # Check for Mermaid code block
    if line.startswith('%% Case ') or line.startswith('flowchart '):
        code_lines = []
        while i < n and (lines[i].strip().startswith('%%') or lines[i].strip().startswith('flowchart') or '-->' in lines[i] or '--' in lines[i]):
            code_lines.append(lines[i].strip())
            i += 1
        blocks.append(('code', code_lines))
        continue

    # Check for Case Study header
    if line.startswith('Case Study 1:') or line.startswith('Case Study 2:') or line.startswith('Case Study 3:') or line.startswith('Case Study 4:'):
        blocks.append(('case_header', line))
        i += 1
        continue

    # Check for major sections
    if line in ['Product Management + AI', 'Case Study Portfolio']:
        blocks.append(('doc_title', line))
        i += 1
        continue
    
    if line in ['Contents', 'Senior PM Review of the Portfolio', 'Appendices', 'Portfolio at a glance', 'What the research changed', 'Diagnosis of the original document']:
        blocks.append(('h1', line))
        i += 1
        continue

    if line in ['Top 10 weaknesses in the original document', 'Top 10 improvements made', 'Top 10 strongest PM decisions in the final version', 'Top 5 areas an interviewer may challenge, and strong answers']:
        blocks.append(('h2', line))
        i += 1
        continue
    
    if line.startswith('A. Product-flow sources') or line.startswith('B. Research-paper notes') or line.startswith('C. References') or line.startswith('D. Final quality check'):
        blocks.append(('h2', line))
        i += 1
        continue

    # Numbered section heading (e.g. 1. Executive Summary)
    m = re.match(r'^(\d+)\.\s+(.*)$', line)
    if m:
        blocks.append(('h2_numbered', (m.group(1), m.group(2))))
        i += 1
        continue

    blocks.append(('p', line))
    i += 1

print(f'Total parsed blocks: {len(blocks)}')

# ==================== BUILD HTML ====================
html_parts = []

def format_text(txt):
    txt = html.escape(txt)
    txt = re.sub(r'\bAssumption\b', '<span class="cs-badge cs-badge--warn">Assumption</span>', txt)
    txt = re.sub(r'\bProposed MVP target\b', '<span class="cs-badge cs-badge--info">Proposed MVP Target</span>', txt)
    txt = re.sub(r'\bHypothesis to validate through user research\b', '<span class="cs-badge cs-badge--purple">Hypothesis</span>', txt)
    txt = re.sub(r'\bEvidence not found / assumption to validate\b', '<span class="cs-badge cs-badge--danger">Evidence Not Found</span>', txt)
    txt = re.sub(r'\bEvidence not found\b', '<span class="cs-badge cs-badge--danger">Evidence Not Found</span>', txt)
    return txt

for b_type, b_data in blocks:
    if b_type == 'doc_title':
        continue
    elif b_type == 'case_header':
        cs_id = 'case-' + b_data.split()[2].replace(':', '')
        tag_color = 'blue' if '1' in b_data else ('purple' if '2' in b_data else ('emerald' if '3' in b_data else 'amber'))
        cat_name = 'E-Commerce & AI' if '1' in b_data else ('EdTech & AI' if '2' in b_data else ('Enterprise SaaS & RAG' if '3' in b_data else 'Geospatial & Navigation'))
        cs_num = b_data.split()[2].replace(':', '')
        html_parts.append(f'''
        <div class="cs-case-anchor" id="{cs_id}"></div>
        <div class="cs-case-header">
          <div class="cs-case-badge-row">
            <span class="cs-tag cs-tag--{tag_color}">{cat_name}</span>
            <span class="cs-tag cs-tag--glow">Case Study 0{cs_num}</span>
          </div>
          <h2 class="cs-case-title">{format_text(b_data)}</h2>
        </div>''')
    elif b_type == 'h1':
        anchor_id = 'review' if 'Senior PM Review' in b_data else ('appendices' if 'Appendices' in b_data else '')
        if anchor_id:
            html_parts.append(f'<div class="cs-case-anchor" id="{anchor_id}"></div>')
        html_parts.append(f'<h2 class="cs-heading-1">{format_text(b_data)}</h2>')
    elif b_type == 'h2':
        html_parts.append(f'<h3 class="cs-heading-2">{format_text(b_data)}</h3>')
    elif b_type == 'h2_numbered':
        num, title = b_data
        html_parts.append(f'''
        <div class="cs-section-header">
          <span class="cs-section-num">{num}</span>
          <h3 class="cs-heading-2">{format_text(title)}</h3>
        </div>''')
    elif b_type == 'figure':
        caption, img_name = b_data
        if img_name:
            img_path = f'assets/case-studies/figures/{img_name}'
            html_parts.append(f'''
            <figure class="cs-figure">
              <div class="cs-figure__frame">
                <img src="{img_path}" alt="{html.escape(caption)}" class="cs-figure__img" loading="lazy" />
              </div>
              <figcaption class="cs-figure__caption"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="vertical-align:middle;margin-right:6px;"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>{html.escape(caption)}</figcaption>
            </figure>''')
        else:
            html_parts.append(f'<p class="cs-figure-label">{format_text(caption)}</p>')
    elif b_type == 'table':
        rows = b_data
        if not rows:
            continue
        hdr = rows[0]
        tbody_rows = rows[1:]
        hdr_html = ''.join(f'<th>{format_text(c)}</th>' for c in hdr)
        body_html = []
        for r in tbody_rows:
            row_cells = ''.join(f'<td>{format_text(c)}</td>' for c in r)
            body_html.append(f'<tr>{row_cells}</tr>')
        
        table_html = f'''
        <div class="cs-table-wrap">
          <table class="cs-table">
            <thead><tr>{hdr_html}</tr></thead>
            <tbody>{''.join(body_html)}</tbody>
          </table>
        </div>'''
        html_parts.append(table_html)
    elif b_type == 'code':
        code_str = html.escape('\n'.join(b_data))
        html_parts.append(f'<div class="cs-code-wrap"><pre><code>{code_str}</code></pre></div>')
    elif b_type == 'p':
        txt = b_data
        if txt.startswith('Product:') and 'Boundary:' in txt:
            html_parts.append(f'<div class="cs-meta-box"><p>{format_text(txt)}</p></div>')
        elif txt.startswith('North Star Metric:') or txt.startswith('North Star:'):
            html_parts.append(f'''
            <div class="cs-callout cs-callout--star">
              <div class="cs-callout__icon">⭐</div>
              <div class="cs-callout__content">
                <strong>North Star Metric</strong>
                <p>{format_text(txt.split(":", 1)[1].strip() if ":" in txt else txt)}</p>
              </div>
            </div>''')
        elif txt.startswith('Problem statement.') or txt.startswith('Problem Statement:'):
            html_parts.append(f'''
            <div class="cs-callout cs-callout--problem">
              <div class="cs-callout__icon">🎯</div>
              <div class="cs-callout__content">
                <strong>Problem Statement</strong>
                <p>{format_text(txt.split(":", 1)[1].strip() if ":" in txt else txt)}</p>
              </div>
            </div>''')
        elif txt.startswith('Central PM question:'):
            html_parts.append(f'''
            <div class="cs-callout cs-callout--question">
              <div class="cs-callout__icon">💡</div>
              <div class="cs-callout__content">
                <strong>Central PM Question</strong>
                <p>{format_text(txt.split(":", 1)[1].strip() if ":" in txt else txt)}</p>
              </div>
            </div>''')
        else:
            html_parts.append(f'<p class="cs-p">{format_text(txt)}</p>')

content_body = '\n'.join(html_parts)

full_html = f'''<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="description" content="Product Management & AI Case Studies Portfolio by Manohar Reddy Bommareddy. Four evidence-checked case studies across E-commerce, EdTech, Enterprise SaaS, and Geospatial Navigation." />
  <title>Product Management + AI Case Studies | Manohar Reddy Bommareddy</title>
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet" />
  <link rel="stylesheet" href="css/style.css" />
  <link rel="stylesheet" href="css/case-studies.css" />
  <script>
    (function () {{
      const saved = localStorage.getItem('theme');
      const prefersLight = window.matchMedia('(prefers-color-scheme: light)').matches;
      document.documentElement.setAttribute('data-theme', saved || (prefersLight ? 'light' : 'dark'));
    }})();
  </script>
</head>
<body class="case-studies-page">
  <div class="bg-grid"></div>
  <div class="bg-glow bg-glow--1"></div>
  <div class="bg-glow bg-glow--2"></div>

  <!-- Top Navigation -->
  <nav class="navbar" id="navbar">
    <div class="container navbar__inner">
      <a href="index.html" class="navbar__logo">MANOHAR<span>.</span></a>
      <button class="navbar__toggle" id="navToggle" aria-label="Toggle menu">
        <span></span><span></span><span></span>
      </button>
      <ul class="navbar__links" id="navLinks">
        <li><a href="index.html#home">Home</a></li>
        <li><a href="index.html#projects">Projects</a></li>
        <li><a href="index.html#case-studies">PM Overview</a></li>
        <li><a href="#case-1">1. FitSure</a></li>
        <li><a href="#case-2">2. LearnPath</a></li>
        <li><a href="#case-3">3. ResolveAI</a></li>
        <li><a href="#case-4">4. RuralRoute</a></li>
        <li><a href="#review">PM Review</a></li>
        <li><a href="index.html#certifications">Certificates</a></li>
      </ul>
      <div class="navbar__actions">
        <a href="assets/case-studies/Product-Management-AI-Case-Studies.docx" download class="btn btn--sm btn--primary nav-download-btn">
          <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" style="vertical-align:middle;margin-right:4px;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>DOCX
        </a>
        <button class="theme-toggle" id="themeToggle" aria-label="Toggle light/dark theme">
          <svg class="theme-toggle__icon theme-toggle__icon--sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42"/></svg>
          <svg class="theme-toggle__icon theme-toggle__icon--moon" viewBox="0 0 24 24" fill="currentColor"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>
        </button>
      </div>
    </div>
  </nav>

  <!-- Page Header -->
  <header class="cs-hero">
    <div class="container">
      <div class="cs-hero__crumbs">
        <a href="index.html">Portfolio</a> &nbsp;/&nbsp; <span>Product Management</span>
      </div>
      <div class="cs-hero__badge-row">
        <span class="cs-badge cs-badge--pm">Product Management + AI</span>
        <span class="cs-badge cs-badge--upgrad">upGrad Certified</span>
        <span class="cs-badge cs-badge--oct">Updated October 2026</span>
      </div>
      <h1 class="cs-hero__title">
        Product Management + AI<br />
        <span class="gradient-text">Case Study Portfolio</span>
      </h1>
      <p class="cs-hero__desc">
        Four rewritten, evidence-checked case studies engineered with rigorous problem boundary definitions, 
        Jobs-to-be-Done segmentation, quantitative RICE prioritisation, machine learning architectures, and 
        break-even ROI economic models.
      </p>

      <div class="cs-hero__stats-grid">
        <div class="cs-stat-card">
          <div class="cs-stat-card__val">4</div>
          <div class="cs-stat-card__lbl">Deep-Dive Case Studies</div>
          <div class="cs-stat-card__sub">E-commerce, EdTech, SaaS, Maps</div>
        </div>
        <div class="cs-stat-card">
          <div class="cs-stat-card__val">8</div>
          <div class="cs-stat-card__lbl">Architectures & Flows</div>
          <div class="cs-stat-card__sub">End-to-End Product Blueprints</div>
        </div>
        <div class="cs-stat-card">
          <div class="cs-stat-card__val">24</div>
          <div class="cs-stat-card__lbl">Sections Per Study</div>
          <div class="cs-stat-card__sub">Rigorous Product Lifecycle Rubric</div>
        </div>
        <div class="cs-stat-card">
          <div class="cs-stat-card__val">100%</div>
          <div class="cs-stat-card__lbl">Evidence-Grounded</div>
          <div class="cs-stat-card__sub">Zalando, YouTube, Google Maps data</div>
        </div>
      </div>

      <div class="cs-hero__actions">
        <a href="#case-1" class="btn btn--primary">Read Case Studies</a>
        <a href="assets/case-studies/Product-Management-AI-Case-Studies.docx" download class="btn btn--outline">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" style="margin-right:6px;"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>Download Full Portfolio (.docx)
        </a>
        <a href="index.html#certifications" class="btn btn--outline">View upGrad Certificate</a>
      </div>
    </div>
  </header>

  <!-- Main Layout with Sidebar -->
  <main class="cs-main">
    <div class="container cs-layout">
      <!-- Sticky Sidebar -->
      <aside class="cs-sidebar" id="csSidebar">
        <div class="cs-sidebar__inner">
          <div class="cs-sidebar__title">Portfolio Index</div>
          <nav class="cs-sidebar__nav">
            <a href="#intro" class="cs-sidebar__link active">Overview & At a Glance</a>
            <a href="#case-1" class="cs-sidebar__link">
              <span class="cs-sidebar__dot cs-sidebar__dot--1"></span>
              1. FitSure (E-Commerce)
            </a>
            <a href="#case-2" class="cs-sidebar__link">
              <span class="cs-sidebar__dot cs-sidebar__dot--2"></span>
              2. LearnPath (YouTube)
            </a>
            <a href="#case-3" class="cs-sidebar__link">
              <span class="cs-sidebar__dot cs-sidebar__dot--3"></span>
              3. ResolveAI (SaaS Support)
            </a>
            <a href="#case-4" class="cs-sidebar__link">
              <span class="cs-sidebar__dot cs-sidebar__dot--4"></span>
              4. RuralRoute (Google Maps)
            </a>
            <a href="#review" class="cs-sidebar__link">
              <span class="cs-sidebar__dot cs-sidebar__dot--5"></span>
              Senior PM Review
            </a>
            <a href="#appendices" class="cs-sidebar__link">
              <span class="cs-sidebar__dot cs-sidebar__dot--6"></span>
              Appendices & Sources
            </a>
          </nav>
          
          <div class="cs-sidebar__card">
            <h4>upGrad Certified</h4>
            <p>Introduction to Product Management</p>
            <div class="cs-sidebar__meta">ID: e925c5a9-91ee...</div>
            <a href="assets/certificates/upgrad-product-management.pdf" target="_blank" class="cs-sidebar__cert-link">View Certificate &rarr;</a>
          </div>
        </div>
      </aside>

      <!-- Content Area -->
      <article class="cs-content">
        <div class="cs-case-anchor" id="intro"></div>
        {content_body}
      </article>
    </div>
  </main>

  <footer class="footer">
    <div class="container footer__inner">
      <p>&copy; 2026 Manohar Reddy Bommareddy. All rights reserved.</p>
      <p class="footer__mono">Built with passion for Product Management, AI &amp; Code</p>
    </div>
  </footer>

  <script src="js/main.js"></script>
  <script src="js/case-studies.js"></script>
</body>
</html>'''

with open(html_out, 'w', encoding='utf-8') as f:
    f.write(full_html)

print('Updated case-studies.html successfully!')

# ==================== BUILD DOCX ====================
doc = docx.Document()

# Set page margins to 0.75 in
for section in doc.sections:
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

# Helper function to style table headers
def style_table(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))
        if i == 0:
            trPr.append(parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w'))))
            for cell in row.cells:
                shading = parse_xml(r'<w:shd {} w:fill="0F2942"/>'.format(nsdecls('w')))
                cell._tc.get_or_add_tcPr().append(shading)
                for p in cell.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(255, 255, 255)
                        run.font.size = Pt(8.5)
        else:
            fill_col = "F8FAFC" if i % 2 == 1 else "FFFFFF"
            for cell in row.cells:
                shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), fill_col))
                cell._tc.get_or_add_tcPr().append(shading)
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(8.5)

for b_type, b_data in blocks:
    if b_type == 'doc_title':
        p = doc.add_paragraph()
        run = p.add_run(b_data)
        run.font.bold = True
        run.font.size = Pt(22)
        run.font.color.rgb = RGBColor(12, 53, 90)
    elif b_type == 'case_header':
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(b_data)
        run.font.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = RGBColor(15, 76, 129)
    elif b_type == 'h1':
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(b_data)
        run.font.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(12, 53, 90)
    elif b_type == 'h2':
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(b_data)
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(29, 78, 216)
    elif b_type == 'h2_numbered':
        num, title = b_data
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(f'{num}. {title}')
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(29, 78, 216)
    elif b_type == 'figure':
        caption, img_name = b_data
        if img_name and os.path.exists(os.path.join(fig_dir, img_name)):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(2)
            run_img = p_img.add_run()
            run_img.add_picture(os.path.join(fig_dir, img_name), width=Inches(6.2))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        run_cap = p_cap.add_run(caption)
        run_cap.font.italic = True
        run_cap.font.size = Pt(8.5)
        run_cap.font.color.rgb = RGBColor(100, 116, 139)
    elif b_type == 'table':
        rows = b_data
        if not rows:
            continue
        ncols = max(len(r) for r in rows)
        tbl = doc.add_table(rows=len(rows), cols=ncols)
        tbl.autofit = True
        for r_idx, r in enumerate(rows):
            for c_idx in range(ncols):
                val = r[c_idx] if c_idx < len(r) else ''
                cell = tbl.cell(r_idx, c_idx)
                cell.text = val
        style_table(tbl)
        doc.add_paragraph().paragraph_format.space_after = Pt(4)
    elif b_type == 'code':
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run('\n'.join(b_data))
        run.font.name = 'Consolas'
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(30, 41, 59)
    elif b_type == 'p':
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(b_data)
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(15, 23, 42)

doc.save(docx_out)
print('Updated Product-Management-AI-Case-Studies.docx successfully!')
