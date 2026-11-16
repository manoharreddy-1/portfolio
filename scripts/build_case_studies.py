import docx
import xml.etree.ElementTree as ET
import os
import html
import re

doc_path = r'c:\Users\manoh\OneDrive\Desktop\manohar\portfolio\assets\case-studies\Product-Management-AI-Case-Studies.docx'
doc = docx.Document(doc_path)
rels = doc.part.rels

img_map = {
    'bc09f25be9a9115399c4fc644f3eeedafddb1bd1.png': 'fitsure-flow.png',
    'c0000ee83c15b815adb66a7fd619bbb44114a83e.png': 'fitsure-architecture.png',
    'f10e9f062d2ff03b8232be1ba865271ac45aa449.png': 'learnpath-flow.png',
    '0de3144660f1244c9482c978121a7eb875a4dca8.png': 'learnpath-architecture.png',
    '7d7312f403820c04d6e444760e5d9f54ea0f2b84.png': 'resolveai-flow.png',
    'a7426f319b7edc9786f9e4f09dd750581ce6424b.png': 'resolveai-architecture.png',
    'b6acd22161b6a96416c862be3da89931aab476d5.png': 'ruralroute-flow.png',
    '4dd72553c5d5c78693f0c43307158fcec76b492d.png': 'ruralroute-architecture.png'
}

def format_text(txt):
    txt = html.escape(txt)
    txt = re.sub(r'\bAssumption\b', '<span class="cs-badge cs-badge--warn">Assumption</span>', txt)
    txt = re.sub(r'\bProposed MVP target\b', '<span class="cs-badge cs-badge--info">Proposed MVP Target</span>', txt)
    txt = re.sub(r'\bHypothesis to validate through user research\b', '<span class="cs-badge cs-badge--purple">Hypothesis</span>', txt)
    txt = re.sub(r'\bEvidence not found / assumption to validate\b', '<span class="cs-badge cs-badge--danger">Evidence Not Found</span>', txt)
    txt = re.sub(r'\bEvidence not found\b', '<span class="cs-badge cs-badge--danger">Evidence Not Found</span>', txt)
    return txt

html_parts = []
body_elements = doc.element.body

in_case_1 = False
in_case_2 = False
in_case_3 = False
in_case_4 = False
in_review = False
in_appendices = False

for elem in body_elements:
    if elem.tag.endswith('p'):
        img_src = None
        for descendant in elem.iter():
            if descendant.tag.endswith('blip'):
                embed = descendant.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed and embed in rels:
                    target = os.path.basename(rels[embed].target_ref)
                    if target in img_map:
                        img_src = 'assets/case-studies/figures/' + img_map[target]
                        break
        
        p = docx.text.paragraph.Paragraph(elem, doc)
        text = p.text.strip()
        style_name = p.style.name if p.style else ''

        if img_src:
            caption_text = text if text else "Case Study Architecture & Workflow"
            html_parts.append(f'''
            <figure class="cs-figure">
              <div class="cs-figure__frame">
                <img src="{img_src}" alt="{html.escape(caption_text)}" class="cs-figure__img" loading="lazy" />
              </div>
              <figcaption class="cs-figure__caption"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="vertical-align:middle;margin-right:6px;"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>{html.escape(caption_text)}</figcaption>
            </figure>''')
            continue

        if not text and not img_src:
            continue

        if style_name == 'Title' or text == 'Product Management + AI' or text == 'Case Study Portfolio':
            continue

        if text.startswith('Case Study 1: Reduce Returns'):
            html_parts.append('<div class="cs-case-anchor" id="case-1"></div>')
            html_parts.append(f'''
            <div class="cs-case-header">
              <div class="cs-case-badge-row">
                <span class="cs-tag cs-tag--blue">E-Commerce & AI</span>
                <span class="cs-tag cs-tag--glow">Case Study 01</span>
              </div>
              <h2 class="cs-case-title">{format_text(text)}</h2>
            </div>''')
        elif text.startswith('Case Study 2: Improve YouTube'):
            html_parts.append('<div class="cs-case-anchor" id="case-2"></div>')
            html_parts.append(f'''
            <div class="cs-case-header">
              <div class="cs-case-badge-row">
                <span class="cs-tag cs-tag--purple">EdTech & AI</span>
                <span class="cs-tag cs-tag--glow">Case Study 02</span>
              </div>
              <h2 class="cs-case-title">{format_text(text)}</h2>
            </div>''')
        elif text.startswith('Case Study 3: AI Support Agent'):
            html_parts.append('<div class="cs-case-anchor" id="case-3"></div>')
            html_parts.append(f'''
            <div class="cs-case-header">
              <div class="cs-case-badge-row">
                <span class="cs-tag cs-tag--emerald">Enterprise SaaS & RAG</span>
                <span class="cs-tag cs-tag--glow">Case Study 03</span>
              </div>
              <h2 class="cs-case-title">{format_text(text)}</h2>
            </div>''')
        elif text.startswith('Case Study 4: Improve Google Maps'):
            html_parts.append('<div class="cs-case-anchor" id="case-4"></div>')
            html_parts.append(f'''
            <div class="cs-case-header">
              <div class="cs-case-badge-row">
                <span class="cs-tag cs-tag--amber">Geospatial & Navigation</span>
                <span class="cs-tag cs-tag--glow">Case Study 04</span>
              </div>
              <h2 class="cs-case-title">{format_text(text)}</h2>
            </div>''')
        elif text == 'Senior PM Review of the Portfolio':
            html_parts.append('<div class="cs-case-anchor" id="review"></div>')
            html_parts.append(f'''
            <div class="cs-case-header">
              <div class="cs-case-badge-row">
                <span class="cs-tag cs-tag--rose">Evaluation & Rubric</span>
                <span class="cs-tag cs-tag--glow">Senior PM Review</span>
              </div>
              <h2 class="cs-case-title">{format_text(text)}</h2>
            </div>''')
        elif text.startswith('Appendices'):
            html_parts.append('<div class="cs-case-anchor" id="appendices"></div>')
            html_parts.append(f'''
            <div class="cs-case-header">
              <div class="cs-case-badge-row">
                <span class="cs-tag cs-tag--cyan">Technical Details</span>
                <span class="cs-tag cs-tag--glow">Appendices & References</span>
              </div>
              <h2 class="cs-case-title">{format_text(text)}</h2>
            </div>''')
        elif style_name == 'Heading 1':
            html_parts.append(f'<h2 class="cs-heading-1">{format_text(text)}</h2>')
        elif style_name == 'Heading 2':
            # Section titles like "1. Executive Summary"
            m = re.match(r'^(\d+)\.\s+(.*)$', text)
            if m:
                sec_num, sec_title = m.groups()
                html_parts.append(f'''
                <div class="cs-section-header">
                  <span class="cs-section-num">{sec_num}</span>
                  <h3 class="cs-heading-2">{format_text(sec_title)}</h3>
                </div>''')
            else:
                html_parts.append(f'<h3 class="cs-heading-2">{format_text(text)}</h3>')
        elif style_name == 'Heading 3':
            html_parts.append(f'<h4 class="cs-heading-3">{format_text(text)}</h4>')
        elif text.startswith('Product:') and 'Boundary:' in text:
            html_parts.append(f'<div class="cs-meta-box"><p>{format_text(text)}</p></div>')
        elif text.startswith('North Star Metric:') or text.startswith('North Star:'):
            html_parts.append(f'''
            <div class="cs-callout cs-callout--star">
              <div class="cs-callout__icon">⭐</div>
              <div class="cs-callout__content">
                <strong>North Star Metric</strong>
                <p>{format_text(text.split(":", 1)[1].strip() if ":" in text else text)}</p>
              </div>
            </div>''')
        elif text.startswith('Problem statement.') or text.startswith('Problem Statement:'):
            html_parts.append(f'''
            <div class="cs-callout cs-callout--problem">
              <div class="cs-callout__icon">🎯</div>
              <div class="cs-callout__content">
                <strong>Problem Statement</strong>
                <p>{format_text(text.split(":", 1)[1].strip() if ":" in text else text)}</p>
              </div>
            </div>''')
        elif text.startswith('Central PM question:'):
            html_parts.append(f'''
            <div class="cs-callout cs-callout--question">
              <div class="cs-callout__icon">💡</div>
              <div class="cs-callout__content">
                <strong>Central PM Question</strong>
                <p>{format_text(text.split(":", 1)[1].strip() if ":" in text else text)}</p>
              </div>
            </div>''')
        elif text.startswith('Figure '):
            html_parts.append(f'<p class="cs-figure-label">{format_text(text)}</p>')
        else:
            html_parts.append(f'<p class="cs-p">{format_text(text)}</p>')

    elif elem.tag.endswith('tbl'):
        tbl = docx.table.Table(elem, doc)
        rows_html = []
        is_first = True
        for row in tbl.rows:
            # deduplicate merged cells
            cells = []
            seen_texts = set()
            for c in row.cells:
                t = c.text.strip()
                cells.append(t)
            
            tag = 'th' if is_first else 'td'
            cell_html = ''.join(f'<{tag}>{format_text(c)}</{tag}>' for c in cells)
            rows_html.append(f'<tr>{cell_html}</tr>')
            is_first = False
        
        if rows_html:
            table_html = f'''
            <div class="cs-table-wrap">
              <table class="cs-table">
                <thead>{rows_html[0]}</thead>
                <tbody>{' '.join(rows_html[1:])}</tbody>
              </table>
            </div>'''
            html_parts.append(table_html)

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

with open(r'c:\Users\manoh\OneDrive\Desktop\manohar\portfolio\case-studies.html', 'w', encoding='utf-8') as f:
    f.write(full_html)

print('Created case-studies.html successfully!')
