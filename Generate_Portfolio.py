import json
import os
import base64
import zipfile
import io
import requests
from collections import OrderedDict
from seatable_api import Base, context

# [log patch] usage log for the DocuGenerate quota
import datetime as _dt

def log_generation(base, doc_type, count, source='script'):
    """Record how many DocuGenerate documents this run used (never stops the script)."""
    if not count or count <= 0:
        return
    try:
        base.append_row('Generation log', {
            'Date': _dt.date.today().isoformat(),
            'Document type': doc_type,
            'Documents': count,
            'Source': source,
        })
        print(f'[log] {count} {doc_type} document(s) added to Generation log')
    except Exception as e:
        print(f'[log] WARNING: could not write to Generation log: {e}')


import os as _os

# Folder of this script; falls back to the working directory if __file__ is unavailable
_script_dir  = _os.path.dirname(_os.path.abspath(globals().get('__file__', 'x'))) \
               if '__file__' in globals() else _os.getcwd()
TEMPLATE_DIR = _os.path.join(_script_dir, 'templates')

def load_template_bytes(filename):
    """Read a base64-encoded base docx from the templates/ folder next to this script."""
    path = _os.path.join(TEMPLATE_DIR, filename)
    try:
        with open(path, 'r', encoding='ascii') as f:
            return base64.b64decode(''.join(f.read().split()))
    except FileNotFoundError:
        raise RuntimeError(f'Base template not found: {path}')


def convert_docx_to_pdf(docx_bytes, name, api_key, base=None, doc_type=None):
    """Convert a finished (pre-rendered) docx to PDF with DocuGenerate; returns PDF bytes.

    Flow: upload docx as temporary template -> generate PDF -> download.
    The temporary template is ALWAYS deleted in the finally block, even when
    generation or download fails, so it never occupies a template slot.
    Raises RuntimeError on failure.
    """
    base_url = 'https://api.docugenerate.com/v1'
    headers = {'Authorization': api_key}
    template_id = None
    try:
        print(f'📤 Uploading "{name}" to DocuGenerate as temporary template...')
        up = requests.post(
            f'{base_url}/template',
            headers=headers,
            files={'file': (f'{name}.docx', docx_bytes,
                            'application/vnd.openxmlformats-officedocument'
                            '.wordprocessingml.document')},
            data={'name': name},
            timeout=120,
        )
        if up.status_code not in (200, 201):
            raise RuntimeError(f'Template upload failed: HTTP {up.status_code} {up.text}')
        template_id = up.json().get('id')
        print(f'✅ Template uploaded (id: {template_id})')

        print('📄 Generating PDF...')
        gen = requests.post(
            f'{base_url}/document',
            headers=headers,
            data={
                'template_id':   template_id,
                'name':          name,
                'output_format': '.pdf',
                'data':          '[{}]',   # no merge tags in a pre-rendered docx
            },
            timeout=120,
        )
        if gen.status_code not in (200, 201):
            raise RuntimeError(f'PDF generation failed: HTTP {gen.status_code} {gen.text}')
        pdf_url = gen.json().get('document_uri')
        if base is not None and doc_type:
            log_generation(base, doc_type, 1)   # counts against the monthly quota

        pdf = requests.get(pdf_url, timeout=120)
        pdf.raise_for_status()
        return pdf.content
    finally:
        if template_id:
            try:
                d = requests.delete(f'{base_url}/template/{template_id}',
                                    headers=headers, timeout=60)
                if d.status_code in (200, 204):
                    print('✅ Temporary template deleted')
                else:
                    print(f'⚠️ Template cleanup failed (HTTP {d.status_code}) - '
                          f'delete template {template_id} manually')
            except Exception as e:
                print(f'⚠️ Template cleanup error: {e} - delete template {template_id} manually')



# ─────────────────────────────────────────────────────────────────────────────
# Nutrycs – Generate Portfolio list from the selected Porfolio view
# Builds docx directly with pre-rendered vMerge rows, converts to PDF
# via Docugenerate /v1/convert, saves to OneDrive and sends by email.
# ─────────────────────────────────────────────────────────────────────────────

# ── CONFIG ───────────────────────────────────────────────────────────────────
DOCUGENERATE_API_KEY = os.environ['DOCUGENERATE_API_KEY']
DOCUMENT_NAME        = 'Nutrycs® Product List'

# SeaTable
TABLE_NAME           = 'Portfolio'
VIEW_NAME            = context.current_view or 'All industries'   # auto-detects the view

# Microsoft Graph API
TENANT_ID            = '89163d46-54fe-4ede-95d2-885f4aaf21c4'
CLIENT_ID            = '549aeefa-b3c8-4256-be2b-6da7e71c1883'
CLIENT_SECRET        = os.environ['MS_GRAPH_CLIENT_SECRET']

# Email settings
FROM_EMAIL           = 'eric.marks@umycs.com'
TO_EMAIL             = 'eric.marks@umycs.com'
EMAIL_SUBJECT        = f'Nutrycs® {VIEW_NAME} Product Portfolio'
EMAIL_BODY           = (
    f'Dear colleague,\n\n'
    f'Please find attached the Nutrycs® {VIEW_NAME} Product Portfolio.\n\n'
    f'Best regards,\nUmycs'
)

# OneDrive settings
ONEDRIVE_FOLDER      = 'Product Management Data\Product Portfolio'
PDF_FILENAME         = f'Nutrycs_{VIEW_NAME}_Portfolio.pdf'
# ─────────────────────────────────────────────────────────────────────────────


# ── STEP 1: Connect to SeaTable ──────────────────────────────────────────────
base = Base(context.api_token, context.server_url)
base.auth()
print('✅ Connected to SeaTable')


# ── STEP 2: Fetch active rows ────────────────────────────────────────────────
rows = base.list_rows(TABLE_NAME, view_name=VIEW_NAME)
rows = [row for row in rows if row.get('Active') in (True, 1, '1', 'true', 'True')]
print(f'✅ Found {len(rows)} active rows in view "{VIEW_NAME}"')

if len(rows) == 0:
    print('❌ No active rows found.')
    raise SystemExit


# ── STEP 3: Map and group rows ───────────────────────────────────────────────
def val(v):
    if v is None: return ''
    if isinstance(v, bool): return ''
    if isinstance(v, list):
        if not v: return ''
        first = v[0]
        if isinstance(first, dict):
            return str(first.get('display_value', first.get('name', '')))
        return str(first)
    return str(v)

def xe(s):
    """XML-escape a string."""
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))

flat = []
for row in rows:
    flat.append({
        'industry':            val(row.get('Industry')) or VIEW_NAME,
        'product_type':        val(row.get('Product Type')),
        'product_name':        val(row.get('Product name')),
        'application':         val(row.get('Application')),
        'product_description': val(row.get('Description')),
        'gm_status':           val(row.get('GM Status')),
        'physical_form':       val(row.get('Physical form')),
        'activity':            val(row.get('Activity')),
        'shelf_life_months':   val(row.get('Shelf Life (Months)')),
        'dietary_information': val(row.get('Dietary Information')),
        'allergens':           val(row.get('Allergens')),
        'pack_size_kg':        val(row.get('Pack Size (kg)')),
    })

# Group: industry → product_type → [products]
industry_groups = OrderedDict()
for p in flat:
    industry_groups.setdefault(p['industry'], OrderedDict()) \
                   .setdefault(p['product_type'], []).append(p)

print(f'✅ Grouped: {len(industry_groups)} industr(y/ies), '
      f'{sum(len(v) for v in industry_groups.values())} product types, '
      f'{len(flat)} products total')


# ── STEP 4: XML building blocks ──────────────────────────────────────────────

RUN_PR = ('<w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Arial" '
          'w:hAnsi="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr>')

PP_LEFT   = ('<w:pPr><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" '
             'w:cs="Calibri"/><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr></w:pPr>')
PP_CENTER = ('<w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Calibri" '
             'w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="12"/>'
             '<w:szCs w:val="12"/></w:rPr></w:pPr>')
PP_VERT   = ('<w:pPr><w:ind w:left="113" w:right="113"/><w:jc w:val="center"/>'
             '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
             '<w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr></w:pPr>')

B_STD  = ('<w:tcBorders>'
          '<w:top w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '</w:tcBorders>')
B_CONT = ('<w:tcBorders>'
          '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '</w:tcBorders>')
B_CONT_LAST = ('<w:tcBorders>'
               '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
               '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
               '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
               '</w:tcBorders>')
B_SL   = ('<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>'
          '<w:bottom w:val="nil"/>'
          '<w:right w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '</w:tcBorders>')
B_SR   = ('<w:tcBorders><w:top w:val="nil"/>'
          '<w:left w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>'
          '<w:bottom w:val="nil"/><w:right w:val="nil"/>'
          '</w:tcBorders>')

M_STD  = ('<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
          '<w:bottom w:w="80" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>')
M_ZERO = ('<w:tcMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
          '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>')


def para(text, pp):
    if text:
        return f'<w:p>{pp}<w:r>{RUN_PR}<w:t xml:space="preserve">{xe(text)}</w:t></w:r></w:p>'
    return f'<w:p>{pp}</w:p>'


def tc(width, borders, margins, shading, extra_pr, content):
    return (f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="pct"/>'
            f'{extra_pr}{borders}<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>'
            f'{margins}<w:vAlign w:val="center"/></w:tcPr>{content}</w:tc>')


def cell_stripe_l():
    return tc(260, B_SL, M_ZERO, 'FFFFFF', '', para('', PP_LEFT))

def cell_stripe_r():
    return tc(233, B_SR, M_ZERO, 'FFFFFF', '', para('', PP_CENTER))

def cell_vmerge_start(w, text, shading='FFFFFF'):
    """vMerge restart — first row of a merged group, with vertical text."""
    return tc(w, B_STD, M_STD, shading,
              '<w:vMerge w:val="restart"/><w:textDirection w:val="btLr"/>',
              para(text, PP_VERT))

def cell_vmerge_cont(w, shading='FFFFFF', last=False):
    """vMerge continuation — empty cell. Last row gets a bottom border."""
    borders = B_CONT_LAST if last else B_CONT
    return tc(w, borders, M_STD, shading,
              '<w:vMerge/><w:textDirection w:val="btLr"/>',
              para('', PP_VERT))

def cell_data(w, text, center=False, shading='FFFFFF'):
    return tc(w, B_STD, M_STD, shading, '', para(text, PP_CENTER if center else PP_LEFT))


def build_row(ind_cell, pt_cell, p, page_break=False, shading='FFFFFF'):
    cells = (
        cell_stripe_l() +
        ind_cell +
        pt_cell +
        cell_data(1020, p['product_name'],        shading=shading) +
        cell_data(1020, p['application'],          shading=shading) +
        cell_data(2461, p['product_description'],  shading=shading) +
        cell_data(647,  p['gm_status'],            center=True, shading=shading) +
        cell_data(768,  p['physical_form'],        center=True, shading=shading) +
        cell_data(848,  p['activity'],             center=True, shading=shading) +
        cell_data(974,  p['shelf_life_months'],    center=True, shading=shading) +
        cell_data(1017, p['dietary_information'],  shading=shading) +
        cell_data(817,  p['allergens'],            center=True, shading=shading) +
        cell_data(757,  p['pack_size_kg'],         center=True, shading=shading) +
        cell_stripe_r()
    )
    return (f'<w:tr><w:trPr><w:cantSplit/>'
            f'<w:trHeight w:val="268"/></w:trPr>{cells}</w:tr>')


# ── STEP 5: (rows are built per-table in Step 6) ──────────────────────────────
print(f'✅ Ready to build tables for {len(flat)} products across {sum(len(v) for v in industry_groups.values())} product types')


# ── STEP 6: Build docx — single table, pageBreakBefore on overflowing groups ──
# Product type groups stay intact. If a group does not fit on remaining page space,
# its first row gets pageBreakBefore — keeps everything in ONE continuous table.

# Page break logic: estimate each row's height in twips based on content length,
# track cumulative height, and insert pageBreakBefore on the first row of any
# product type group that doesn't fully fit on the remaining page space.
# Calibrated from observed PDF output (landscape A4).

PAGE_AVAIL   = 8889   # twips available for data rows (content - margins - table header)
CHAR_W       = 82     # twips per character at 12pt Calibri (calibrated)
LINE_H       = 200    # twips per line
CELL_PAD     = 160    # top + bottom cell padding

COL_WIDTHS = {
    'product_name': 1020, 'application': 1020, 'product_description': 2461,
    'gm_status': 647, 'physical_form': 768, 'activity': 848,
    'shelf_life_months': 974,
    'dietary_information': 1017, 'allergens': 817, 'pack_size_kg': 757,
}

def estimate_row_height(p):
    max_lines = 1
    for col, w in COL_WIDTHS.items():
        text = str(p.get(col, '') or '')
        if not text: continue
        chars_per_line = max(1, w // CHAR_W)
        lines = max(1, -(-len(text) // chars_per_line))
        max_lines = max(max_lines, lines)
    return max(268, max_lines * LINE_H + CELL_PAD)

# Alternate shading between product type groups for visual separation.
# Industry and Product Type text repeats on every row — fully page-break safe.
# ── Build tables: one per page, new table started when a group doesn't fit ────
# Each table has its own header row. When a product type group would overflow
# the remaining page space, the current table is closed and a new one opens
# on a fresh page. vMerge works cleanly because it never crosses a page boundary.

import base64 as _b64t
TABLE_OPEN_XML  = _b64t.b64decode("PHc6dGJsPjx3OnRibFByPjx3OnRibFcgdzp3PSIxNDAwOCIgdzp0eXBlPSJkeGEiLz48dzp0YmxJbmQgdzp3PSIwIiB3OnR5cGU9ImR4YSIvPjx3OnRibEJvcmRlcnM+PHc6dG9wIHc6dmFsPSJub25lIiB3OnN6PSIwIiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJhdXRvIi8+PHc6bGVmdCB3OnZhbD0ibm9uZSIgdzpzej0iMCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iYXV0byIvPjx3OmJvdHRvbSB3OnZhbD0ibm9uZSIgdzpzej0iMCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iYXV0byIvPjx3OnJpZ2h0IHc6dmFsPSJub25lIiB3OnN6PSIwIiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJhdXRvIi8+PHc6aW5zaWRlSCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzppbnNpZGVWIHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjwvdzp0YmxCb3JkZXJzPjx3OnRibExheW91dCB3OnR5cGU9ImZpeGVkIi8+PHc6dGJsQ2VsbE1hcj48dzpsZWZ0IHc6dz0iMTAiIHc6dHlwZT0iZHhhIi8+PHc6cmlnaHQgdzp3PSIxMCIgdzp0eXBlPSJkeGEiLz48L3c6dGJsQ2VsbE1hcj48dzp0YmxMb29rIHc6dmFsPSIwMDAwIiB3OmZpcnN0Um93PSIwIiB3Omxhc3RSb3c9IjAiIHc6Zmlyc3RDb2x1bW49IjAiIHc6bGFzdENvbHVtbj0iMCIgdzpub0hCYW5kPSIwIiB3Om5vVkJhbmQ9IjAiLz48L3c6dGJsUHI+PHc6dGJsR3JpZD48dzpncmlkQ29sIHc6dz0iNzI5Ii8+PHc6Z3JpZENvbCB3Onc9Ijc2NSIvPjx3OmdyaWRDb2wgdzp3PSI3NjUiLz48dzpncmlkQ29sIHc6dz0iMTAyMCIvPjx3OmdyaWRDb2wgdzp3PSIxMDIwIi8+PHc6Z3JpZENvbCB3Onc9IjI0NjEiLz48dzpncmlkQ29sIHc6dz0iNjQ3Ii8+PHc6Z3JpZENvbCB3Onc9Ijc2OCIvPjx3OmdyaWRDb2wgdzp3PSI3NjgiLz48dzpncmlkQ29sIHc6dz0iODQ4Ii8+PHc6Z3JpZENvbCB3Onc9IjEwMTciLz48dzpncmlkQ29sIHc6dz0iODE3Ii8+PHc6Z3JpZENvbCB3Onc9Ijc1NyIvPjx3OmdyaWRDb2wgdzp3PSI2NTIiLz48L3c6dGJsR3JpZD48dzp0ciB3OnJzaWRSPSIwMEQ3N0U1QiIgdzpyc2lkUlByPSIwMDk3MzQ1MyIgdzE0OnBhcmFJZD0iMDkzNkMzMjMiIHcxNDp0ZXh0SWQ9Ijc3Nzc3Nzc3IiB3OnJzaWRUcj0iMDA4ODREQjkiPjx3OnRyUHI+PHc6dHJIZWlnaHQgdzp2YWw9IjQ2NyIvPjx3OnRibEhlYWRlci8+PC93OnRyUHI+PHc6dGM+PHc6dGNQcj48dzp0Y1cgdzp3PSI3MjkiIHc6dHlwZT0iZHhhIi8+PHc6dGNCb3JkZXJzPjx3OnRvcCB3OnZhbD0ibmlsIi8+PHc6bGVmdCB3OnZhbD0ibmlsIi8+PHc6Ym90dG9tIHc6dmFsPSJuaWwiLz48dzpyaWdodCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48L3c6dGNCb3JkZXJzPjx3OnNoZCB3OnZhbD0iY2xlYXIiIHc6Y29sb3I9ImF1dG8iIHc6ZmlsbD0iRkZGRkZGIi8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iMCIgdzp0eXBlPSJkeGEiLz48dzpsZWZ0IHc6dz0iMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSIwIiB3OnR5cGU9ImR4YSIvPjx3OnJpZ2h0IHc6dz0iMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9IjQyOUJFMEQ1IiB3MTQ6dGV4dElkPSI3Nzc3Nzc3NyIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDA5NzM0NTMiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA2QTBGMjgiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48dzpsYW5nIHc6dmFsPSJlbi1VUyIvPjwvdzpyUHI+PC93OnBQcj48L3c6cD48L3c6dGM+PHc6dGM+PHc6dGNQcj48dzp0Y1cgdzp3PSI3NjUiIHc6dHlwZT0iZHhhIi8+PHc6dGNCb3JkZXJzPjx3OnRvcCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpsZWZ0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmJvdHRvbSB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpyaWdodCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48L3c6dGNCb3JkZXJzPjx3OnNoZCB3OnZhbD0iY2xlYXIiIHc6Y29sb3I9ImF1dG8iIHc6ZmlsbD0iQTZBNkE2IiB3OnRoZW1lRmlsbD0iYmFja2dyb3VuZDEiIHc6dGhlbWVGaWxsU2hhZGU9IkE2Ii8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6bGVmdCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpyaWdodCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9IjM3MjM0RjUyIiB3MTQ6dGV4dElkPSI2MTVFNzU1RiIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDA5NzM0NTMiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA5MzBBNUYiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnByb29mRXJyIHc6dHlwZT0ic3BlbGxTdGFydCIvPjx3OnI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0PkluZHVzdHJ5PC93OnQ+PC93OnI+PHc6cHJvb2ZFcnIgdzp0eXBlPSJzcGVsbEVuZCIvPjwvdzpwPjwvdzp0Yz48dzp0Yz48dzp0Y1ByPjx3OnRjVyB3Onc9Ijc2NSIgdzp0eXBlPSJkeGEiLz48dzp0Y0JvcmRlcnM+PHc6dG9wIHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmxlZnQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6Ym90dG9tIHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OnJpZ2h0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjwvdzp0Y0JvcmRlcnM+PHc6c2hkIHc6dmFsPSJjbGVhciIgdzpjb2xvcj0iYXV0byIgdzpmaWxsPSJBNkE2QTYiIHc6dGhlbWVGaWxsPSJiYWNrZ3JvdW5kMSIgdzp0aGVtZUZpbGxTaGFkZT0iQTYiLz48dzp0Y01hcj48dzp0b3Agdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpsZWZ0IHc6dz0iMTAwIiB3OnR5cGU9ImR4YSIvPjx3OmJvdHRvbSB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OnJpZ2h0IHc6dz0iMTAwIiB3OnR5cGU9ImR4YSIvPjwvdzp0Y01hcj48dzp2QWxpZ24gdzp2YWw9ImNlbnRlciIvPjwvdzp0Y1ByPjx3OnAgdzE0OnBhcmFJZD0iNzEwRDJEQUYiIHcxNDp0ZXh0SWQ9IjBBNDVFOUI0IiB3OnJzaWRSPSIwMEQ3N0U1QiIgdzpyc2lkUlByPSIwMDk3MzQ1MyIgdzpyc2lkUkRlZmF1bHQ9IjAwRDc3RTVCIiB3OnJzaWRQPSIwMDVDM0JCNSI+PHc6cFByPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0PlByb2R1Y3QgVHlwZTwvdzp0PjwvdzpyPjwvdzpwPjwvdzp0Yz48dzp0Yz48dzp0Y1ByPjx3OnRjVyB3Onc9IjEwMjAiIHc6dHlwZT0iZHhhIi8+PHc6dGNCb3JkZXJzPjx3OnRvcCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpsZWZ0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmJvdHRvbSB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpyaWdodCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48L3c6dGNCb3JkZXJzPjx3OnNoZCB3OnZhbD0iY2xlYXIiIHc6Y29sb3I9ImF1dG8iIHc6ZmlsbD0iQTZBNkE2IiB3OnRoZW1lRmlsbD0iYmFja2dyb3VuZDEiIHc6dGhlbWVGaWxsU2hhZGU9IkE2Ii8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6bGVmdCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpyaWdodCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9IjIxOUU3ODY3IiB3MTQ6dGV4dElkPSI3Nzc3Nzc3NyIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDA5NzM0NTMiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA2QTBGMjgiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjwvdzpwUHI+PHc6ciB3OnJzaWRSUHI9IjAwOTczNDUzIj48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6ZWFzdEFzaWE9IkFyaWFsIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6Yi8+PHc6YkNzLz48dzpjb2xvciB3OnZhbD0iRkZGRkZGIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjx3OnQ+UHJvZHVjdCBOYW1lPC93OnQ+PC93OnI+PC93OnA+PC93OnRjPjx3OnRjPjx3OnRjUHI+PHc6dGNXIHc6dz0iMTAyMCIgdzp0eXBlPSJkeGEiLz48dzp0Y0JvcmRlcnM+PHc6dG9wIHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmxlZnQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6Ym90dG9tIHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OnJpZ2h0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjwvdzp0Y0JvcmRlcnM+PHc6c2hkIHc6dmFsPSJjbGVhciIgdzpjb2xvcj0iYXV0byIgdzpmaWxsPSJBNkE2QTYiIHc6dGhlbWVGaWxsPSJiYWNrZ3JvdW5kMSIgdzp0aGVtZUZpbGxTaGFkZT0iQTYiLz48dzp0Y01hcj48dzp0b3Agdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpsZWZ0IHc6dz0iMTAwIiB3OnR5cGU9ImR4YSIvPjx3OmJvdHRvbSB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OnJpZ2h0IHc6dz0iMTAwIiB3OnR5cGU9ImR4YSIvPjwvdzp0Y01hcj48dzp2QWxpZ24gdzp2YWw9ImNlbnRlciIvPjwvdzp0Y1ByPjx3OnAgdzE0OnBhcmFJZD0iNTc3RERCOTAiIHcxNDp0ZXh0SWQ9Ijc3Nzc3Nzc3IiB3OnJzaWRSPSIwMEQ3N0U1QiIgdzpyc2lkUlByPSIwMDk3MzQ1MyIgdzpyc2lkUkRlZmF1bHQ9IjAwRDc3RTVCIiB3OnJzaWRQPSIwMDZBMEYyOCI+PHc6cFByPjx3OmpjIHc6dmFsPSJjZW50ZXIiLz48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6aEFuc2k9IkNhbGlicmkiIHc6Y3M9IkNhbGlicmkiLz48dzpzeiB3OnZhbD0iMTYiLz48dzpzekNzIHc6dmFsPSIxNiIvPjwvdzpyUHI+PC93OnBQcj48dzpyIHc6cnNpZFJQcj0iMDA5NzM0NTMiPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzplYXN0QXNpYT0iQXJpYWwiIHc6aEFuc2k9IkNhbGlicmkiIHc6Y3M9IkNhbGlicmkiLz48dzpiLz48dzpiQ3MvPjx3OmNvbG9yIHc6dmFsPSJGRkZGRkYiLz48dzpzeiB3OnZhbD0iMTYiLz48dzpzekNzIHc6dmFsPSIxNiIvPjwvdzpyUHI+PHc6dD5BcHBsaWNhdGlvbjwvdzp0PjwvdzpyPjwvdzpwPjwvdzp0Yz48dzp0Yz48dzp0Y1ByPjx3OnRjVyB3Onc9IjI0NjEiIHc6dHlwZT0iZHhhIi8+PHc6dGNCb3JkZXJzPjx3OnRvcCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpsZWZ0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmJvdHRvbSB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpyaWdodCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48L3c6dGNCb3JkZXJzPjx3OnNoZCB3OnZhbD0iY2xlYXIiIHc6Y29sb3I9ImF1dG8iIHc6ZmlsbD0iQTZBNkE2IiB3OnRoZW1lRmlsbD0iYmFja2dyb3VuZDEiIHc6dGhlbWVGaWxsU2hhZGU9IkE2Ii8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6bGVmdCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpyaWdodCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9IjEyNUNGQzdCIiB3MTQ6dGV4dElkPSI3Nzc3Nzc3NyIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDA5NzM0NTMiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA2QTBGMjgiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjwvdzpwUHI+PHc6ciB3OnJzaWRSUHI9IjAwOTczNDUzIj48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6ZWFzdEFzaWE9IkFyaWFsIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6Yi8+PHc6YkNzLz48dzpjb2xvciB3OnZhbD0iRkZGRkZGIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjx3OnQgeG1sOnNwYWNlPSJwcmVzZXJ2ZSI+UHJvZHVjdCA8L3c6dD48L3c6cj48dzpwcm9vZkVyciB3OnR5cGU9InNwZWxsU3RhcnQiLz48dzpyIHc6cnNpZFJQcj0iMDA5NzM0NTMiPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzplYXN0QXNpYT0iQXJpYWwiIHc6aEFuc2k9IkNhbGlicmkiIHc6Y3M9IkNhbGlicmkiLz48dzpiLz48dzpiQ3MvPjx3OmNvbG9yIHc6dmFsPSJGRkZGRkYiLz48dzpzeiB3OnZhbD0iMTYiLz48dzpzekNzIHc6dmFsPSIxNiIvPjwvdzpyUHI+PHc6dD5EZXNjcmlwdGlvbjwvdzp0PjwvdzpyPjx3OnByb29mRXJyIHc6dHlwZT0ic3BlbGxFbmQiLz48L3c6cD48L3c6dGM+PHc6dGM+PHc6dGNQcj48dzp0Y1cgdzp3PSI2NDciIHc6dHlwZT0iZHhhIi8+PHc6dGNCb3JkZXJzPjx3OnRvcCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpsZWZ0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmJvdHRvbSB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpyaWdodCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48L3c6dGNCb3JkZXJzPjx3OnNoZCB3OnZhbD0iY2xlYXIiIHc6Y29sb3I9ImF1dG8iIHc6ZmlsbD0iQTZBNkE2IiB3OnRoZW1lRmlsbD0iYmFja2dyb3VuZDEiIHc6dGhlbWVGaWxsU2hhZGU9IkE2Ii8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6bGVmdCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpyaWdodCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9IjE0NEVCNDAyIiB3MTQ6dGV4dElkPSI3Nzc3Nzc3NyIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDA5NzM0NTMiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA2QTBGMjgiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjwvdzpwUHI+PHc6ciB3OnJzaWRSUHI9IjAwOTczNDUzIj48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6ZWFzdEFzaWE9IkFyaWFsIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6Yi8+PHc6YkNzLz48dzpjb2xvciB3OnZhbD0iRkZGRkZGIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjx3OnQ+R00gU3RhdHVzPC93OnQ+PC93OnI+PC93OnA+PC93OnRjPjx3OnRjPjx3OnRjUHI+PHc6dGNXIHc6dz0iNzY4IiB3OnR5cGU9ImR4YSIvPjx3OnRjQm9yZGVycz48dzp0b3Agdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6bGVmdCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpib3R0b20gdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6cmlnaHQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PC93OnRjQm9yZGVycz48dzpzaGQgdzp2YWw9ImNsZWFyIiB3OmNvbG9yPSJhdXRvIiB3OmZpbGw9IkE2QTZBNiIgdzp0aGVtZUZpbGw9ImJhY2tncm91bmQxIiB3OnRoZW1lRmlsbFNoYWRlPSJBNiIvPjx3OnRjTWFyPjx3OnRvcCB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OmxlZnQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PHc6Ym90dG9tIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6cmlnaHQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PC93OnRjTWFyPjx3OnZBbGlnbiB3OnZhbD0iY2VudGVyIi8+PC93OnRjUHI+PHc6cCB3MTQ6cGFyYUlkPSIyQTVDNjI3QiIgdzE0OnRleHRJZD0iNzc3Nzc3NzciIHc6cnNpZFI9IjAwRDc3RTVCIiB3OnJzaWRSUHI9IjAwOTczNDUzIiB3OnJzaWRSRGVmYXVsdD0iMDBENzdFNUIiIHc6cnNpZFA9IjAwNkEwRjI4Ij48dzpwUHI+PHc6amMgdzp2YWw9ImNlbnRlciIvPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnByb29mRXJyIHc6dHlwZT0ic3BlbGxTdGFydCIvPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0PlBoeXNpY2FsPC93OnQ+PC93OnI+PHc6cHJvb2ZFcnIgdzp0eXBlPSJzcGVsbEVuZCIvPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0IHhtbDpzcGFjZT0icHJlc2VydmUiPiBGb3JtPC93OnQ+PC93OnI+PC93OnA+PC93OnRjPjx3OnRjPjx3OnRjUHI+PHc6dGNXIHc6dz0iNzY4IiB3OnR5cGU9ImR4YSIvPjx3OnRjQm9yZGVycz48dzp0b3Agdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6bGVmdCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpib3R0b20gdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6cmlnaHQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PC93OnRjQm9yZGVycz48dzpzaGQgdzp2YWw9ImNsZWFyIiB3OmNvbG9yPSJhdXRvIiB3OmZpbGw9IkE2QTZBNiIgdzp0aGVtZUZpbGw9ImJhY2tncm91bmQxIiB3OnRoZW1lRmlsbFNoYWRlPSJBNiIvPjx3OnRjTWFyPjx3OnRvcCB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OmxlZnQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PHc6Ym90dG9tIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6cmlnaHQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PC93OnRjTWFyPjx3OnZBbGlnbiB3OnZhbD0iY2VudGVyIi8+PC93OnRjUHI+PHc6cCB3MTQ6cGFyYUlkPSIzNTNBRDM4QSIgdzE0OnRleHRJZD0iNzc3Nzc3NzciIHc6cnNpZFI9IjAwRDc3RTVCIiB3OnJzaWRSUHI9IjAwOTczNDUzIiB3OnJzaWRSRGVmYXVsdD0iMDBENzdFNUIiIHc6cnNpZFA9IjAwOTMwQTVGIj48dzpwUHI+PHc6aW5kIHc6bGVmdD0iLTIwIi8+PHc6amMgdzp2YWw9ImNlbnRlciIvPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0PkFjdGl2aXR5PC93OnQ+PC93OnI+PC93OnA+PC93OnRjPjx3OnRjPjx3OnRjUHI+PHc6dGNXIHc6dz0iODQ4IiB3OnR5cGU9ImR4YSIvPjx3OnRjQm9yZGVycz48dzp0b3Agdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6bGVmdCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpib3R0b20gdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6cmlnaHQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PC93OnRjQm9yZGVycz48dzpzaGQgdzp2YWw9ImNsZWFyIiB3OmNvbG9yPSJhdXRvIiB3OmZpbGw9IkE2QTZBNiIgdzp0aGVtZUZpbGw9ImJhY2tncm91bmQxIiB3OnRoZW1lRmlsbFNoYWRlPSJBNiIvPjx3OnRjTWFyPjx3OnRvcCB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OmxlZnQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PHc6Ym90dG9tIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6cmlnaHQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PC93OnRjTWFyPjx3OnZBbGlnbiB3OnZhbD0iY2VudGVyIi8+PC93OnRjUHI+PHc6cD48dzpwUHI+PHc6amMgdzp2YWw9ImNlbnRlciIvPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0IHhtbDpzcGFjZT0icHJlc2VydmUiPlNoZWxmPC93OnQ+PC93OnI+PC93OnA+PHc6cD48dzpwUHI+PHc6amMgdzp2YWw9ImNlbnRlciIvPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0IHhtbDpzcGFjZT0icHJlc2VydmUiPkxpZmU8L3c6dD48L3c6cj48L3c6cD48dzpwPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjwvdzpwUHI+PHc6cj48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6ZWFzdEFzaWE9IkFyaWFsIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6Yi8+PHc6YkNzLz48dzpjb2xvciB3OnZhbD0iRkZGRkZGIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjx3OnQgeG1sOnNwYWNlPSJwcmVzZXJ2ZSI+KE1vbnRocyk8L3c6dD48L3c6cj48L3c6cD48L3c6dGM+PHc6dGM+PHc6dGNQcj48dzp0Y1cgdzp3PSIxMDE3IiB3OnR5cGU9ImR4YSIvPjx3OnRjQm9yZGVycz48dzp0b3Agdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6bGVmdCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpib3R0b20gdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6cmlnaHQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PC93OnRjQm9yZGVycz48dzpzaGQgdzp2YWw9ImNsZWFyIiB3OmNvbG9yPSJhdXRvIiB3OmZpbGw9IkE2QTZBNiIgdzp0aGVtZUZpbGw9ImJhY2tncm91bmQxIiB3OnRoZW1lRmlsbFNoYWRlPSJBNiIvPjx3OnRjTWFyPjx3OnRvcCB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OmxlZnQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PHc6Ym90dG9tIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6cmlnaHQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PC93OnRjTWFyPjx3OnZBbGlnbiB3OnZhbD0iY2VudGVyIi8+PC93OnRjUHI+PHc6cCB3MTQ6cGFyYUlkPSI0QTFGQUZBQSIgdzE0OnRleHRJZD0iNzc3Nzc3NzciIHc6cnNpZFI9IjAwRDc3RTVCIiB3OnJzaWRSUHI9IjAwOTczNDUzIiB3OnJzaWRSRGVmYXVsdD0iMDBENzdFNUIiIHc6cnNpZFA9IjAwNkEwRjI4Ij48dzpwUHI+PHc6amMgdzp2YWw9ImNlbnRlciIvPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnByb29mRXJyIHc6dHlwZT0ic3BlbGxTdGFydCIvPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0PkRpZXRhcnk8L3c6dD48L3c6cj48dzpwcm9vZkVyciB3OnR5cGU9InNwZWxsRW5kIi8+PHc6ciB3OnJzaWRSUHI9IjAwOTczNDUzIj48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6ZWFzdEFzaWE9IkFyaWFsIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6Yi8+PHc6YkNzLz48dzpjb2xvciB3OnZhbD0iRkZGRkZGIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjx3OnQgeG1sOnNwYWNlPSJwcmVzZXJ2ZSI+IEluZm9ybWF0aW9uPC93OnQ+PC93OnI+PC93OnA+PC93OnRjPjx3OnRjPjx3OnRjUHI+PHc6dGNXIHc6dz0iODE3IiB3OnR5cGU9ImR4YSIvPjx3OnRjQm9yZGVycz48dzp0b3Agdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6bGVmdCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpib3R0b20gdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6cmlnaHQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PC93OnRjQm9yZGVycz48dzpzaGQgdzp2YWw9ImNsZWFyIiB3OmNvbG9yPSJhdXRvIiB3OmZpbGw9IkE2QTZBNiIgdzp0aGVtZUZpbGw9ImJhY2tncm91bmQxIiB3OnRoZW1lRmlsbFNoYWRlPSJBNiIvPjx3OnRjTWFyPjx3OnRvcCB3Onc9IjgwIiB3OnR5cGU9ImR4YSIvPjx3OmxlZnQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PHc6Ym90dG9tIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6cmlnaHQgdzp3PSIxMDAiIHc6dHlwZT0iZHhhIi8+PC93OnRjTWFyPjx3OnZBbGlnbiB3OnZhbD0iY2VudGVyIi8+PC93OnRjUHI+PHc6cCB3MTQ6cGFyYUlkPSIzQ0EyRTY0QyIgdzE0OnRleHRJZD0iNzc3Nzc3NzciIHc6cnNpZFI9IjAwRDc3RTVCIiB3OnJzaWRSUHI9IjAwOTczNDUzIiB3OnJzaWRSRGVmYXVsdD0iMDBENzdFNUIiIHc6cnNpZFA9IjAwNkEwRjI4Ij48dzpwUHI+PHc6amMgdzp2YWw9ImNlbnRlciIvPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48L3c6cFByPjx3OnByb29mRXJyIHc6dHlwZT0ic3BlbGxTdGFydCIvPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0PkFsbGVyZ2Vuczwvdzp0PjwvdzpyPjx3OnByb29mRXJyIHc6dHlwZT0ic3BlbGxFbmQiLz48L3c6cD48L3c6dGM+PHc6dGM+PHc6dGNQcj48dzp0Y1cgdzp3PSI3NTciIHc6dHlwZT0iZHhhIi8+PHc6dGNCb3JkZXJzPjx3OnRvcCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpsZWZ0IHc6dmFsPSJzaW5nbGUiIHc6c3o9IjQiIHc6c3BhY2U9IjAiIHc6Y29sb3I9IkNDQ0NDQyIvPjx3OmJvdHRvbSB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48dzpyaWdodCB3OnZhbD0ic2luZ2xlIiB3OnN6PSI0IiB3OnNwYWNlPSIwIiB3OmNvbG9yPSJDQ0NDQ0MiLz48L3c6dGNCb3JkZXJzPjx3OnNoZCB3OnZhbD0iY2xlYXIiIHc6Y29sb3I9ImF1dG8iIHc6ZmlsbD0iQTZBNkE2IiB3OnRoZW1lRmlsbD0iYmFja2dyb3VuZDEiIHc6dGhlbWVGaWxsU2hhZGU9IkE2Ii8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iODAiIHc6dHlwZT0iZHhhIi8+PHc6bGVmdCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSI4MCIgdzp0eXBlPSJkeGEiLz48dzpyaWdodCB3Onc9IjEwMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9IjUwNjlDREM1IiB3MTQ6dGV4dElkPSI3Nzc3Nzc3NyIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDA5NzM0NTMiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA2QTBGMjgiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjwvdzpwUHI+PHc6ciB3OnJzaWRSUHI9IjAwOTczNDUzIj48dzpyUHI+PHc6ckZvbnRzIHc6YXNjaWk9IkNhbGlicmkiIHc6ZWFzdEFzaWE9IkFyaWFsIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6Yi8+PHc6YkNzLz48dzpjb2xvciB3OnZhbD0iRkZGRkZGIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48L3c6clByPjx3OnQgeG1sOnNwYWNlPSJwcmVzZXJ2ZSI+UGFjayA8L3c6dD48L3c6cj48dzpwcm9vZkVyciB3OnR5cGU9InNwZWxsU3RhcnQiLz48dzpyIHc6cnNpZFJQcj0iMDA5NzM0NTMiPjx3OnJQcj48dzpyRm9udHMgdzphc2NpaT0iQ2FsaWJyaSIgdzplYXN0QXNpYT0iQXJpYWwiIHc6aEFuc2k9IkNhbGlicmkiIHc6Y3M9IkNhbGlicmkiLz48dzpiLz48dzpiQ3MvPjx3OmNvbG9yIHc6dmFsPSJGRkZGRkYiLz48dzpzeiB3OnZhbD0iMTYiLz48dzpzekNzIHc6dmFsPSIxNiIvPjwvdzpyUHI+PHc6dD5TaXplPC93OnQ+PC93OnI+PHc6cHJvb2ZFcnIgdzp0eXBlPSJzcGVsbEVuZCIvPjx3OnIgdzpyc2lkUlByPSIwMDk3MzQ1MyI+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmVhc3RBc2lhPSJBcmlhbCIgdzpoQW5zaT0iQ2FsaWJyaSIgdzpjcz0iQ2FsaWJyaSIvPjx3OmIvPjx3OmJDcy8+PHc6Y29sb3Igdzp2YWw9IkZGRkZGRiIvPjx3OnN6IHc6dmFsPSIxNiIvPjx3OnN6Q3Mgdzp2YWw9IjE2Ii8+PC93OnJQcj48dzp0IHhtbDpzcGFjZT0icHJlc2VydmUiPiAoa2cpPC93OnQ+PC93OnI+PC93OnA+PC93OnRjPjx3OnRjPjx3OnRjUHI+PHc6dGNXIHc6dz0iNjUyIiB3OnR5cGU9ImR4YSIvPjx3OnRjQm9yZGVycz48dzp0b3Agdzp2YWw9Im5pbCIvPjx3OmxlZnQgdzp2YWw9InNpbmdsZSIgdzpzej0iNCIgdzpzcGFjZT0iMCIgdzpjb2xvcj0iQ0NDQ0NDIi8+PHc6Ym90dG9tIHc6dmFsPSJuaWwiLz48dzpyaWdodCB3OnZhbD0ibmlsIi8+PC93OnRjQm9yZGVycz48dzpzaGQgdzp2YWw9ImNsZWFyIiB3OmNvbG9yPSJhdXRvIiB3OmZpbGw9IkZGRkZGRiIgdzp0aGVtZUZpbGw9ImJhY2tncm91bmQxIi8+PHc6dGNNYXI+PHc6dG9wIHc6dz0iMCIgdzp0eXBlPSJkeGEiLz48dzpsZWZ0IHc6dz0iMCIgdzp0eXBlPSJkeGEiLz48dzpib3R0b20gdzp3PSIwIiB3OnR5cGU9ImR4YSIvPjx3OnJpZ2h0IHc6dz0iMCIgdzp0eXBlPSJkeGEiLz48L3c6dGNNYXI+PHc6dkFsaWduIHc6dmFsPSJjZW50ZXIiLz48L3c6dGNQcj48dzpwIHcxNDpwYXJhSWQ9Ijc4NkQwM0M2IiB3MTQ6dGV4dElkPSI3Nzc3Nzc3NyIgdzpyc2lkUj0iMDBENzdFNUIiIHc6cnNpZFJQcj0iMDBENzdFNUIiIHc6cnNpZFJEZWZhdWx0PSIwMEQ3N0U1QiIgdzpyc2lkUD0iMDA2QTBGMjgiPjx3OnBQcj48dzpqYyB3OnZhbD0iY2VudGVyIi8+PHc6clByPjx3OnJGb250cyB3OmFzY2lpPSJDYWxpYnJpIiB3OmhBbnNpPSJDYWxpYnJpIiB3OmNzPSJDYWxpYnJpIi8+PHc6c3ogdzp2YWw9IjE2Ii8+PHc6c3pDcyB3OnZhbD0iMTYiLz48dzpsYW5nIHc6dmFsPSJlbi1VUyIvPjwvdzpyUHI+PC93OnBQcj48L3c6cD48L3c6dGM+PC93OnRyPg==").decode('utf-8')
TABLE_CLOSE_XML = '</w:tbl>'

# Separator between tables: minimal paragraph + page break
PAGE_BREAK = (
    '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/>'
    '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>'
    '<w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
    '<w:br w:type="page"/></w:r></w:p>'
)

PAGE_AVAIL = 8889   # twips available for data rows per page
CHAR_W     = 82     # twips per char at 12pt Calibri
LINE_H     = 200    # twips per line
CELL_PAD   = 160    # top+bottom cell padding

COL_WIDTHS = {
    'product_name': 1020, 'application': 1020, 'product_description': 2461,
    'gm_status': 647, 'physical_form': 768, 'activity': 848,
    'shelf_life_months': 974, 
    'dietary_information': 1017, 'allergens': 817, 'pack_size_kg': 757,
}

def estimate_row_height(p):
    max_lines = 1
    for col, w in COL_WIDTHS.items():
        text = str(p.get(col, '') or '')
        if not text: continue
        chars_per_line = max(1, w // CHAR_W)
        lines = max(1, -(-len(text) // chars_per_line))
        max_lines = max(max_lines, lines)
    return max(268, max_lines * LINE_H + CELL_PAD)

SHADING_A = 'FFFFFF'
SHADING_B = 'F0F0F0'

# Build list of (table_xml) — one entry per page
tables      = []           # list of complete table XML strings
cur_rows    = []           # rows accumulated for current table
used        = 0            # twips used on current page
group_index = 0            # for alternating shading

for industry, type_dict in industry_groups.items():
    for product_type, products in type_dict.items():
        shading   = SHADING_A if group_index % 2 == 0 else SHADING_B
        group_h   = sum(estimate_row_height(p) for p in products)
        at_top    = (used == 0)

        # If this group doesn't fit AND we're not at the top, close current table
        # and start a new one on the next page
        if not at_top and group_h > (PAGE_AVAIL - used):
            tables.append(TABLE_OPEN_XML + '\n'.join(cur_rows) + TABLE_CLOSE_XML)
            cur_rows = []
            used     = 0

        # Add all rows of this group to the current table
        last_i = len(products) - 1
        for i, p in enumerate(products):
            is_last  = (i == last_i)
            ind_cell = cell_vmerge_start(765, industry,     shading) if i == 0 else cell_vmerge_cont(765, shading, last=is_last)
            pt_cell  = cell_vmerge_start(765, product_type, shading) if i == 0 else cell_vmerge_cont(765, shading, last=is_last)
            cur_rows.append(build_row(ind_cell, pt_cell, p, shading=shading))

        used        += group_h
        group_index += 1

# Close the last table
if cur_rows:
    tables.append(TABLE_OPEN_XML + '\n'.join(cur_rows) + TABLE_CLOSE_XML)

# Join tables with page breaks between them
body_content = PAGE_BREAK.join(tables)
print(f'\u2705 Built {len(tables)} table(s) across {group_index} product type(s), {len(flat)} products')

# Reassemble docx from original embedded template
import zipfile as _zf, io as _io
_tmpl_bytes = load_template_bytes('portfolio_template.b64')
src_zip     = _zf.ZipFile(_io.BytesIO(_tmpl_bytes))
doc_xml     = src_zip.read('word/document.xml').decode('utf-8')

# Replace entire table with our multi-table body_content
tbl_start   = doc_xml.find('<w:tbl>')
tbl_end     = doc_xml.rfind('</w:tbl>') + len('</w:tbl>')
doc_xml_new = doc_xml[:tbl_start] + body_content + doc_xml[tbl_end:]

out_buf = _io.BytesIO()
with _zf.ZipFile(out_buf, 'w', _zf.ZIP_DEFLATED) as zout:
    for name in src_zip.namelist():
        data = doc_xml_new.encode('utf-8') if name == 'word/document.xml' \
               else src_zip.read(name)
        zout.writestr(name, data)

docx_bytes = out_buf.getvalue()
print(f'\u2705 Docx assembled ({len(docx_bytes):,} bytes)')

# ── STEP 7: Convert the pre-rendered docx to PDF via DocuGenerate ────────────
# convert_docx_to_pdf() uploads a temporary template, generates the PDF and
# always deletes the template again (try/finally), so only one template slot
# is used at a time and a failed run cannot leave templates behind.
pdf_content = convert_docx_to_pdf(docx_bytes, DOCUMENT_NAME, DOCUGENERATE_API_KEY,
                                  base=base, doc_type='Portfolio')
print(f'✅ PDF downloaded ({len(pdf_content):,} bytes)')


# ── STEP 8: Microsoft Graph access token ─────────────────────────────────────
print('🔐 Getting Microsoft Graph access token...')

tok = requests.post(
    f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token',
    data={'grant_type': 'client_credentials', 'client_id': CLIENT_ID,
          'client_secret': CLIENT_SECRET,
          'scope': 'https://graph.microsoft.com/.default'}
)

if tok.status_code != 200:
    print(f'❌ Token error: {tok.text}')
    raise SystemExit

access_token  = tok.json()['access_token']
graph_headers = {'Authorization': f'Bearer {access_token}',
                 'Content-Type': 'application/json'}
print('✅ Access token obtained')


# ── STEP 9: Save to OneDrive ──────────────────────────────────────────────────
print(f'💾 Saving PDF to OneDrive "{ONEDRIVE_FOLDER}"...')

up = requests.put(
    f'https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/drive/root:/'
    f'{ONEDRIVE_FOLDER}/{PDF_FILENAME}:/content',
    headers={'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/pdf'},
    data=pdf_content
)

if up.status_code in [200, 201]:
    print(f'✅ Saved to OneDrive: {up.json().get("webUrl")}')
else:
    print(f'⚠️ OneDrive upload failed: {up.status_code} {up.text}')


# ── STEP 10: Send email ───────────────────────────────────────────────────────
print(f'📧 Sending email to {TO_EMAIL}...')

mail = requests.post(
    f'https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/sendMail',
    headers=graph_headers,
    json={'message': {
        'subject': EMAIL_SUBJECT,
        'body': {'contentType': 'Text', 'content': EMAIL_BODY},
        'toRecipients': [{'emailAddress': {'address': TO_EMAIL}}],
        'attachments': [{
            '@odata.type':  '#microsoft.graph.fileAttachment',
            'name':         PDF_FILENAME,
            'contentType':  'application/pdf',
            'contentBytes': base64.b64encode(pdf_content).decode('utf-8')
        }]
    }, 'saveToSentItems': True}
)

if mail.status_code == 202:
    print(f'✅ Email sent to {TO_EMAIL}')
else:
    print(f'❌ Email error: {mail.status_code} {mail.text}')

print('🎉 All done!')

