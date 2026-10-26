import base64
import re
import requests
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



# =============================================================================
# Umycs - Generate TPDS documents via Docugenerate
# =============================================================================

# -- CONFIG -------------------------------------------------------------------
DOCUGENERATE_API_KEY = 'PASTE-NEW-DOCUGENERATE-KEY'   # rotate the old key, never commit the real one
TEMPLATE_ID          = '7kvTHZfhZirTHsetkxFS'

# SeaTable
TABLE_NAME           = 'TPDS view'

# Microsoft Graph API
TENANT_ID            = '89163d46-54fe-4ede-95d2-885f4aaf21c4'
CLIENT_ID            = '549aeefa-b3c8-4256-be2b-6da7e71c1883'
CLIENT_SECRET        = 'PASTE-NEW-CLIENT-SECRET'      # rotate the old secret, never commit the real one

# Email settings
FROM_EMAIL           = 'eric.marks@umycs.com'
TO_EMAIL             = ''     # leave empty to skip email

# OneDrive settings
ONEDRIVE_FOLDER      = r'Product Management Data\Technical Product Data Sheets'

# Languages / labels
LABELS_TABLE         = 'Labels'   # columns: Key, EN, DE, ... (one column per language)
DEFAULT_LANG         = 'EN'       # used for rows without a Language and as label fallback
USE_DATE_TPDS        = False      # False: issue date = today; True: use 'Date TPDS' if filled

# Translation memory (all non-label text of the data sheet)
TM_TABLE             = 'Translation memory'  # Source, Field, one column per language, Reviewed languages
ON_MISSING           = 'skip'     # 'skip': do not generate a language with untranslated text
                                  # 'english': generate anyway, untranslated text stays English
USE_DRAFTS           = False      # False: only translations ticked in "Reviewed languages"
                                  # True: also use filled but unreviewed translations
# -----------------------------------------------------------------------------


# -- STEP 1: Connect to SeaTable ----------------------------------------------
base = Base(context.api_token, context.server_url)
base.auth()
print('Connected to SeaTable')


# -- STEP 2: Fetch rows where "Generate document" = true ----------------------
all_rows = base.list_rows(TABLE_NAME)
rows = [row for row in all_rows if row.get('Generate document') == True]
print(f'Found {len(rows)} rows to process')

if len(rows) == 0:
    print('No rows with "Generate document" = true. Done.')
    raise SystemExit


# -- STEP 3: Get Microsoft Graph access token ---------------------------------
print('Getting Microsoft Graph access token...')
token_response = requests.post(
    f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token',
    data={
        'grant_type':    'client_credentials',
        'client_id':     CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope':         'https://graph.microsoft.com/.default'
    }
)

if token_response.status_code != 200:
    print(f'Token error: {token_response.text}')
    raise SystemExit

access_token  = token_response.json()['access_token']
graph_headers = {'Authorization': f'Bearer {access_token}'}
print('Access token obtained')


# -- HELPERS ------------------------------------------------------------------
def s(row, col):
    v = row.get(col)
    if v is None:
        return ''
    if isinstance(v, bool):
        return ''
    if isinstance(v, list):
        parts = []
        for item in v:
            if isinstance(item, dict):
                parts.append(str(item.get('display_value', item.get('name', ''))))
            else:
                parts.append(str(item))
        return ', '.join(p for p in parts if p)
    return str(v)

def table_rows(row, table_num, max_rows):
    result = []
    for i in range(1, max_rows + 1):
        prop = s(row, f'Table {table_num} {i} A')
        val  = s(row, f'Table {table_num} {i} B')
        if prop or val:
            result.append({'Property': prop, 'Value': val})
    return result

def linked_statement_lines(row, col):
    """Get the text of every linked Statements row (col must be a
    multi-link column whose display value is the Statement text), one
    line per linked row, for use with a Docugenerate {#Tag}{.}{/Tag}
    array loop (one paragraph per linked statement)."""
    v = row.get(col)
    if not v or not isinstance(v, list):
        return []
    lines = []
    for item in v:
        if isinstance(item, dict):
            text = item.get('display_value') or item.get('name') or ''
            text = str(text)
        else:
            text = str(item) if item is not None else ''
        text = text.replace('&#x20;', ' ').replace('&#X20;', ' ').strip()
        if text:
            lines.append(text)
    return lines

def split_last(lines):
    """Split a list of lines into (all-but-last, last) so the last line can be
    styled differently (e.g. left-aligned) from the rest (e.g. justified)."""
    if not lines:
        return [], ''
    return lines[:-1], lines[-1]


# [log patch] documents generated in this run
generated = 0


# -- LABELS & DATES -----------------------------------------------------------
# Built-in English labels: used when the Labels table is missing or a cell is empty.
# Keys are the {{Label...}} placeholders of the Docugenerate template.
DEFAULT_LABELS = {
    'LabelDocTitle':              'Technical Product Data Sheet',
    'LabelIssueDate':             'Issue Date:',
    'LabelProductDescription':    'Product Description',
    'LabelApplication':           'Application',
    'LabelComposition':           'Composition',
    'LabelStorage':               'Storage',
    'LabelShelfLife':             'Shelf Life',
    'LabelPackagingSize':         'Packaging Size',
    'LabelGMOInformation':        'GMO Information',
    'LabelComplianceLabelling':   'Compliance & Labelling',
    'LabelHealthSafety':          'Health & Safety',
    'LabelQualityCertifications': 'Quality & Food Safety Certifications',
}

MONTHS = {
    'EN': ['January', 'February', 'March', 'April', 'May', 'June', 'July',
           'August', 'September', 'October', 'November', 'December'],
    'DE': ['Januar', 'Februar', 'M\u00e4rz', 'April', 'Mai', 'Juni', 'Juli',
           'August', 'September', 'Oktober', 'November', 'Dezember'],
}

def load_labels():
    """{lang: {Key: text}} from the Labels table. EN starts from DEFAULT_LABELS."""
    labels = {DEFAULT_LANG: dict(DEFAULT_LABELS)}
    try:
        label_rows = base.list_rows(LABELS_TABLE)
        columns    = [c['name'] for c in base.list_columns(LABELS_TABLE)]
    except Exception as e:
        print(f'WARNING: could not read table "{LABELS_TABLE}", using built-in English labels: {e}')
        return labels
    for lang in columns:
        if lang == 'Key':
            continue
        labels.setdefault(lang, {})
        for r in label_rows:
            key, text = r.get('Key'), r.get(lang)
            if key and text:
                labels[lang][key] = str(text)
    return labels

LABELS = load_labels()
print('Label languages available: ' + ', '.join(sorted(LABELS)))

def labels_for(lang):
    """Labels for one language; anything missing falls back to English."""
    if lang not in LABELS:
        print(f'  WARNING: no label column for "{lang}", using {DEFAULT_LANG} labels')
    return {**LABELS[DEFAULT_LANG], **LABELS.get(lang, {})}

def row_languages(row):
    """Languages selected in the row's multi-select; EN if none."""
    langs = row.get('Language') or []
    if isinstance(langs, (str, dict)):
        langs = [langs]
    out = [str(l.get('name', '')) if isinstance(l, dict) else str(l) for l in langs]
    out = [l for l in out if l]
    return out or [DEFAULT_LANG]

def issue_dates(row, lang):
    """{IssueDate, IssueDateShort} for header/body and footer, per language."""
    d = None
    if USE_DATE_TPDS and row.get('Date TPDS'):
        d = _dt.date.fromisoformat(str(row['Date TPDS'])[:10])
    d = d or _dt.date.today()
    month = MONTHS.get(lang, MONTHS[DEFAULT_LANG])[d.month - 1]
    if lang == 'DE':
        return {'IssueDate': f'{d.day}. {month} {d.year}',
                'IssueDateShort': f'{d.day}.{d.month}.{d.year}'}
    return {'IssueDate': f'{month} {d.day}, {d.year}',
            'IssueDateShort': f'{d.day}-{d.month}-{d.year}'}

def lang_tag(lang):
    """Display-name suffix; none for the default language."""
    return '' if lang == DEFAULT_LANG else f' ({lang})'

def file_tag(lang):
    """File-name suffix; none for the default language (keeps existing names)."""
    return '' if lang == DEFAULT_LANG else f'_{lang}'


# -- TRANSLATION MEMORY -------------------------------------------------------
# One row per distinct English text: Source, Field, one column per language
# (DE, FR, ...) and "Reviewed languages". The product code inside a text is
# stored as {P} so that boilerplate shared by many products is translated once.
def norm(text):
    """Lookup key: &#x20; -> space, any run of whitespace -> one space."""
    return ' '.join(str(text).replace('&#x20;', ' ').replace('&#X20;', ' ').split())

def needs_translation(text):
    """False for empty text and for pure numbers/units (no word of 3+ letters)."""
    return bool(re.search(r'[^\W\d_]{3,}', text))

def load_tm():
    """{normalised source: TM row}; reads the table in pages of 1000."""
    tm, start = {}, 0
    try:
        while True:
            chunk = base.list_rows(TM_TABLE, start=start, limit=1000)
            for r in chunk:
                if r.get('Source'):
                    tm[norm(r['Source'])] = r
            if len(chunk) < 1000:
                break
            start += 1000
    except Exception as e:
        print(f'WARNING: could not read table "{TM_TABLE}": {e}')
    return tm

TM = load_tm()
print(f'Translation memory: {len(TM)} entries')
NEW_SOURCES = {}   # normalised source -> field, for texts the TM has not seen yet

class Translator:
    """Translates the texts of one document into one language via the TM."""
    def __init__(self, lang, product_name):
        self.lang    = lang
        parts        = re.sub(r'^\s*Nutrycs\s*[®™]?\s*', '', product_name).split()
        self.short   = ' '.join(parts)                      # 'Nutrycs®  TGT 101' -> 'TGT 101'
        self.pattern = re.compile(r'\s+'.join(re.escape(p) for p in parts)) if parts else None
        self.missing = {}    # sources without a usable translation -> field
        self.drafts  = 0     # unreviewed translations used (USE_DRAFTS)

    def text(self, value, field):
        if not value or not needs_translation(value):
            return value
        src = norm(value)
        if self.pattern:
            src = self.pattern.sub('{P}', src)
        entry = TM.get(src)
        if entry is None:
            NEW_SOURCES.setdefault(src, field)
            self.missing[src] = field
            return value
        out      = entry.get(self.lang)
        reviewed = self.lang in (entry.get('Reviewed languages') or [])
        if not out or not (reviewed or USE_DRAFTS):
            self.missing[src] = field
            return value
        if not reviewed:
            self.drafts += 1
        return str(out).replace('{P}', self.short)

TEXT_KEYS = {
    'ProductDescription': 'Product description', 'Application': 'Application',
    'Composition': 'Composition', 'Storage': 'Storage', 'ShelfLife': 'Shelf life',
    'PackagingSize': 'Packaging size', 'GMOInformation': 'GMO information',
    'GMOadditionalnotes': 'GMO additional information', 'DisclaimerLastLine': 'Disclaimer',
}
LINES_KEYS = {
    'ComplianceLabellingLines': 'Compliance & Labelling', 'HealthSafetyLines': 'Health & Safety',
    'QualityFoodSafetyCertificationsLines': 'Quality & Food Safety Certifications',
    'DisclaimerJustifiedLines': 'Disclaimer',
}
for _n in range(1, 5):
    for _suffix, _label in (('Title', 'title'), ('HeaderA', 'header A'),
                            ('HeaderB', 'header B'), ('Additionalnotes', 'additional notes')):
        TEXT_KEYS[f'Table{_n}{_suffix}'] = f'Table {_n} {_label}'

def translate_payload(payload, tr):
    """Translate the data-sheet texts of a payload in place (labels are handled separately)."""
    for key, field in TEXT_KEYS.items():
        payload[key] = tr.text(payload.get(key, ''), field)
    for key, field in LINES_KEYS.items():
        payload[key] = [tr.text(v, field) for v in payload.get(key, [])]
    for n in range(1, 5):
        payload[f'Table{n}'] = [
            {'Property': tr.text(r['Property'], f'Table {n} property'),
             'Value':    tr.text(r['Value'],    f'Table {n} value')}
            for r in payload.get(f'Table{n}', [])
        ]

def flush_new_sources():
    """Add the texts the TM has not seen yet, so they can be translated."""
    if not NEW_SOURCES:
        return
    new_rows = [{'Source': src, 'Field': field} for src, field in NEW_SOURCES.items()]
    try:
        for i in range(0, len(new_rows), 500):
            base.batch_append_rows(TM_TABLE, new_rows[i:i + 500])
        print(f'[tm] {len(new_rows)} new text(s) added to "{TM_TABLE}"')
    except Exception as e:
        print(f'[tm] WARNING: could not write to "{TM_TABLE}": {e}')


# -- STEP 4: Process each row -------------------------------------------------
for row in rows:
    product_name = s(row, 'Product name')
    row_id       = row.get('_id')
    languages    = row_languages(row)
    print(f'\nProcessing: {product_name} ({", ".join(languages)})')
    all_ok       = True

    disclaimer_justified, disclaimer_last = split_last(linked_statement_lines(row, 'Disclaimer text'))

    for lang in languages:
        # Build Docugenerate payload
        payload_data = {
            'ProductName':                     s(row, 'Product name'),
            'ProductDescription':              s(row, 'Product description'),
            'Application':                     s(row, 'Application'),
            'Composition':                     s(row, 'Composition'),
            'Storage':                         s(row, 'Storage'),
            'ShelfLife':                       s(row, 'Shelf Life'),
            'PackagingSize':                   s(row, 'Packaging Size'),
            'GMOInformation':                  s(row, 'GMO Information'),
            'GMOadditionalnotes':              s(row, 'GMO Additional information'),
            'ComplianceLabellingLines':        linked_statement_lines(row, 'Compliance & Labelling'),
            'HasComplianceLabelling':          bool(linked_statement_lines(row, 'Compliance & Labelling')),
            'HealthSafetyLines':               linked_statement_lines(row, 'Health & Safety'),
            'HasHealthSafety':                 bool(linked_statement_lines(row, 'Health & Safety')),
            'QualityFoodSafetyCertificationsLines': linked_statement_lines(row, 'Quality & Food Safety Certifications'),
            'HasQualityFoodSafetyCertifications':   bool(linked_statement_lines(row, 'Quality & Food Safety Certifications')),
            'DisclaimerJustifiedLines':        disclaimer_justified,
            'DisclaimerLastLine':              disclaimer_last,
            'HasDisclaimer':                   bool(disclaimer_justified or disclaimer_last),
            'Table1Title':                     s(row, 'Table 1 Title'),
            'Table1HeaderA':                   s(row, 'Table 1 Header A'),
            'Table1HeaderB':                   s(row, 'Table 1 Header B'),
            'Table1':                          table_rows(row, 1, 6),
            'Table1Additionalnotes':           s(row, 'Table 1 Additional notes'),
            'Table2Title':                     s(row, 'Table 2 Title'),
            'Table2HeaderA':                   s(row, 'Table 2 Header A'),
            'Table2HeaderB':                   s(row, 'Table 2 Header B'),
            'Table2':                          table_rows(row, 2, 4),
            'Table2Additionalnotes':           s(row, 'Table 2 Additional notes'),
            'Table3Title':                     s(row, 'Table 3 Title'),
            'Table3HeaderA':                   s(row, 'Table 3 Header A'),
            'Table3HeaderB':                   s(row, 'Table 3 Header B'),
            'Table3':                          table_rows(row, 3, 13),
            'Table3Additionalnotes':           s(row, 'Table 3 Additional notes'),
            'Table4Title':                     s(row, 'Table 4 Title'),
            'Table4HeaderA':                   s(row, 'Table 4 Header A'),
            'Table4HeaderB':                   s(row, 'Table 4 Header B'),
            'Table4':                          table_rows(row, 4, 14),
            'Table4Additionalnotes':           s(row, 'Table 4 Additional notes'),
        }
        if lang != DEFAULT_LANG:
            tr = Translator(lang, product_name)
            translate_payload(payload_data, tr)
            if tr.drafts:
                print(f'  NOTE: {tr.drafts} unreviewed (draft) translation(s) used [{lang}]')
            if tr.missing:
                print(f'  {len(tr.missing)} untranslated or unreviewed text(s) [{lang}]')
                if ON_MISSING == 'skip':
                    print(f'  Skipping {lang}: fill the {lang} column in "{TM_TABLE}", tick {lang} under '
                          f'"Reviewed languages", then run again')
                    all_ok = False
                    continue
        payload_data.update(labels_for(lang))
        payload_data.update(issue_dates(row, lang))

        # -- STEP 4a: Generate PDF via Docugenerate -------------------------------
        dg_response = requests.post(
            'https://api.docugenerate.com/v1/document',
            headers={
                'Authorization': DOCUGENERATE_API_KEY,
                'Content-Type':  'application/json'
            },
            json={
                'template_id':   TEMPLATE_ID,
                'name':          f'TPDS - Nutrycs® {product_name}{lang_tag(lang)}',
                'output_format': '.pdf',
                'data':          [payload_data]
            }
        )

        if dg_response.status_code not in [200, 201]:
            print(f'  Docugenerate error ({lang}): {dg_response.status_code} {dg_response.text}')
            all_ok = False
            continue

        generated += 1   # [log patch] DocuGenerate accepted the document
        pdf_url     = dg_response.json().get('document_uri')
        pdf_content = requests.get(pdf_url).content
        pdf_name    = f'TPDS_Nutrycs_{product_name.replace(" ", "_")}{file_tag(lang)}.pdf'
        print(f'  PDF generated [{lang}] ({len(pdf_content)} bytes)')

        # -- STEP 4b: Save PDF to OneDrive ----------------------------------------
        upload_response = requests.put(
            f'https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/drive/root:/{ONEDRIVE_FOLDER}/{pdf_name}:/content',
            headers={**graph_headers, 'Content-Type': 'application/pdf'},
            data=pdf_content
        )
        if upload_response.status_code in [200, 201]:
            print(f'  Saved to OneDrive')
        else:
            print(f'  OneDrive warning: {upload_response.status_code} {upload_response.text}')

        # -- STEP 4c: Send email if TO_EMAIL is set --------------------------------
        if TO_EMAIL:
            email_payload = {
                'message': {
                    'subject': f'TPDS - Nutrycs® {product_name}{lang_tag(lang)}',
                    'body': {
                        'contentType': 'Text',
                        'content': (
                            f'Dear colleague,\n\n'
                            f'Please find attached the Technical Product Data Sheet for '
                            f'Nutrycs® {product_name}.\n\n'
                            f'Best regards,\nUmycs'
                        )
                    },
                    'toRecipients': [{'emailAddress': {'address': TO_EMAIL}}],
                    'attachments': [{
                        '@odata.type':  '#microsoft.graph.fileAttachment',
                        'name':         pdf_name,
                        'contentType':  'application/pdf',
                        'contentBytes': base64.b64encode(pdf_content).decode('utf-8')
                    }]
                },
                'saveToSentItems': True
            }
            email_response = requests.post(
                f'https://graph.microsoft.com/v1.0/users/{FROM_EMAIL}/sendMail',
                headers={**graph_headers, 'Content-Type': 'application/json'},
                json=email_payload
            )
            if email_response.status_code == 202:
                print(f'  Email sent to {TO_EMAIL}')
            else:
                print(f'  Email error: {email_response.status_code} {email_response.text}')

    # -- STEP 4d: Update row status in SeaTable --------------------------------
    if all_ok:
        base.update_row(TABLE_NAME, row_id, {
            'Status':            'Generated',
            'Generate document': False
        })
        print(f'  Status updated to "Generated"')
    else:
        print('  WARNING: not every language was generated - row left ticked for a re-run')


flush_new_sources()

# [log patch]
log_generation(base, 'TPDS', generated)

print('\nAll done!')