import os
import sqlite3
import struct
import re
import zipfile
import time
from io import BytesIO
from threading import Thread

from flask import Flask, render_template, request, jsonify, send_file

app = Flask(__name__)

TEMP_DB = "/tmp/shadow_library.eni"
DB_STATUS = {"ready": False, "message": "upload_needed", "last_sync": None}

# --- ONEDRIVE / AZURE CONFIG (set as Render environment variables) ---
AZURE_TENANT_ID    = os.environ.get('AZURE_TENANT_ID', '')
AZURE_CLIENT_ID    = os.environ.get('AZURE_CLIENT_ID', '')
AZURE_CLIENT_SECRET = os.environ.get('AZURE_CLIENT_SECRET', '')
ONEDRIVE_GROUP_NAME = os.environ.get('ONEDRIVE_GROUP_NAME', 'GRP-IWTSC')
ONEDRIVE_FILE_PATH  = os.environ.get('ONEDRIVE_FILE_PATH',
    'EndNote Resource Library/IWTSC Endnote Library.Data/sdb/sdb.eni')

REF_TYPE_MAP = {
    'Journal Article': 17, 'Video': 3, 'Book': 6, 'Book Section': 5,
    'Web Page': 31, 'Thesis': 32, 'Conference Paper': 47, 'Report': 27
}


# ---------------------------------------------------------------------------
# OneDrive sync via Microsoft Graph API
# ---------------------------------------------------------------------------

def get_graph_token():
    import msal
    authority = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}"
    msal_app = msal.ConfidentialClientApplication(
        AZURE_CLIENT_ID,
        authority=authority,
        client_credential=AZURE_CLIENT_SECRET,
    )
    result = msal_app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if 'access_token' not in result:
        raise Exception(result.get('error_description', 'Token acquisition failed'))
    return result['access_token']


def sync_from_onedrive():
    global DB_STATUS
    if not all([AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET]):
        DB_STATUS = {
            "ready": False,
            "message": "Azure credentials not configured. "
                       "Set AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET in Render.",
            "last_sync": None,
        }
        return False

    try:
        import requests
        token = get_graph_token()
        headers = {'Authorization': f'Bearer {token}'}

        # Find the Microsoft 365 group
        resp = requests.get(
            f"https://graph.microsoft.com/v1.0/groups"
            f"?$filter=startswith(displayName,'{ONEDRIVE_GROUP_NAME}')&$select=id,displayName",
            headers=headers, timeout=30,
        )
        groups = resp.json().get('value', [])
        if not groups:
            DB_STATUS = {
                "ready": False,
                "message": f"Microsoft 365 group '{ONEDRIVE_GROUP_NAME}' not found.",
                "last_sync": None,
            }
            return False

        group_id = groups[0]['id']

        # Download the .eni database file
        file_url = (
            f"https://graph.microsoft.com/v1.0/groups/{group_id}"
            f"/drive/root:/{ONEDRIVE_FILE_PATH}:/content"
        )
        r = requests.get(file_url, headers=headers, timeout=120)
        if r.status_code != 200:
            DB_STATUS = {
                "ready": False,
                "message": f"File not found in OneDrive (HTTP {r.status_code}). "
                           f"Path: {ONEDRIVE_FILE_PATH}",
                "last_sync": None,
            }
            return False

        with open(TEMP_DB, 'wb') as f:
            f.write(r.content)

        DB_STATUS = {
            "ready": True,
            "message": "Database synced from OneDrive.",
            "last_sync": time.strftime('%Y-%m-%d %H:%M UTC'),
        }
        return True

    except Exception as e:
        DB_STATUS = {"ready": False, "message": f"Sync error: {e}", "last_sync": None}
        return False


# ---------------------------------------------------------------------------
# Reference formatting helpers (unchanged from desktop)
# ---------------------------------------------------------------------------

def format_authors(raw_author_string, format_type):
    if not raw_author_string:
        return "Unknown Author"
    authors = [a.strip() for a in raw_author_string.replace('\r', '\n').split('\n') if a.strip()]
    formatted = []
    for a in authors:
        parts = a.split(',')
        if len(parts) == 2:
            last, first = parts[0].strip(), parts[1].strip().split()
            initials = (
                "".join([n[0] + "." for n in first if n])
                if format_type == 'APA'
                else "".join([n[0] for n in first if n])
            )
            formatted.append(f"{last}, {initials}" if format_type == 'APA' else f"{last} {initials}")
        else:
            formatted.append(a)
    if format_type == 'APA':
        return (
            ", ".join(formatted[:-1]) + ", & " + formatted[-1]
            if len(formatted) > 1
            else (formatted[0] if formatted else "")
        )
    return ", ".join(formatted)


def build_search_query(filters, for_excel=False):
    cols = "id, author, year, title, reference_type, secondary_title, volume, number, pages, electronic_resource_number, url, abstract"
    if for_excel:
        cols = "author, year, title, reference_type, secondary_title, volume, number, pages, date, access_date, isbn, electronic_resource_number, notes, label, abstract, accession_number, url, author_address, publisher"
    query = f"SELECT {cols} FROM refs WHERE trash_state = 0"
    params = []

    primary_filters = [f for f in filters if f['type'] == 'Primary' and f['value'] != 'All']
    secondaries = [f for f in filters if f['type'] == 'Secondary']

    if primary_filters:
        conn = sqlite3.connect(TEMP_DB)
        cursor = conn.cursor()
        cursor.execute("SELECT members, spec FROM groups WHERE spec IS NOT NULL")
        valid_ids = set()
        target_names = [f['value'] for f in primary_filters]
        for row in cursor.fetchall():
            blob, spec = row[0], row[1]
            spec_str = spec.decode('utf-8', 'ignore') if isinstance(spec, bytes) else str(spec)
            match = re.search(r'<name>(.*?)</name>', spec_str, re.IGNORECASE)
            if match and match.group(1).strip() in target_names and blob:
                count = len(blob) // 4
                ints = struct.unpack(f'<{count}I', blob[:count * 4])
                valid_ids.update([str(i) for i in ints[1:]])
        conn.close()
        query += f" AND id IN ({','.join(valid_ids)})" if valid_ids else " AND 1=0"

    col_map = {
        'Keyword': 'keywords', 'Author': 'author', 'Primary Author': 'author',
        'Year': 'year', 'Journal': 'secondary_title', 'Title': 'title',
        'Reference Type': 'reference_type',
    }
    for f in secondaries:
        col = col_map.get(f['category'])
        if not col:
            continue
        if f['category'] == 'Reference Type':
            query += f" AND {col} = ?"
            params.append(REF_TYPE_MAP.get(f['value'], 17))
        else:
            query += f" AND {col} LIKE ?"
            params.append(f"%{f['value']}%")
    return query, params


# ---------------------------------------------------------------------------
# Flask routes
# ---------------------------------------------------------------------------

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/status')
def status():
    return jsonify(DB_STATUS)


@app.route('/api/sync', methods=['POST'])
def manual_sync():
    Thread(target=sync_from_onedrive, daemon=True).start()
    return jsonify({"message": "Sync started..."})


@app.route('/api/upload', methods=['POST'])
def upload_library():
    if 'file' not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.eni'):
        return jsonify({"error": "Please upload a .eni file (the library database file)"}), 400
    try:
        f.save(TEMP_DB)
        DB_STATUS["ready"] = True
        DB_STATUS["message"] = "Library uploaded successfully."
        DB_STATUS["last_sync"] = time.strftime('%Y-%m-%d %H:%M UTC')
        return jsonify({"success": True})
    except Exception as e:
        DB_STATUS["ready"] = False
        DB_STATUS["message"] = f"Upload error: {e}"
        return jsonify({"error": str(e)}), 500


@app.route('/api/get_filters')
def get_filters():
    if not DB_STATUS['ready']:
        return jsonify({"groups": [], "error": DB_STATUS['message']})
    try:
        conn = sqlite3.connect(TEMP_DB)
        cursor = conn.cursor()
        cursor.execute("SELECT spec FROM groups WHERE spec IS NOT NULL")
        unique_groups = set()
        blacklist = ['Found PDF', 'Found URL', 'Not Found', 'Searching',
                     'Authentication required', 'URL', 'Recent Search']
        for g in cursor.fetchall():
            raw = g[0].decode('utf-8', 'ignore') if isinstance(g[0], bytes) else str(g[0])
            match = re.search(r'<name>(.*?)</name>', raw, re.IGNORECASE)
            if match:
                name = match.group(1).strip()
                if any(b in name for b in blacklist):
                    continue
                if len(name) > 55:
                    continue
                if 'TYPE;0' in raw or 'TYPE;4' in raw:
                    unique_groups.add(name)
        conn.close()
        return jsonify({"groups": sorted(list(unique_groups))})
    except Exception as e:
        return jsonify({"groups": [], "error": str(e)})


@app.route('/api/generate', methods=['POST'])
def generate_refs():
    if not DB_STATUS['ready']:
        return jsonify({"html": f"<em>{DB_STATUS.get('message', 'Library not loaded')}</em>", "count": 0})
    data = request.json
    format_type = data.get('format', 'APA')
    filters = data.get('filters', [])
    try:
        query, params = build_search_query(filters)
        conn = sqlite3.connect(TEMP_DB)
        cursor = conn.cursor()
        cursor.execute(query, params)
        results = cursor.fetchall()
        conn.close()
        html_output = []
        for r in results:
            if format_type == 'Pure Links Only':
                raw_url = r[10] or ""
                raw_doi = r[9] or ""
                links = [u.strip() for u in raw_url.replace('\r', '\n').split('\n')
                         if u.strip().startswith('http')]
                if not links and raw_doi:
                    links.append(f"https://doi.org/{raw_doi}")
                for l in links:
                    html_output.append(f"<p><a href='{l}' target='_blank'>{l}</a></p>")
            else:
                authors = format_authors(r[1], format_type)
                year = r[2] or "n.d."
                title = r[3] or "Untitled"
                journal = r[5] or ""
                html_output.append(
                    f"<p style='margin-bottom:12px;'>{authors} ({year}). {title}. <i>{journal}</i>.</p>"
                )
        return jsonify({
            "html": "".join(html_output) if html_output else "<em>No matches found.</em>",
            "count": len(html_output),
        })
    except Exception as e:
        return jsonify({"html": f"<em>Error: {e}</em>", "count": 0})


@app.route('/api/export', methods=['POST'])
def export_excel():
    if not DB_STATUS['ready']:
        return jsonify({"error": DB_STATUS['message']}), 503
    import pandas as pd
    data = request.json
    filters = data.get('filters', [])
    try:
        query, params = build_search_query(filters, for_excel=True)
        conn = sqlite3.connect(TEMP_DB)
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        df.columns = [
            'Authors', 'Year', 'Title', 'Reference Type', 'Journal', 'Volume', 'Issue',
            'Pages', 'Date Published', 'Date Accessed', 'ISSN', 'DOI', 'Notes', 'Label',
            'Abstract', 'PMID/PMCID', 'Link', 'Affiliation', 'Publisher',
        ]
        rev_map = {v: k for k, v in REF_TYPE_MAP.items()}
        df['Reference Type'] = df['Reference Type'].map(rev_map).fillna('Other')
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='References')
        output.seek(0)
        return send_file(output, download_name="Lanzi_Lab_References.xlsx", as_attachment=True)
    except Exception as e:
        return jsonify({"error": f"Export failed: {e}"}), 500


@app.route('/api/download_pdfs', methods=['POST'])
def download_pdfs():
    # PDF files live in OneDrive locally; not available in the web version.
    return jsonify({"error": "pdf_unavailable"}), 501


# ---------------------------------------------------------------------------
# Auto-sync on startup (background thread so Flask starts immediately)
# ---------------------------------------------------------------------------

def _startup_sync():
    time.sleep(3)
    # Skip sync if Azure credentials are not real
    if AZURE_TENANT_ID in ('', 'placeholder') or AZURE_CLIENT_ID in ('', 'placeholder'):
        return  # Waiting for manual file upload
    sync_from_onedrive()

Thread(target=_startup_sync, daemon=True).start()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
