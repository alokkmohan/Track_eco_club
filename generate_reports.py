import pandas as pd
import io
import os
import json
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import openpyxl

os.makedirs("reports", exist_ok=True)

DISTRICTS = {
    "Lakhimpur Kheri":  ("KHERI",       "KHERI",          True),
    "Auraiya":          ("AURAIYA",      "AURAIYA",        True),
    "Chandauli":        ("CHANDAULI",    "CHANDAULI",      True),
    "Bulandshahr":      ("BULANDSHAHR",  "BULANDSHAHR",    True),
    "Hathras":          ("HATHRAS",      "HATHRAS",        True),
    "Shravasti":        ("SHRAV|SHRAW",  "SHRAV|SHRAW",    False),
    "Gonda":            ("GONDA",        "GONDA",          True),
    "Basti":            ("BASTI",        "BASTI",          True),
    "Bareilly":         ("BAREILLY",     "BAREILLY",       True),
    "Ballia":           ("BALLIA",       "BALLIA",         True),
    "Rampur":           ("RAMPUR",       "RAMPUR",         True),
    "Moradabad":        ("MORADABAD",    "MORADABAD",      True),
    "Etawah":           ("ETAWAH",       "ETAWAH",         True),
    "Azamgarh":         ("AZAMGARH",     "AZAMGARH",       True),
    "Amethi":           ("AMETHI",       "AMETHI",         False),
    "Kannauj":          ("KANNAUJ",      "KANNAUJ",        True),
    "Jaunpur":          ("JAUNPUR",      "JAUNPUR",        True),
    "Hardoi":           ("HARDOI",       "HARDOI",         True),
    "Hapur":            ("HAPUR",        "HAPUR",          False),
    "Amroha":           ("AMROHA",       "AMROHA|JYOTIBA", False),
    "Bhadohi":          ("BHADOHI",      "BHADOI",         False),
    "Prayagraj":        ("PRAYAGRAJ",    "PRAYAGRAJ",      True),
    "Sambhal":          ("SAMBHAL",      "SAMBHAL",        False),
    "Mirzapur":         ("MIRZAPUR",     "MIRZAPUR",       True),
    "Meerut":           ("MEERUT",       "MEERUT",         True),
    "Farrukhabad":      ("FARRUKHABAD",  "FARRUKHABAD",    True),
}

def normalize_udise(val):
    s = str(val).strip().split('.')[0].lstrip('0')
    return s.zfill(10)

def load_secondary_list():
    return {
        sheet: pd.read_excel("data/Final_Secondary_School_List.xlsx", sheet_name=sheet)
        for sheet in ["Govt Schools", "Aided Schools", "Private Schools"]
    }

def load_notifications():
    return pd.read_excel("data/All_Schools_with_Notifications_UP.xlsx")

def get_district_data(notifications_df, secondary_data, dist_label):
    sec_pattern, notif_pattern, exact = DISTRICTS[dist_label]

    if exact:
        notif_d = notifications_df[notifications_df['District'].str.strip().str.upper() == notif_pattern].copy()
    else:
        notif_d = notifications_df[notifications_df['District'].str.strip().str.upper().str.contains(notif_pattern, regex=True, na=False)].copy()

    notif_d['UDISE_norm'] = notif_d['UDISE ID'].apply(normalize_udise)
    notif_set = set(notif_d['UDISE_norm'])

    result = {}
    for sheet, df in secondary_data.items():
        dist_col  = [c for c in df.columns if 'district' in c.lower()][0]
        udise_col = [c for c in df.columns if 'udise' in c.lower()][0]

        if exact:
            d = df[df[dist_col].astype(str).str.strip().str.upper() == sec_pattern].copy()
        else:
            d = df[df[dist_col].astype(str).str.strip().str.upper().str.contains(sec_pattern, regex=True, na=False)].copy()

        d[dist_col]  = dist_label
        d['UDISE_norm'] = d[udise_col].apply(normalize_udise)
        d[udise_col] = d['UDISE_norm']
        d['Upload Status'] = d['UDISE_norm'].apply(lambda x: 'Uploaded' if x in notif_set else 'Pending')
        d = d.drop(columns=['UDISE_norm'])
        result[sheet] = d

    return result

def build_excel(district_data):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    for sheet, df in district_data.items():
        df_out = df.reset_index(drop=True)
        df_out.index = df_out.index + 1
        df_out.index.name = 'S.No.'
        df_out.to_excel(writer, sheet_name=sheet, index=True)

    writer.close()
    output.seek(0)

    wb = openpyxl.load_workbook(output)
    green_fill  = PatternFill('solid', fgColor='C6EFCE')
    red_fill    = PatternFill('solid', fgColor='FFC7CE')
    header_fill = PatternFill('solid', fgColor='1F4E79')
    header_font = Font(color='FFFFFF', bold=True)
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),  bottom=Side(style='thin'))

    for ws in wb.worksheets:
        status_col    = None
        udise_col_idx = None
        for cell in ws[1]:
            if cell.value == 'Upload Status':
                status_col = cell.column
            if cell.value and 'udise' in str(cell.value).lower():
                udise_col_idx = cell.column
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = thin
                cell.alignment = Alignment(vertical='center')
                if udise_col_idx and cell.column == udise_col_idx:
                    cell.number_format = '@'
                    if cell.value:
                        cell.value = str(cell.value).zfill(10)
            if status_col:
                sc = ws.cell(row=row[0].row, column=status_col)
                if sc.value == 'Uploaded':
                    sc.fill = green_fill
                    sc.font = Font(bold=True, color='276221')
                elif sc.value == 'Pending':
                    sc.fill = red_fill
                    sc.font = Font(bold=True, color='9C0006')
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_len = max(max_len, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 4, 45)
        ws.freeze_panes = 'B2'

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final

def generate_index(summary, generated_time):
    total_grand    = sum(sum(s[sh]["total"]    for sh in ["Govt Schools","Aided Schools","Private Schools"]) for s in summary)
    uploaded_grand = sum(sum(s[sh]["uploaded"] for sh in ["Govt Schools","Aided Schools","Private Schools"]) for s in summary)
    pending_grand  = total_grand - uploaded_grand
    pct_grand      = round(uploaded_grand / total_grand * 100) if total_grand > 0 else 0

    cards = ""
    for s in sorted(summary, key=lambda x: x["district"]):
        dist = s["district"]
        priv_missing = s["Private Schools"]["total"] == 0
        total_d    = sum(s[sh]["total"]    for sh in ["Govt Schools","Aided Schools","Private Schools"])
        uploaded_d = sum(s[sh]["uploaded"] for sh in ["Govt Schools","Aided Schools","Private Schools"])
        pending_d  = total_d - uploaded_d
        pct        = round(uploaded_d / total_d * 100) if total_d > 0 else 0
        bar_color  = "bg-success" if pct >= 75 else "bg-warning" if pct >= 40 else "bg-danger"
        fname      = f"reports/{dist}_ECO_Club_Status.xlsx"

        rows = ""
        for sh, icon in [("Govt Schools","🏛️"), ("Aided Schools","🤝"), ("Private Schools","🏫")]:
            t = s[sh]["total"]; u = s[sh]["uploaded"]; p = s[sh]["pending"]
            if sh == "Private Schools" and priv_missing:
                rows += f"""<tr class="table-warning">
              <td>{icon} Private</td>
              <td colspan="3" class="text-center text-warning fw-bold small">⚠️ List pending</td>
            </tr>"""
            else:
                rows += f"""<tr>
              <td>{icon} {sh.replace(' Schools','')}</td>
              <td class="text-center">{t}</td>
              <td class="text-center fw-bold text-success">{u}</td>
              <td class="text-center text-danger">{p}</td>
            </tr>"""

        priv_badge = ' <span class="badge bg-warning text-dark ms-1" title="Private school list not yet added">⚠️ Pvt</span>' if priv_missing else ""

        cards += f"""
        <div class="col-xl-4 col-md-6 mb-4 district-card" data-name="{dist.lower()}">
          <div class="card h-100 shadow-sm border-0">
            <div class="card-header text-white fw-bold d-flex justify-content-between align-items-center" style="background:#1F4E79;border-radius:12px 12px 0 0">
              <span>📍 {dist}</span>{priv_badge}
            </div>
            <div class="card-body pb-2">
              <div class="d-flex justify-content-between mb-1">
                <small class="text-muted">{uploaded_d} / {total_d} uploaded</small>
                <small class="fw-bold">{pct}%</small>
              </div>
              <div class="progress mb-3" style="height:8px;border-radius:4px">
                <div class="{bar_color} progress-bar" style="width:{pct}%"></div>
              </div>
              <table class="table table-sm table-borderless mb-0 small">
                <thead><tr>
                  <th></th>
                  <th class="text-center text-muted">Total</th>
                  <th class="text-center text-success">✅</th>
                  <th class="text-center text-danger">❌</th>
                </tr></thead>
                <tbody>{rows}</tbody>
              </table>
            </div>
            <div class="card-footer bg-white border-0 pt-0">
              <a href="{fname}" class="btn btn-success btn-sm w-100 fw-bold" download>
                ⬇️ Report Download करें
              </a>
            </div>
          </div>
        </div>"""

    html = f"""<!DOCTYPE html>
<html lang="hi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>ECO Club — District Reports</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
  <style>
    body {{ background:#f0f4f8; font-family: 'Segoe UI', sans-serif; }}
    .hero {{ background: linear-gradient(135deg, #0d3320, #1a6644); color:#fff; padding:2.5rem 0 2rem; }}
    .stat-box {{ background:rgba(255,255,255,.15); border-radius:10px; padding:.8rem 1.4rem; min-width:110px; }}
    .card {{ border-radius:12px; }}
    #searchBox {{ border-radius:25px; padding-left:1.2rem; max-width:420px; }}
    .progress {{ background:#dee2e6; }}
  </style>
</head>
<body>
<div class="hero">
  <div class="container text-center">
    <h1 class="fw-bold mb-1">🌿 ECO Club</h1>
    <p class="lead mb-3">UP Board Secondary Schools — Notification Upload Status</p>
    <div class="d-flex flex-wrap justify-content-center gap-3 mb-3">
      <div class="stat-box"><div class="fs-4 fw-bold">{total_grand}</div><div class="small">Total Schools</div></div>
      <div class="stat-box" style="background:rgba(40,167,69,.5)"><div class="fs-4 fw-bold">{uploaded_grand}</div><div class="small">Uploaded ✅</div></div>
      <div class="stat-box" style="background:rgba(220,53,69,.5)"><div class="fs-4 fw-bold">{pending_grand}</div><div class="small">Pending ❌</div></div>
      <div class="stat-box"><div class="fs-4 fw-bold">{pct_grand}%</div><div class="small">Completion</div></div>
    </div>
    <div class="small opacity-75">Last updated: {generated_time}</div>
  </div>
</div>

<div class="container py-4">
  <div class="d-flex justify-content-center mb-4">
    <input id="searchBox" type="text" class="form-control form-control-lg shadow-sm"
           placeholder="🔍 अपना जिला खोजें..." oninput="filterDistricts(this.value)">
  </div>
  <div class="row" id="grid">{cards}</div>
  <p id="noResult" class="text-center text-muted d-none mt-4">कोई जिला नहीं मिला।</p>
</div>

<footer class="text-center py-3 text-muted small">
  ECO Club — UP Board Secondary Schools &nbsp;|&nbsp; Data auto-updated on push
</footer>

<script>
function filterDistricts(q) {{
  q = q.toLowerCase();
  let visible = 0;
  document.querySelectorAll('.district-card').forEach(c => {{
    const show = c.dataset.name.includes(q);
    c.style.display = show ? '' : 'none';
    if (show) visible++;
  }});
  document.getElementById('noResult').classList.toggle('d-none', visible > 0);
}}
</script>
</body>
</html>"""

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)

def main():
    print("Loading data files...")
    notifications_df = load_notifications()
    secondary_data   = load_secondary_list()

    summary = []
    for dist_label in sorted(DISTRICTS.keys()):
        print(f"  Generating: {dist_label}...")
        district_data = get_district_data(notifications_df, secondary_data, dist_label)
        excel_bytes   = build_excel(district_data)

        fname = f"reports/{dist_label}_ECO_Club_Status.xlsx"
        with open(fname, "wb") as f:
            f.write(excel_bytes.read())

        stats = {"district": dist_label}
        for sheet in ["Govt Schools", "Aided Schools", "Private Schools"]:
            df = district_data[sheet]
            total    = len(df)
            uploaded = int((df['Upload Status'] == 'Uploaded').sum())
            stats[sheet] = {"total": total, "uploaded": uploaded, "pending": total - uploaded}
        summary.append(stats)

    generated_time = datetime.utcnow().strftime("%d %b %Y, %H:%M UTC")
    generate_index(summary, generated_time)
    print(f"\nDone! {len(DISTRICTS)} districts | index.html generated.")

if __name__ == "__main__":
    main()
