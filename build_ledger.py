# -*- coding: utf-8 -*-
import os
import sys
import json
import datetime
import glob
import re
import zipfile
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_PATH = os.path.join(BASE_DIR, 'ledger_data.json')
XLSX_PATH = os.path.join(BASE_DIR, '오토콜 회비 장부.xlsx')
HTML_PATH = os.path.join(BASE_DIR, 'index.html')

def parse_excel_transactions(xlsx_path):
    if not os.path.exists(xlsx_path):
        return None
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for si in ss_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                    t_parts = [t.text for t in si.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') if t.text]
                    shared_strings.append(''.join(t_parts))

            sheet_root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
            rows_data = {}
            for row in sheet_root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                r_num = int(row.get('r'))
                row_dict = {}
                for c in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    ref = c.get('r')
                    col_letters = ''.join([ch for ch in ref if ch.isalpha()])
                    t = c.get('t')
                    v = c.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                    val = v.text if v is not None and v.text is not None else ''
                    if t == 's' and val.isdigit():
                        val = shared_strings[int(val)]
                    row_dict[col_letters] = val
                rows_data[r_num] = row_dict

        txs = []
        for r_num in sorted(rows_data.keys()):
            if r_num < 4:
                continue
            rd = rows_data[r_num]
            seq = rd.get('B', '').strip()
            month = rd.get('C', '').strip()
            extra = rd.get('D', '').strip()
            person = rd.get('E', '').strip()
            amount_str = rd.get('F', '').strip()
            date = rd.get('G', '').strip()
            note = rd.get('H', '').strip()

            if seq and seq.isdigit() and (person or amount_str):
                try:
                    amount = int(float(amount_str))
                except ValueError:
                    amount = 0
                if person or amount != 0:
                    txs.append({
                        'id': int(seq),
                        'month': month,
                        'extraMonth': extra,
                        'person': person,
                        'amount': amount,
                        'date': date,
                        'note': note
                    })
        return txs
    except Exception as e:
        print(f'⚠️ [안내] 엑셀 파싱 참고: {e}')
        return None

def main():
    if not os.path.exists(JSON_PATH):
        print(f'❌ {JSON_PATH} 파일을 찾을 수 없습니다.')
        sys.exit(1)

    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        app_data = json.load(f)

    # 1. Check Downloads folder for recent web backup
    dl_backup = os.path.expandvars(r'%USERPROFILE%\Downloads\autocall_ledger_backup.json')
    if os.path.exists(dl_backup):
        try:
            dl_mtime = os.path.getmtime(dl_backup)
            json_mtime = os.path.getmtime(JSON_PATH)
            if dl_mtime > json_mtime:
                with open(dl_backup, 'r', encoding='utf-8') as f:
                    imported = json.load(f)
                if len(imported.get('transactions', [])) >= len(app_data.get('transactions', [])):
                    print(f'  [동기화] 다운로드 폴더의 백업 파일({len(imported["transactions"])}건)을 자동 반영했습니다.')
                    app_data = imported
        except Exception as e:
            pass

    # 2. Check Excel file for any new transactions
    excel_txs = parse_excel_transactions(XLSX_PATH)
    if excel_txs and len(excel_txs) > len(app_data.get('transactions', [])):
        existing_ids = {t['id'] for t in app_data['transactions']}
        new_items = [t for t in excel_txs if t['id'] not in existing_ids]
        if new_items:
            print(f'  [동기화] 엑셀 파일에서 새 거래 내역 {len(new_items)}건을 감지하여 추가했습니다.')
            for nt in new_items:
                app_data['transactions'].append(nt)
                person = nt['person']
                if person in app_data['members']:
                    m = nt['month']
                    fee = app_data.get('monthlyFees', {}).get(m, 20000)
                    if m in app_data.get('months', []):
                        app_data['monthlyMatrix'][person][m] = (app_data['monthlyMatrix'][person].get(m, 0) or 0) + fee
                    if nt.get('extraMonth'):
                        for em in nt['extraMonth'].replace(',', ' ').split():
                            em = em.strip()
                            if em in app_data.get('months', []):
                                fee_em = app_data.get('monthlyFees', {}).get(em, 20000)
                                app_data['monthlyMatrix'][person][em] = (app_data['monthlyMatrix'][person].get(em, 0) or 0) + fee_em

    now_str = datetime.datetime.now().strftime('%Y.%m.%d %H:%M')
    app_data['updatedAt'] = now_str

    # Save updated ledger_data.json
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(app_data, f, ensure_ascii=False, indent=2)

    # 3. Inject into index.html
    if not os.path.exists(HTML_PATH):
        print(f'❌ {HTML_PATH} 파일을 찾을 수 없습니다.')
        sys.exit(1)

    with open(HTML_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    # Cache control meta tags
    if 'Cache-Control' not in html:
        meta_cache = '''  <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">\n  <meta http-equiv="Pragma" content="no-cache">\n  <meta http-equiv="Expires" content="0">'''
        html = html.replace('<meta name="viewport" content="width=device-width, initial-scale=1.0">',
                            '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n' + meta_cache)

    # Header update badge
    if 'lastUpdatedText' not in html:
        old_header_p = '<p class="text-xs text-slate-400 hidden sm:block">2026 오토콜 정기 회비 스마트 관리 시스템</p>'
        new_header_div = '''<div class="flex items-center space-x-2 text-xs text-slate-400 mt-0.5">
              <span class="hidden sm:inline">2026 오토콜 정기 회비 스마트 관리 시스템</span>
              <span class="hidden sm:inline text-slate-600">•</span>
              <span class="text-emerald-400 font-medium inline-flex items-center">
                <i class="fa-solid fa-clock-rotate-left mr-1"></i>최종 업데이트: <span id="lastUpdatedText">''' + now_str + '''</span>
              </span>
            </div>'''
        html = html.replace(old_header_p, new_header_div)
    else:
        html = re.sub(r'<span id="lastUpdatedText">.*?</span>', f'<span id="lastUpdatedText">{now_str}</span>', html, count=1)

    # Fix loadData for viewers
    old_loaddata = '''    // Load Data from LocalStorage or Default
    function loadData() {
      const saved = localStorage.getItem('autocall_ledger_data');
      if (saved) {
        try {
          appData = JSON.parse(saved);
        } catch (e) {
          console.error('Failed to parse local storage data', e);
          appData = JSON.parse(JSON.stringify(INITIAL_DATA));
        }
      } else {
        appData = JSON.parse(JSON.stringify(INITIAL_DATA));
      }'''

    new_loaddata = '''    // Load Data from Server (INITIAL_DATA) or Admin Local Draft
    function loadData() {
      const adminSession = sessionStorage.getItem('autocall_admin_logged_in');
      const isAdminSession = (adminSession === 'true');
      const saved = localStorage.getItem('autocall_ledger_data');

      if (isAdminSession && saved) {
        try {
          appData = JSON.parse(saved);
        } catch (e) {
          console.error('Failed to parse local storage data', e);
          appData = JSON.parse(JSON.stringify(INITIAL_DATA));
        }
      } else {
        appData = JSON.parse(JSON.stringify(INITIAL_DATA));
        if (!isAdminSession && saved) {
          localStorage.removeItem('autocall_ledger_data');
        }
      }'''

    if old_loaddata in html:
        html = html.replace(old_loaddata, new_loaddata)

    # Replace INITIAL_DATA
    json_str = json.dumps(app_data, ensure_ascii=False)
    html = re.sub(r'const INITIAL_DATA = [\s\S]*?;\s*\n\s*// State', f'const INITIAL_DATA = {json_str};\n\n    // State', html)

    with open(HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    last_tx = app_data['transactions'][-1] if app_data['transactions'] else {}
    print('  [빌드 성공] index.html 최신 데이터 주입 완료!')
    print(f'   * 총 거래 건수: {len(app_data["transactions"])}건')
    print(f'   * 최종 거래: [ID {last_tx.get("id")}] {last_tx.get("person")} - {last_tx.get("amount", 0):,}원 ({last_tx.get("date")})')
    print(f'   * 동기화 일시: {now_str}')

if __name__ == '__main__':
    main()
