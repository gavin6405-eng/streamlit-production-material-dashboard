import math
from datetime import datetime,date,timedelta
from zoneinfo import ZoneInfo
MODULES = {'production': '2026排程', 'material': '缺料追蹤', 'iqc': 'IQC 待驗', 'warehouse': '倉庫概況'}

FIELDS = ['編號', '專案', '料號或機型', '名稱', '供應商或地點', '數量', '完成數量', '需求日期', '預計日期', '狀態', '優先級', '負責人', '備註']

STATES = {'production': ['未開始', '組立中', '檢驗中', '已完成', '暫停'], 'material': ['待確認', '齊料', '待回覆', '採購中', '已到貨', '已結案'], 'iqc': ['待驗', '檢驗中', '合格', '不合格', '已結案'], 'warehouse': ['使用中', '空儲位', '停用']}

DONE = {'production': {'已完成'}, 'material': {'齊料', '已到貨', '已結案'}, 'iqc': {'合格', '已結案'}, 'warehouse': {'空儲位', '停用'}}

def today():
    return datetime.now(ZoneInfo('Asia/Taipei')).date()

def validate(row, module):
    r = {f: str(row.get(f, '') if row.get(f) is not None else '').strip() for f in FIELDS}
    if not r['編號']:
        raise ValueError('編號不可空白')
    if any((len(v) > 1000 for v in r.values())):
        raise ValueError('欄位長度不可超過1000字')
    for f in ['數量', '完成數量']:
        try:
            n = float(r[f] or '0')
        except ValueError:
            raise ValueError(f + '須為數字')
        if not math.isfinite(n) or n < 0:
            raise ValueError(f + '須為非負數')
        r[f] = n
    if r['完成數量'] > r['數量']:
        raise ValueError('完成數量不可大於數量')
    for f in ['需求日期', '預計日期']:
        if r[f]:
            try:
                r[f] = date.fromisoformat(r[f][:10]).isoformat()
            except ValueError:
                raise ValueError(f + '須為 YYYY-MM-DD')
    if r['狀態'] not in STATES[module]:
        raise ValueError('狀態不符此模組選項')
    if r['優先級'] not in ['一般', '急件']:
        raise ValueError('優先級須為一般或急件')
    return r

TAIPEI = ZoneInfo('Asia/Taipei')

MATERIAL_COMPARE_FIELDS = ['Frame/ Frame set', 'PU', 'Facility', '其他託外模組', 'RB', 'LP', 'AL', 'FFU', 'X-table', '加工件', '市購件', '齊料']

def now_stamp():
    return datetime.now(TAIPEI).isoformat(timespec='seconds')

def parse_day(v):
    s = str(v or '').strip()
    if not s:
        return None
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        return None

def elapsed_days(stamp):
    s = str(stamp or '').strip()
    if not s:
        return 0
    try:
        d = datetime.fromisoformat(s).date()
    except ValueError:
        try:
            d = date.fromisoformat(s[:10])
        except ValueError:
            return 0
    return max((today() - d).days, 0)

def display_value(v):
    s = str(v or '').strip()
    return s if s else '空白'

def compare_schedule_row(new, old=None, stamp=None):
    """保留上一次生產進度基準，讓每日上傳可判斷是否有推進。"""
    r = dict(new)
    stamp = stamp or now_stamp()
    current = str(r.get('組立進度', '') or '').strip()
    previous = str((old or {}).get('組立進度', '') or '').strip()
    r['_last_imported_at'] = stamp
    r['_previous_progress'] = previous
    if not old:
        r['_progress_changed_at'] = stamp
        r['_last_compare_changed'] = None
        r['_progress_diff'] = '首次匯入／建立基準'
    elif current != previous:
        r['_progress_changed_at'] = stamp
        r['_last_compare_changed'] = True
        r['_progress_diff'] = f'{display_value(previous)} → {display_value(current)}'
    else:
        r['_progress_changed_at'] = old.get('_progress_changed_at') or old.get('_last_imported_at') or stamp
        r['_last_compare_changed'] = False
        r['_progress_diff'] = '無變更'
    return r

def compare_material_row(new, old=None, stamp=None):
    """比較缺料相關欄位的新舊差異，保留最近一次有變化的時間。"""
    r = dict(new)
    stamp = stamp or now_stamp()
    changes = []
    r['_last_imported_at'] = stamp
    if old:
        for f in MATERIAL_COMPARE_FIELDS:
            ov = str(old.get(f, '') or '').strip()
            nv = str(r.get(f, '') or '').strip()
            if ov != nv:
                changes.append(f'{f}：{display_value(ov)} → {display_value(nv)}')
    if not old:
        r['_material_changed_at'] = stamp
        r['_last_compare_changed'] = None
        r['_material_diff'] = '首次匯入／建立基準'
    elif changes:
        r['_material_changed_at'] = stamp
        r['_last_compare_changed'] = True
        r['_material_diff'] = '；'.join(changes)[:1800]
    else:
        r['_material_changed_at'] = old.get('_material_changed_at') or old.get('_last_imported_at') or stamp
        r['_last_compare_changed'] = False
        r['_material_diff'] = '無變更'
    return r

def decorate_compare(r, module):
    """建立畫面用欄位；未變動天數每天自動增加，不必修改原資料。"""
    if module == 'production':
        days = elapsed_days(r.get('_progress_changed_at'))
        last_changed = r.get('_progress_changed_at', '')
        if r.get('狀態') in DONE['production']:
            result = '已完成／停止異常追蹤'
        elif days >= 1:
            result = f'異常：進度未變動 {days} 天'
        elif r.get('_last_compare_changed') is True:
            result = '正常：進度有變更'
        elif r.get('_last_compare_changed') is False:
            result = '進度未變動（未滿 1 天）'
        elif r.get('_source') == '2026排程':
            result = '首次匯入／建立比對基準'
        else:
            result = '—'
        r['比對結果'] = result
        r['進度差異'] = r.get('_progress_diff', '—')
        r['未變動天數'] = f'{days} 天' if r.get('_source') == '2026排程' else '—'
        r['上次進度變更'] = str(last_changed).replace('T', ' ') if last_changed else '—'
        r['最後比對'] = str(r.get('_last_imported_at', '')).replace('T', ' ') or '—'
    elif module == 'material':
        days = elapsed_days(r.get('_material_changed_at'))
        last_changed = r.get('_material_changed_at', '')
        if r.get('狀態') in DONE['material']:
            result = '已完成／停止異常追蹤'
        elif days >= 1:
            result = f'異常：缺料未變動 {days} 天'
        elif r.get('_last_compare_changed') is True:
            result = '缺料有變更'
        elif r.get('_last_compare_changed') is False:
            result = '缺料未變動（未滿 1 天）'
        elif r.get('_source') == '在線缺料':
            result = '首次匯入／建立比對基準'
        else:
            result = '—'
        r['比對結果'] = result
        r['缺料差異'] = r.get('_material_diff', '—')
        r['未變動天數'] = f'{days} 天' if r.get('_source') == '在線缺料' else '—'
        r['上次缺料變更'] = str(last_changed).replace('T', ' ') if last_changed else '—'
        r['最後比對'] = str(r.get('_last_imported_at', '')).replace('T', ' ') or '—'
    return r

def risk(r, module):
    if r['狀態'] in DONE[module]:
        return ('green', '已完成／不占用')
    if module == 'warehouse':
        return ('green', '使用中')
    if module == 'production':
        due = parse_day(r.get('入庫日') or r.get('預計日期'))
        stagnant = elapsed_days(r.get('_progress_changed_at')) if r.get('_source') == '2026排程' else 0
        if due and due < today():
            overdue = (today() - due).days
            suffix = f'／進度未變 {stagnant} 天' if stagnant >= 1 else ''
            return ('red', f'入庫逾期 {overdue} 天{suffix}')
        if stagnant >= 1:
            return ('red', f'進度異常／未變動 {stagnant} 天')
        if not due:
            return ('yellow', '入庫日未填')
        if r['優先級'] == '急件' or due <= today() + timedelta(days=3):
            return ('yellow', '入庫日三日內／關注')
        return ('green', '正常／進度有更新')
    if module == 'material' and r.get('_source') == '在線缺料':
        if r.get('齊料') == '齊料':
            return ('green', '齊料')
        due = parse_day(r.get('入庫日') or r.get('需求日期'))
        stagnant = elapsed_days(r.get('_material_changed_at'))
        if due and due < today():
            overdue = (today() - due).days
            suffix = f'／缺料未變 {stagnant} 天' if stagnant >= 1 else ''
            return ('red', f'缺料逾期 {overdue} 天{suffix}')
        if stagnant >= 1:
            return ('red', f'缺料異常／未變動 {stagnant} 天')
        if not due:
            return ('yellow', '入庫日未填／缺料待確認')
        if due <= today() + timedelta(days=3):
            return ('yellow', '缺料／入庫日三日內')
        return ('yellow', '缺料待處理')
    due = parse_day(r.get('需求日期'))
    eta = parse_day(r.get('預計日期'))
    if r['狀態'] == '不合格' or (due and (due < today() or (eta and eta > due))):
        return ('red', '逾期／異常')
    if r['優先級'] == '急件' or (due and due <= today() + timedelta(days=3)):
        return ('yellow', '急件／三日內到期')
    if not due:
        return ('yellow', '需求日期未填')
    return ('green', '正常')
