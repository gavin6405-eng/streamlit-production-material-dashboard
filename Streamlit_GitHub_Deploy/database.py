import json
import psycopg
from logic import MODULES, validate, compare_schedule_row, compare_material_row

def connect(url):
    return psycopg.connect(url, connect_timeout=15)

def initialize(url):
    with connect(url) as c:
        c.execute('CREATE TABLE IF NOT EXISTS dashboard_records(module TEXT NOT NULL,id TEXT NOT NULL,payload TEXT NOT NULL,version INTEGER NOT NULL DEFAULT 1,PRIMARY KEY(module,id))')

def load(url):
    with connect(url) as c:
        return {m:{rid:(json.loads(p),v) for mod,rid,p,v in c.execute('SELECT module,id,payload,version FROM dashboard_records WHERE module=%s',(m,))} for m in MODULES}

def save(url,module,items,expected,compare=False,missing_only=False):
    inserted=0
    if module not in MODULES: raise ValueError('未知模組')
    with connect(url) as c:
        for row in items:
            row={**row,**validate(row,module)}
            rid=row['編號']; old=expected.get(rid)
            if compare and module in ['production','material']:
                row=(compare_schedule_row if module=='production' else compare_material_row)(row,old[0] if old else None)
            payload=json.dumps(row,ensure_ascii=False)
            if missing_only:
                cur=c.execute('INSERT INTO dashboard_records VALUES(%s,%s,%s,1) ON CONFLICT DO NOTHING',(module,rid,payload))
            elif old:
                cur=c.execute('UPDATE dashboard_records SET payload=%s,version=version+1 WHERE module=%s AND id=%s AND version=%s',(payload,module,rid,old[1]))
                if cur.rowcount!=1: raise ValueError('資料已被他人修改，請重新整理後再操作；本次未儲存。')
            else:
                cur=c.execute('INSERT INTO dashboard_records VALUES(%s,%s,%s,1) ON CONFLICT DO NOTHING',(module,rid,payload))
                if cur.rowcount!=1: raise ValueError('編號已存在，請重新整理；本次未儲存。')
            inserted+=cur.rowcount
    return inserted
