import io,json,hmac,sqlite3,tempfile
from pathlib import Path
from collections import Counter
import pandas as pd
import streamlit as st
from logic import *
import database as db
from schedule_import import read_schedule
from material_import import read_material

st.set_page_config(page_title='超慧科技｜生管資材戰情看板',page_icon='📊',layout='wide')
st.title('超慧科技｜生管資材戰情看板')
try:
    url=st.secrets['DATABASE_URL']
    passwords={r:str(st.secrets[r+'_password']) for r in ['admin','viewer']}
    if any(len(p)<12 or p.startswith('CHANGE_') for p in passwords.values()) or passwords['admin']==passwords['viewer']:
        raise ValueError('請設定兩組不同且至少12字的密碼')
except Exception:
    st.info('首次部署：請依 README 在 Streamlit Secrets 設定 DATABASE_URL、admin_password、viewer_password。設定完成後重新啟動。')
    st.stop()
if not st.session_state.get('role'):
    with st.form('login'):
        role=st.selectbox('登入身分',['viewer','admin'],format_func=lambda r:'管理者' if r=='admin' else '瀏覽者')
        password=st.text_input('密碼',type='password')
        if st.form_submit_button('登入'):
            if hmac.compare_digest(password.encode(),passwords[role].encode()):
                st.session_state.role=role;st.rerun()
            else:st.error('密碼不正確')
    st.stop()
admin=st.session_state.role=='admin'
with st.sidebar:
    st.image('static/logo.png',width=160)
    page=st.radio('功能',['戰情總覽',*MODULES.values(),'資料備份／舊版匯入'])
    if st.button('重新整理'):st.rerun()
    if st.button('登出'):
        st.session_state.clear();st.rerun()
try:
    db.initialize(url);all_data=db.load(url)
except Exception:
    st.error('無法連線至雲端資料庫。請確認 DATABASE_URL、SSL 與資料庫服務狀態；系統未切換為空白本機資料。');st.stop()

def table(rows,module):
    out=[]
    for r in rows:
        r=dict(r);decorate_compare(r,module)
        color,label=risk(r,module)
        out.append({'風險':{'red':'🔴','yellow':'🟡','green':'🟢'}[color],'風險說明':label,**{k:v for k,v in r.items() if not k.startswith('_')}})
    return pd.DataFrame(out).fillna('')

if page=='戰情總覽':
    for col,(m,title) in zip(st.columns(4),MODULES.items()):
        rows=[r for r,v in all_data[m].values()]
        col.metric(title,len(rows));col.caption(f'未結案 {sum(r["狀態"] not in DONE[m] for r in rows)} 筆')
    for m,title in MODULES.items():
        st.subheader(title)
        rows=[r for r,v in all_data[m].values()]
        if m=='production' and rows:
            st.bar_chart(pd.Series(Counter(str(r.get('組立進度') or r['狀態']) for r in rows),name='台數'))
        st.dataframe(table(rows,m),use_container_width=True,hide_index=True)
elif page=='資料備份／舊版匯入':
    if not admin:st.info('請使用管理者登入。');st.stop()
    backup={m:[r for r,v in entries.values()] for m,entries in all_data.items()}
    st.download_button('下載全部資料 JSON 備份',json.dumps(backup,ensure_ascii=False,indent=2),'dashboard_backup.json','application/json')
    st.caption('舊版 dashboard.db 位於 Windows：%LOCALAPPDATA%\\SuperPlusTech\\MaterialsDashboard\\data。匯入前先關閉舊程式並備份；同編號保留雲端現有資料。')
    f=st.file_uploader('匯入舊版資料庫或本版 JSON 備份',type=['db','json'])
    if f and st.button('合併匯入（不覆寫現有資料）'):
        try:
            if f.name.endswith('.json'):data=json.loads(f.getvalue())
            else:
                with tempfile.TemporaryDirectory() as tmp:
                    path=Path(tmp)/'legacy.db';path.write_bytes(f.getvalue())
                    with sqlite3.connect(f'file:{path}?mode=ro',uri=True) as c:
                        data={m:[] for m in MODULES}
                        for m,p in c.execute('SELECT module,payload FROM records'):
                            if m not in MODULES:raise ValueError('資料庫包含未知模組')
                            data[m].append(json.loads(p))
            for m,rows in data.items():
                if m not in MODULES:raise ValueError('備份包含未知模組')
                for r in rows:validate(r,m)
            for m,rows in data.items():
                count=db.save(url,m,rows,{},missing_only=True)
                st.success(f'{MODULES[m]}：新增 {count} 筆，其餘同編號保留')
        except Exception as e:st.error('匯入失敗，已完成的模組可安全重試。'+(str(e) if isinstance(e,ValueError) else '請檢查檔案及資料庫連線。'))
else:
    module=next(m for m,t in MODULES.items() if t==page)
    entries=all_data[module];rows=[r for r,v in entries.values()]
    main,right=st.columns([3,1])
    with main:
        keyword=st.text_input('搜尋製令／料號／專案／備註')
        selected=[r for r in rows if keyword.lower() in json.dumps(r,ensure_ascii=False).lower()]
        frame=table(selected,module)
        st.dataframe(frame,use_container_width=True,hide_index=True)
        st.download_button('下載目前清單 CSV',frame.to_csv(index=False).encode('utf-8-sig'),f'{module}.csv','text/csv')
    with right:
        st.subheader('Excel 上傳')
        if admin and module in ['production','material']:
            f=st.file_uploader('原格式 Excel',type=['xlsx'],key=module)
            if f and st.button('預覽匯入'):
                try:
                    imported,warnings=(read_schedule if module=='production' else read_material)(io.BytesIO(f.getvalue()))
                    st.session_state['preview_'+module]=(imported,warnings,entries)
                except Exception as e:st.error(str(e))
            preview=st.session_state.get('preview_'+module)
            if preview:
                imported,warnings,baseline=preview
                st.write(f'待匯入 {len(imported)} 筆（同編號更新，其他資料保留）')
                st.dataframe(table(imported,module),hide_index=True)
                with st.expander('匯入提示'):st.write(warnings or ['無'])
                if st.button('確認寫入'):
                    try:
                        db.save(url,module,imported,baseline,compare=True)
                        del st.session_state['preview_'+module];st.rerun()
                    except ValueError as e:st.error(str(e))
                    except Exception:st.error('寫入失敗；此次交易已回復，請檢查資料庫連線。')
        else:st.caption('管理者可匯入排程／缺料原檔；IQC 與倉庫使用下方表單維護。')
    if admin:
        st.subheader('新增／編輯資料')
        rid=st.selectbox('選擇資料',['新增',*entries.keys()])
        editkey=(module,rid)
        if st.session_state.get('editkey')!=editkey:
            st.session_state.editkey=editkey
            st.session_state.editbase=entries.get(rid,({},0))
        old,version=st.session_state.editbase
        with st.form('edit_'+module+'_'+rid):
            values={}
            cols=st.columns(3)
            for i,field in enumerate(FIELDS):
                with cols[i%3]:
                    if field in ['狀態','優先級']:
                        opts=STATES[module] if field=='狀態' else ['一般','急件']
                        values[field]=st.selectbox(field,opts,index=opts.index(old[field]) if old.get(field) in opts else 0)
                    else:values[field]=st.text_input(field,str(old.get(field,'')),disabled=field=='編號' and rid!='新增')
            if old.get('_source') in ['2026排程','在線缺料']:
                values['組立進度']=st.text_input('組立進度',str(old.get('組立進度','')))
            if st.form_submit_button('儲存'):
                try:
                    row={**old,**values}
                    if row.get('_source'):
                        row.update({'客戶':row['專案'],'P/N':row['料號或機型'],'Type':row['名稱'],'組立地點':row['供應商或地點'],'組立人員':row['負責人'],'入庫日':row['預計日期'] if module=='production' else row['需求日期']})
                        if module=='material':row['齊料']='齊料' if row['狀態'] in DONE[module] else ''
                    if module=='iqc':row.setdefault('收料時間',now_stamp())
                    db.save(url,module,[row],{rid:(old,version)} if old else {},compare=bool(old.get('_source')))
                    st.session_state.pop('editkey',None);st.rerun()
                except ValueError as e:st.error(str(e))
                except Exception:st.error('儲存失敗，請檢查資料庫連線後重試。')
