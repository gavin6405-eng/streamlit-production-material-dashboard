# 超慧科技｜Streamlit 生管資材戰情看板

這是由 FIX10 Flask 程式轉換的 Streamlit 版，不需 Windows 本機啟動器。
入口：`streamlit_app.py`；Python：3.12；GitHub 分支：`main`。

## 最快上線方式

1. 解壓縮本套件。若不用指令，至 https://github.com/new 建立 **Private** 儲存庫，選 Add file → Upload files，上傳本資料夾內檔案（不要只上傳 ZIP）。確認 streamlit_app.py 位於儲存庫最上層。
2. 想自動建立 GitHub：電腦先安装 Git 及 GitHub CLI (https://cli.github.com/)，在此資料夾開 PowerShell 執行 `powershell -ExecutionPolicy Bypass -File .\PUBLISH_GITHUB.ps1`。它只建立私人 GitHub 儲存庫與上傳程式，不啟動本機看板。首次會要求 GitHub 瀏覽器登入；不覆寫既有 repository。
3. 準備可由網路連線的 PostgreSQL 資料庫，取得 SSL 連線字串。請由公司資訊部門提供，或使用自行管理的雲端 PostgreSQL。資料庫帳號需能建立 dashboard_records 表及讀寫資料。持續保留此資料庫與備份。
4. 前往 https://share.streamlit.io/ → Create app，選擇上述 repository、main、streamlit_app.py。Advanced settings 選 Python 3.12，Secrets 貼上 secrets.example.toml 內容，換成真實 DATABASE_URL 及兩組不同、至少12字的密碼，再按 Deploy。
5. 開啟產生的網址，以 admin 管理／viewer 瀏覽。不要將真實 secrets.toml、舊資料庫或公司 Excel 上傳 GitHub。

帳號登入、外部資料庫開通及 Secrets 設定必須由帳號持有人完成；套件不包含帳號或資料庫憑證，不能跳過這些設定直接上線。
官方部署說明：https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/deploy

## 功能

- 四大模組：2026排程、缺料追蹤、IQC待驗、倉庫概況；戰情總覽、組立進度台數、風險燈號、搜尋、CSV 匯出。
- 右側上傳既有「2026排程」及「在線缺料」Excel，沿用 FIX10 解析與欄位順序。
- 匯入前預覽；相同編號更新，未包含的舊資料不刪除。Excel 中相同製令／料號／Type／Category 重複時仍沿用原版出現順序識別，請勿任意調換重複列順序。
- 進度／缺料差異與未變動天數沿用原版規則。IQC 新增時固定記錄收料時間。
- 管理者新增／編輯，瀏覽者只讀；並行更新採版本檢查，衝突時拒絕整批匯入。
- 資料持續存於外部 PostgreSQL，更新程式不清空、不初始化示範資料。連線失敗會明確停止，不以空白資料替代。

## 舊資料移轉（首次執行一次）

此來源 ZIP 沒有 dashboard.db，所以本包未包含你的既有進度。舊程式的資料不會自動出現在雲端。

1. 先關閉舊程式，備份 `%LOCALAPPDATA%\SuperPlusTech\MaterialsDashboard\data` 整個資料夾。若先前自訂 DATA_DIR，使用該路徑。
2. 管理者登入新看板 →「資料備份／舊版匯入」，上傳 dashboard.db，點「合併匯入」。
3. 逐一核對四個模組筆數／進度。相同編號保留雲端值，不覆寫；舊 payload 額外欄位與比對時間保留。
4. 下載全部 JSON 備份。若中途連線中斷，可重試；已存在編號會略過。各模組分別交易，不同模組可能部分成功。

雲端登入密碼以 Secrets 新設定為準，不會搬移舊版 config.json 密碼。

## 更新與備份

後續只更新 GitHub 程式檔，保留同一 DATABASE_URL。不要刪除資料庫或換成新空白資料庫。雲端資料庫的停用、帳號到期或刪除仍會影響資料可用性，請配置資料庫服務的備份並定期下載 JSON。

本套件已提供 GitHub Actions 語法及單元檢查；實際雲端連線需填入你的設定後驗證。沒有宣稱已建立 GitHub 儲存庫或已部署網址。
