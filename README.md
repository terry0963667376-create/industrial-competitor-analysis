# 工業產品規格書競品比較 MVP

免費、規則式的 Streamlit 雙 PDF 規格擷取、人工編輯、競品比較及 Excel 匯出示範網站。

## 功能
1. 上傳產品 A 與 B 的 PDF（需有文字層）。
2. 檢視、修正標準化規格。
3. 產生比較表並下載多工作表 Excel 報告。

## 本機執行
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## 免費部署到 Streamlit Community Cloud
1. 在 GitHub 建立 Repository，將此資料夾內的檔案上傳到 Repository 根目錄（包括 `.python-version`）。
2. 開啟 https://share.streamlit.io/，使用 GitHub 帳號登入並建立 App。
3. 選擇 Repository、`main` 分支及主程式 `app.py`，按 Deploy。
4. 部署後取得 `*.streamlit.app` 網址，使用公開的測試 PDF 驗證上傳、編輯、匯出。

## 注意
- 目前未使用 AI，也未連接付費 API。
- 只支援有文字層的 PDF；影像掃描 PDF 不一定能擷取。
- 規格配對與分析屬候選結果，須人工確認；未擷取到不代表產品沒有該規格。
- 公開展示時請只上傳公開或有權使用的資料；不要將公司機密或 API 金鑰提交至公開儲存庫。
- `.python-version` 指定雲端 Python 3.12，不影響你本機 Python 3.15。
