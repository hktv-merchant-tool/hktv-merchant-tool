# HKTVmall B2B 智慧招商工具

基於「2026 HKTVmall New Merchant Acquisition Deck [CHI].pptx」的 B2B 招商提案動態生成工具。

## 功能

- 輸入商戶名稱、產品品類、自選痛點
- 選擇語氣風格（6種）和輸出格式（5種）
- 填寫自訂 Prompt 指令
- 即時生成客製化招商提案，包含 HKTVmall 全套核心數據
- 資料來自 PPT：310萬+用戶、78.9億 GMV、7大倉庫、HK$25,000年費等

## 部署

本工具支援以下部署方式：

### Streamlit Community Cloud（免費，推薦）

1. 將本專案上傳至 GitHub 倉庫
2. 前往 https://streamlit.io/cloud
3. 連接 GitHub 帳號，選擇倉庫與分支
4. 設定 `app.py` 為主程式
5. 點擊 Deploy — 完成！

### 本地運行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 技術栈

- Streamlit
- python-pptx
- 內嵌 PPT 知識庫（無需 API key）
