# PPTX 簡報範本準備指南

## 支援的 Placeholder 標籤

在 PowerPoint 中，以**雙大括號** `{{TAG}}` 格式寫入文字方塊，AI 會自動替換：

| 標籤 | 說明 | 建議位置 |
|------|------|---------|
| `{{MERCHANT_NAME}}` | 商戶名稱 | 封面標題、成功案例頁、年費頁 |
| `{{PAIN_POINT_SOLUTION}}` | 第一個痛點的 AI 解決方案（一句話） | 痛點/方案頁 |
| `{{PAIN_POINT_LIST}}` | 所有痛點解決方案（多行列表） | 痛點/方案頁 |
| `{{CATEGORY}}` | 產品品類 | 品類相關頁 |
| `{{DATE_TODAY}}` | 日期（目前固定為 2026） | 封面/版本頁 |

## 範例：如何埋標籤

1. 用 PowerPoint 開啟 `2026 HKTVmall New Merchant Acquisition Deck [CHI].pptx`
2. 在**封面頁**（Slide 1），找到「商戶合作方案」文字，改為：
   `{{MERCHANT_NAME}} 商戶合作方案`
3. 在**成功案例頁**（Slide 31-34），把商戶名稱（如「蘇寧」）替換為 `{{MERCHANT_NAME}}`
4. 在**年費頁**（Slide 37），在「HK$25,000」後方加上 `{{MERCHANT_NAME}}`
5. 儲存為新檔案（**不要**覆蓋原始簡報），上傳至 Streamlit App

## 字型設定

直接在 PowerPoint 中對 `{{TAG}}` 設定字型、大小、顏色，置換後**完全繼承**原本樣式。
