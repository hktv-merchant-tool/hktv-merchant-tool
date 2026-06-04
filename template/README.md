# PPTX 簡報範本準備指南

## 支援的 Placeholder 標籤

在 PowerPoint 中，以**雙大括號** `{{TAG}}` 格式寫入文字方塊，AI 會自動替換：

| 標籤 | 說明 | 來源 |
|------|------|------|
| `{{MERCHANT_NAME}}` | 商戶名稱 | 用戶輸入 |
| `{{CATEGORY}}` | 產品品類 | 用戶選擇 |
| `{{CATEGORY_GMV}}` | 品類 GMV 數據 | HKTVmall 數據庫（如「33.9億港元（佔比 43.0%）」） |
| `{{CATEGORY_TOOLS}}` | 品類熱門工具 | HKTVmall 數據庫（如「個人專屬價、直播頻道」） |
| `{{CATEGORY_PITCH}}` | 品類完整優勢話術 | HKTVmall 數據庫（完整段落） |
| `{{PAIN_POINT_SOLUTION}}` | 第一個痛點的 AI 解決方案 | 痛點選擇 |
| `{{PAIN_POINT_LIST}}` | 所有痛點對應的解決方案 | 痛點選擇 |
| `{{DATE_TODAY}}` | 年份（目前固定為 2026） | 固定值 |

## 已在 Template 中埋好的位置

使用 `/home/clyee/hktv-merchant-tool/template/HKTVmall_Template_Base.pptx`（已預埋）：
- **S1 封面**：`{{MERCHANT_NAME}} 商戶合作方案`
- **S31 蘇寧案例**：`{{MERCHANT_NAME}}`
- **S32 新乳酪品牌**：`{{MERCHANT_NAME}}`
- **S33 入駐新品牌**：`{{MERCHANT_NAME}}`
- **S34 小品牌**：`{{MERCHANT_NAME}}`
- **S37 年費頁**：`{{MERCHANT_NAME}}`

## 如何自定義更多位置

1. 用 PowerPoint 開啟 `HKTVmall_Template_Base.pptx`
2. 找到想客製化的文字方塊，直接輸入 `{{CATEGORY_GMV}}` 等標籤
3. 對標籤設定字型、大小、顏色（樣式會被完整繼承）
4. 儲存上傳

## v2.0 品類 GMV 數據（已更新至 CATEGORY_DATA）

- 超級市場：**33.9億港元（佔比 43.0%）**
- 護理保健 / 護膚化妝：**19.8億港元（佔比 25.1%）**
- 寵物用品：**8.6億港元（佔比 10.9%）**
- 家居電器：**6.6億港元（佔比 8.4%）**
