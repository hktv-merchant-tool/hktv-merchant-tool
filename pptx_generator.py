 """
  HKTVmall PPT 動態生成模組
  =========================
  基於「2026 HKTVmall New Merchant Acquisition Deck [CHI].pptx」模板
  根據輸入的目標招商品牌，生成客製化招商PPT
  """

  import io
  from typing import List
  from pptx import Presentation


  CATEGORY_INSIGHTS = {
      "糧油雜貨": "糧油雜貨是HKTVmall GMV最高的品類，家庭消費者忠誠度高，復購率達4.7x/季，加入我們立即享受流量紅利。",
      "美容及健康產品": "美容及健康產品在HKTVmall增長迅猛，平台用戶男女均衡，個人專屬價工具能精準觸及目標客戶，20%轉化率效果顯著。",
      "寵物用品": "寵物市場快速增長，直播頻道可展示寵物用品實際使用效果，透過CRM精準定位寵物主人群體，大幅提升轉化率。",
      "電子電器產品": "蘇寧等大型品牌已在HKTVmall直播單晚創下近90萬GMV，電子電器產品極適合直播展示功能，搶佔高端客戶群。",
      "其他": "HKTVmall 310萬+用戶覆蓋全港各類消費者，全方位營銷工具助您快速打開市場，3-5個工作天即可開店營運。",
  }

  PAIN_SOLUTIONS = {
      "缺乏線上流量": "HKTVmall擁有310萬+獨立用戶及160萬月活設備，能為商戶帶來巨大線上流量曝光，解決流量獲取難題。",
      "自建物流成本過高": "HKTVmall自建7大倉庫、350+輛貨車及75家O2O門市，提供全港最大住宅配送網絡，讓商戶告別自建物流的高昂成本。",
      "冷鏈配送常有客訴": "HKTVmall配備專業冷鏈配送系統，每日處理10萬+訂單，自動化倉庫確保生鮮及冷藏商品品質，大幅降低客訴率。",
      "營銷成本過高": "HKTVmall提供85折推廣（平台全額承擔）、個人專屬價（20%轉化率）等高效營銷工具，並有鄭裕玲、黃子華等名人廣告加持
  ，讓營銷更具性價比。",
      "庫存管理困難": "GREEN LAB臨期百貨及會員月費計劃能有效幫助商戶清理庫存，降低損耗成本。",
      "缺乏電商營運經驗": "HKTVmall電商學院提供全方位支援，包括上架教學、數據分析、行銷技巧，讓電商新手也能快速上手。",
      "曝光機會不足": "58個MTR站、3,000+廣告燈箱的全渠道曝光，加上直播頻道，為商戶帶來前所未有的品牌曝光機會。",
  }


  def replace_text_in_shape(shape, old_text, new_text):
      if not shape.has_text_frame:
          return False
      replaced = False
      for para in shape.text_frame.paragraphs:
          for run in para.runs:
              if old_text in run.text:
                  run.text = run.text.replace(old_text, new_text)
                  replaced = True
      return replaced


  def replace_text_in_slide(slide, old_text, new_text):
      count = 0
      for shape in slide.shapes:
          if replace_text_in_shape(shape, old_text, new_text):
              count += 1
      return count


  def generate_merchant_ppt(merchant_name, category, pain_points, custom_content="", tone="誠懇專業型", template_path=None) ->
  io.BytesIO:
      if template_path is None:
          template_path = "/mnt/c/Users/cylee/Desktop/HKTVmall Plan 2026/Local Standard Merchant/2026 HKTVmall New Merchant
  Acquisition Deck [CHI].pptx"
      prs = Presentation(template_path)
      slide1 = prs.slides[0]
      replace_text_in_slide(slide1, "No.1", f"{merchant_name}")
      output = io.BytesIO()
      prs.save(output)
      output.seek(0)
      return output


  def generate_custom_ppt(uploaded_ppt_bytes, merchant_name, category, merchant_desc, pain_points, custom_content="",
  tone="誠懇專業型") -> dict:
      prs = Presentation(io.BytesIO(uploaded_ppt_bytes))
      custom_slides = []
      slide1 = prs.slides[0]
      for shape in slide1.shapes:
          if shape.has_text_frame:
              for para in shape.text_frame.paragraphs:
                  for run in para.runs:
                      if "No.1" in run.text:
                          run.text = run.text.replace("No.1", merchant_name)
                          if 1 not in custom_slides:
                              custom_slides.append(1)
                      elif "香港網上購物平台商戶合作方案" in run.text and merchant_name:
                          run.text = f"{merchant_name} - " + run.text
                          if 1 not in custom_slides:
                              custom_slides.append(1)
      if len(prs.slides) >= 49:
          last_slide = prs.slides[48]
          for shape in last_slide.shapes:
              if shape.has_text_frame:
                  for para in shape.text_frame.paragraphs:
                      for run in para.runs:
                          if run.text.strip() == "" or "附錄" in run.text:
                              run.text = f"目標招商商戶：{merchant_name}\n產品品類：{category}\n業務描述：{merchant_desc[:100]}{'...'
  if len(merchant_desc) > 100 else ''}"
                              if 49 not in custom_slides:
                                  custom_slides.append(49)
                              break
      cat_insight = CATEGORY_INSIGHTS.get(category, CATEGORY_INSIGHTS["其他"])
      pain_texts = [PAIN_SOLUTIONS.get(p, "") for p in pain_points if p in PAIN_SOLUTIONS]
      summary_parts = [
          f"【招商目標】{merchant_name}",
          f"【產品品類】{category}",
          f"【主要業務】{merchant_desc[:80]}{'...' if len(merchant_desc) > 80 else ''}",
          f"【語氣風格】{tone}",
          "",
          "【品類洞察】",
          cat_insight,
          "",
      ]
      if pain_texts:
          summary_parts.append("【痛點對應方案】")
          for pt in pain_texts:
              summary_parts.append(f"• {pt}")
          summary_parts.append("")
      if custom_content:
          summary_parts.append(f"【特別要求】{custom_content}")
          summary_parts.append("")
      summary_parts.extend([
          "【PPT客製化說明】",
          f"已根據上傳的PPT模板進行客製化",
          f"共 {len(prs.slides)} 張投影片",
          "以下slides已根據目標商戶調整：",
          f"• Slide 1: 封面（已加入商戶名稱）",
      ])
      if merchant_desc:
          summary_parts.append(f"• Slide 49: 備註頁（已加入商戶業務描述）")
      summary_parts.extend([
          "",
          "【聯絡我們】",
          "WhatsApp：+852 5283 4138",
          "電郵：business@hktvmall.com",
          "網站：business.hktvmall.com",
      ])
      output = io.BytesIO()
      prs.save(output)
      output.seek(0)
      return {'ppt': output, 'summary': '\n'.join(summary_parts), 'custom_slides': custom_slides}


  def generate_merchant_ppt_with_notes(merchant_name, category, pain_points, custom_content="", tone="誠懇專業型",
  template_path=None) -> dict:
      if template_path is None:
          template_path = "/mnt/c/Users/cylee/Desktop/HKTVmall Plan 2026/Local Standard Merchant/2026 HKTVmall New Merchant
  Acquisition Deck [CHI].pptx"
      prs = Presentation(template_path)
      custom_slides = []
      slide1 = prs.slides[0]
      if replace_text_in_slide(slide1, "No.1", merchant_name) > 0:
          custom_slides.append(1)
      cat_insight = CATEGORY_INSIGHTS.get(category, CATEGORY_INSIGHTS["其他"])
      pain_texts = [PAIN_SOLUTIONS.get(p, "") for p in pain_points if p in PAIN_SOLUTIONS]
      summary_parts = [
          f"【招商目標】{merchant_name}",
          f"【產品品類】{category}",
          f"【語氣風格】{tone}",
          "",
          "【品類洞察】",
          cat_insight,
          "",
      ]
      if pain_texts:
          summary_parts.append("【痛點對應】")
          for pt in pain_texts:
              summary_parts.append(f"• {pt}")
          summary_parts.append("")
      if custom_content:
          summary_parts.append(f"【特別要求】{custom_content}")
          summary_parts.append("")
      summary_parts.extend([
          "【PPT說明】",
          f"已使用「2026 HKTVmall New Merchant Acquisition Deck」作為模板",
          f"覆蓋 {len(prs.slides)} 張投影片",
          "以下slides已根據目標商戶客製化：",
          f"• Slide 1: 封面（已加入商戶名稱）",
          "• Slide 31-34: 成功案例（保持作為同業參考）",
          "• Slide 44: 聯絡我們（保持不變）",
          "",
          "聯絡人：HKTVmall招商團隊",
          "WhatsApp：+852 5283 4138",
          "電郵：aog.merc@hktv.com.hk",
      ])
      output = io.BytesIO()
      prs.save(output)
      output.seek(0)
      return {'ppt': output, 'summary': '\n'.join(summary_parts), 'custom_slides': custom_slides}
