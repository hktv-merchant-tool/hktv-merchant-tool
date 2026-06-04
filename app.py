"""
HKTVmall B2B 智慧招商動態生成工具
Built on 2026 HKTVmall New Merchant Acquisition Deck [CHI].pptx
"""

import streamlit as st
import random

# ──────────────────────────────────────────────
# 內置知識庫（來自 PPT）
# ──────────────────────────────────────────────
KNOWLEDGE_BASE = {
    "平台概覽": {
        "獨立用戶": "310萬+",
        "月活躍設備": "160萬",
        "每季購買頻率": "4.7x",
        "2025總GMV": "HK$78.9億（7.89 Billion）",
        "平台地位": "香港No.1網上購書平台",
        "上市代號": "香港聯交所代號：1137",
        "成立年份": "1992年",
    },
    "物流網絡": {
        "倉庫數量": "7大貨倉（屯門、葵涌、上水、青衣、將軍澳等五大區域）",
        "貨車數量": "350+輛",
        "O2O門市": "75家",
        "每日訂單處理量": "10萬+",
        "自動化系統": "德國自動化執貨系統",
        "配送範圍": "本地及海外（澳門、澳洲、加拿大、英國）",
    },
    "營銷工具": {
        "85折推廣": "HKTVmall全額承擔85折，單日銷售額突破4,000萬",
        "個人專屬價": "根據顧客購買偏好打造個性化優惠，轉化率達20%",
        "名人廣告": "鄭裕玲、黃子華等名人代言",
        "廣告覆蓋": "58個MTR站、3,000+廣告燈箱、23列全車身包裝列車",
        "直播頻道": "商客互聯直播，單場直播可售出477+件商品",
        "GREEN_LAB臨期百貨": "有效清理庫存",
        "會員月費計劃": "搜羅500+熱銷商品，價格低於市場最低水平",
    },
    "收費模式": {
        "常規商戶加盟年費": "HK$25,000",
        "銷售佣金": "根據產品類別計算",
        "廣告金回贈": "按比例發放，需於當年12月31日前用完",
        "數碼引流計劃": "將社交媒體流量引流至HKTVmall，享特惠佣金率",
        "結算週期": "每月結算，15個工作日內款項轉至商戶銀行帳戶",
    },
    "開店流程": {
        "所需文件": "商業登記證(BR)、銀行月結單",
        "開店步驟": "開立帳戶 → 建立電子合約 → 開展業務",
        "開店時間": "3至5個工作天",
        "截單時間": "下午1點",
        "配送時效": "HKTVmall次日配送",
    },
    "成功案例": {
        "蘇寧": "2025年直播單晚GMV接近90萬，產品數量1,700+件",
        "中小品牌": "2025年Q1已售出1,000+件",
        "小品牌": "業績年增長3倍，實體店轉型成功",
        "商戶評價": "5星評價",
    },
    "客戶數據": {
        "客戶性別": "男女均衡",
        "年齡層": "全面滲透各年齡層",
        "重複購買率": "高重複購買率",
        "用戶特徵": "忠實客戶，高回購率",
    },
}

PAIN_POINT_MAPPING = {
    "缺乏線上流量": "HKTVmall擁有310萬+獨立用戶及160萬月活設備，能為商戶帶來巨大線上流量曝光，解決流量獲取難題。",
    "自建物流成本過高": "HKTVmall自建7大倉庫、350+輛貨車及75家O2O門市，提供全港最大住宅配送網絡，讓商戶告別自建物流的高昂成本。",
    "冷鏈配送常有客訴": "HKTVmall配備專業冷鏈配送系統，每日處理10萬+訂單，自動化倉庫確保生鮮及冷藏商品品質，大幅降低客訴率。",
    "營銷成本過高": "HKTVmall提供85折推廣（平台全額承擔）、個人專屬價（20%轉化率）等高效營銷工具，並有鄭裕玲、黃子華等名人廣告加持，讓營銷更具性價比。",
    "庫存管理困難": "GREEN LAB臨期百貨及會員月費計劃能有效幫助商戶清理庫存，降低損耗成本。",
    "缺乏電商營運經驗": "HKTVmall電商學院提供全方位支援，包括上架教學、數據分析、行銷技巧，讓電商新手也能快速上手。",
    "曝光機會不足": "58個MTR站、3,000+廣告燈箱的全渠道曝光，加上直播頻道，為商戶帶來前所未有的品牌曝光機會。",
}

CATEGORY_DATA = {
    "超級市場": {
        "gmv佔比": "最高",
        "目標客群": "家庭消費者、重複購買率高",
        "物流要求": "常溫/冷藏/急凍",
        "熱門工具": ["GREEN LAB臨期百貨", "會員月費計劃", "個人專屬價"],
        "優勢話術": "超級市場品類是HKTVmall GMV最高的品類，家庭消費者忠誠度高，复購率達4.7x/季，加入我們立即享受流量紅利。",
    },
    "護理保健": {
        "gmv佔比": "中高",
        "目標客群": "女性消費者、年輕白領、注重健康人士",
        "物流要求": "常溫",
        "熱門工具": ["個人專屬價", "直播頻道", "ChicChat互動工具"],
        "優勢話術": "護理保健產品在HKTVmall增長迅猛，平台用戶男女均衡，個人專屬價工具能精準觸及目標客戶，20%轉化率效果顯著。",
    },
    "護膚化妝": {
        "gmv佔比": "中高",
        "目標客群": "女性消費者、年輕白領",
        "物流要求": "常溫",
        "熱門工具": ["個人專屬價", "直播頻道", "KOL推廣"],
        "優勢話術": "護膚化妝品市場高速增長，直播頻道可展示產品實際效果，透過個人專屬價工具精準觸及目標客戶，大幅提升轉化率。",
    },
    "家品傢俬": {
        "gmv佔比": "中",
        "目標客群": "家庭消費者、置業人士",
        "物流要求": "常溫/大型",
        "熱門工具": ["直播頻道", "橫幅廣告", "個人專屬價"],
        "優勢話術": "家品傢俬需求穩定，直播可展示產品實際擺放效果，配合個人專屬價優惠，有效吸引家庭客戶下单。",
    },
    "女士服飾": {
        "gmv佔比": "中",
        "目標客群": "女性消費者、年輕時尚",
        "物流要求": "常溫",
        "熱門工具": ["直播頻道", "個人專屬價", "KOL推廣"],
        "優勢話術": "女性服飾市場龐大，直播展示穿搭效果，配合個人專屬價優惠，20%轉化率效果顯著，吸引時尚女性客户。",
    },
    "男士服飾": {
        "gmv佔比": "中低",
        "目標客群": "男性消費者、專業人士",
        "物流要求": "常溫",
        "熱門工具": ["個人專屬價", "橫幅廣告", "CRM精準廣告"],
        "優勢話術": "男士服飾需求稳定，透过CRM精准定位职场男士，配合个人专属价优惠，有效提升复购率。",
    },
    "母嬰育兒": {
        "gmv佔比": "中高",
        "目標客群": "新手父母、家庭客戶",
        "物流要求": "常溫/冷藏",
        "熱門工具": ["個人專屬價", "直播頻道", "會員月費計劃"],
        "優勢話術": "母嬰育兒市場需求殷切，家庭消費者忠誠度高，會員月費計劃能綁定長期客戶，配合個人專屬價優惠大幅提升轉化率。",
    },
    "童裝 嬰兒服飾": {
        "gmv佔比": "中",
        "目標客群": "新手父母、家庭客戶",
        "物流要求": "常溫",
        "熱門工具": ["個人專屬價", "直播頻道", "會員月費計劃"],
        "優勢話術": "童裝及嬰兒服飾市場穩定增長，家庭消費者忠诚度高，配合直播展示穿著效果，有效吸引年輕父母群體。",
    },
    "大腦場": {
        "gmv佔比": "中",
        "目標客群": "大眾消費者",
        "物流要求": "常溫",
        "熱門工具": ["直播頻道", "橫幅廣告", "個人專屬價"],
        "優勢話術": "HKTVmall 310萬+用戶覆蓋全港各類消費者，全方位營銷工具助您快速打開市場，3-5個工作天即可開店營運。",
    },
    "家居電器": {
        "gmv佔比": "中高",
        "目標客群": "家庭消費者、科技愛好者",
        "物流要求": "常溫/大型",
        "熱門工具": ["直播頻道", "橫幅廣告", "關鍵字廣告"],
        "優勢話術": "蘇寧等大型品牌已在HKTVmall直播單晚創下近90萬GMV，家居電器產品極適合直播展示功能，搶佔高端客戶群。",
    },
    "運動旅行": {
        "gmv佔比": "中",
        "目標客群": "運動愛好者、旅行人士",
        "物流要求": "常溫/大型",
        "熱門工具": ["個人專屬價", "直播頻道", "KOL推廣"],
        "優勢話術": "運動及旅行用品需求持續增長，直播展示產品功能效果，配合個人專屬價優惠，有效吸引運動及旅遊愛好者。",
    },
    "寵物用品": {
        "gmv佔比": "中高",
        "目標客群": "寵物主人、家庭客戶",
        "物流要求": "常溫",
        "熱門工具": ["直播頻道", "個人專屬價", "CRM精準廣告"],
        "優勢話術": "寵物市場快速增長，直播頻道可展示寵物用品實際使用效果，透過CRM精準定位寵物主人群體，大幅提升轉化率。",
    },
    "玩具圖書": {
        "gmv佔比": "中",
        "目標客群": "家庭客戶、兒童家長",
        "物流要求": "常溫",
        "熱門工具": ["個人專屬價", "直播頻道", "會員月費計劃"],
        "優勢話術": "玩具及圖書市場穩定，家庭消費者忠诚度高，會員月費計劃能綁定家長客戶，配合個人專屬價優惠大幅提升轉化率。",
    },
    "吃喝玩樂電子券": {
        "gmv佔比": "中低",
        "目標客群": "年輕消費者、追求優惠用戶",
        "物流要求": "常溫/電子券",
        "熱門工具": ["個人專屬價", "直播頻道", "ChicChat互動工具"],
        "優勢話術": "電子券品類增長迅速，直播頻道可展示優惠內容，個人專屬價工具精準觸及追求優惠的年輕客戶群，轉化率效果顯著。",
    },
}

TONE_OPTIONS = [
    "誠懇專業型",
    "強勢說服型",
    "溫暖關懷型",
    "數據導向型",
    "簡潔利落型",
    "輕鬆友好型",
]

FORMAT_OPTIONS = [
    "Pitching Email",
    "WhatsApp短訊（200字以內）",
    "電話推銷話術",
    "社交媒體帖文",
    "商業提案摘要",
]

# ──────────────────────────────────────────────
# Prompt 拼接引擎
# ──────────────────────────────────────────────
def build_system_prompt(merchant_name, category, pain_points, custom_prompt, tone, output_format):
    kb = KNOWLEDGE_BASE
    cat_data = CATEGORY_DATA.get(category, next(iter(CATEGORY_DATA.values())))
    pain_texts = [PAIN_POINT_MAPPING[p] for p in pain_points if p in PAIN_POINT_MAPPING]

    prompt_parts = []
    prompt_parts.append(f"你是一位資深的HKTVmall招商專員，請根據以下信息生成一段針對目標商戶的招商推銷內容。\n")
    prompt_parts.append(f"【目標商戶】{merchant_name}")
    prompt_parts.append(f"【產品品類】{category}")
    prompt_parts.append(f"【品類特徵】GMV佔比：{cat_data['gmv佔比']}｜目標客群：{cat_data['目標客群']}｜物流要求：{cat_data['物流要求']}")
    prompt_parts.append(f"【品類專屬話術】{cat_data['優勢話術']}")

    if pain_texts:
        prompt_parts.append(f"【商戶痛點及對應解決方案】\n" + "\n".join(f"• {t}" for t in pain_texts))

    prompt_parts.append(f"【HKTVmall核心數據】")
    prompt_parts.append(f"• 獨立用戶：{kb['平台概覽']['獨立用戶']}")
    prompt_parts.append(f"• 月活躍設備：{kb['平台概覽']['月活躍設備']}")
    prompt_parts.append(f"• 每季購買頻率：{kb['平台概覽']['每季購買頻率']}")
    prompt_parts.append(f"• 2025總GMV：{kb['平台概覽']['2025總GMV']}")
    prompt_parts.append(f"• 倉庫網絡：{kb['物流網絡']['倉庫數量']}")
    prompt_parts.append(f"• 配送車隊：{kb['物流網絡']['貨車數量']} + {kb['物流網絡']['O2O門市']}")
    prompt_parts.append(f"• 每日訂單處理：{kb['物流網絡']['每日訂單處理量']}")
    prompt_parts.append(f"• 加盟年費：{kb['收費模式']['常規商戶加盟年費']}")
    prompt_parts.append(f"• 個人專屬價轉化率：20%")
    prompt_parts.append(f"• 開店速度：3-5個工作天")
    prompt_parts.append(f"• 貨款結算：15個工作日內")
    prompt_parts.append(f"• 海外市場：澳門、澳洲、加拿大、英國")

    prompt_parts.append(f"\n【輸出格式要求】{output_format}")
    prompt_parts.append(f"【語氣風格】{tone}")
    if custom_prompt.strip():
        prompt_parts.append(f"\n【使用者自訂指示】\n{custom_prompt.strip()}")

    return "\n".join(prompt_parts)


# ──────────────────────────────────────────────
# 本地結構化生成（無需 API key）
# ──────────────────────────────────────────────
def generate_recruitment_content(merchant_name, category, pain_points, custom_prompt, tone, output_format):
    kb = KNOWLEDGE_BASE
    cat_data = CATEGORY_DATA.get(category, next(iter(CATEGORY_DATA.values())))
    pain_texts = [PAIN_POINT_MAPPING[p] for p in pain_points if p in PAIN_POINT_MAPPING]

    openings = {
        "誠懇專業型": f"您好，{merchant_name}團隊，\n\n非常高興有機會向您介紹 HKTVmall — 香港 No.1 網上購物平台，以及我們為您量身訂造的商戶合作方案。",
        "強勢說服型": f"{merchant_name}，您好！\n\n我必須直接告訴您：在香港電商市場，已經有超過310萬消費者在HKTVmall購物，而您的品牌還未入駐，這是一個不容忽視的商機。",
        "溫暖關懷型": f"您好呀，{merchant_name}！\n\n知道您一直在用心經營品牌，我們很希望能把 HKTVmall 的資源帶給您，讓您的生意事半功倍。",
        "數據導向型": f"{merchant_name} 商戶負責人您好，\n\n以下是一份基於HKTVmall真實平台數據的商戶合作方案，供您決策參考：",
        "簡潔利落型": f"{merchant_name}您好，HKTVmall誠邀入駐，重點如下：",
        "輕鬆友好型": f"嘿，{merchant_name}！\n\n有個超棒的消息要告訴您 — HKTVmall 正在招募優質商戶，我覺得您超適合！",
    }

    pain_section = ""
    if pain_texts:
        pain_lines = "\n".join(f"  ✦ {t}" for t in pain_texts)
        pain_section = f"\n針對您可能面臨的挑戰，我們有以下解決方案：\n{pain_lines}\n"

    fee_section = f"""
【入駐費用 — 簡單透明】
  ✦ 加盟年費：{kb['收費模式']['常規商戶加盟年費']}（含廣告金回贈）
  ✦ 銷售佣金：根據品類計算，具競爭力
  ✦ 數碼引流計劃：社交媒體引流享特惠佣金率

【開店流程 — 快捷簡便】
  ✦ 只需：商業登記證 + 銀行月結單
  ✦ 開店時間：3至5個工作天
  ✦ 截單時間：每日下午1點，次日送達

【結算服務 — 準時可靠】
  ✦ 每月結算，15個工作日內款項直入商戶帳戶
"""

    ctas = {
        "誠懇專業型": "\n如您對入駐HKTVmall有任何疑問，歡迎隨時與我聯絡，我們的招商團隊將全程為您提供支援。期待與您合作，共創佳績！",
        "強勢說服型": "\n現在就行動吧！立即聯絡我們，3-5個工作天內即可開店營運，在310萬用戶面前展示您的品牌！商機稍縱即逝。",
        "溫暖關懷型": "\n隨時歡迎您提出任何問題，我會全程耐心跟進。希望有機會與您合作，讓您的品牌被更多香港家庭看見 💚",
        "數據導向型": "\n立即聯絡HKTVmall招商團隊，開啟您的電商增長之旅。",
        "簡潔利落型": "\n聯絡我們，立即入駐！",
        "輕鬆友好型": "\n快點找我聊聊吧，我超期待看到您的品牌登陸HKTVmall！",
    }

    custom_section = ""
    if custom_prompt.strip():
        custom_section = f"\n\n【特別提示】以下為根據您的指示額外調整：\n{custom_prompt.strip()}"

    content = f"""
{'='*60}
HKTVmall 商戶招商提案
{'='*60}

{openings.get(tone, openings["誠懇專業型"])}

{pain_section}
【為何選擇 HKTVmall？】
  ✦ 香港No.1網上購物平台 — {kb['平台概覽']['獨立用戶']}獨立用戶、{kb['平台概覽']['月活躍設備']}月活設備
  ✦ 每季購買頻率 {kb['平台概覽']['每季購買頻率']} — 忠實客戶，高重複購買率
  ✦ 2025年總GMV達{kb['平台概覽']['2025總GMV']}，平台持續高速增長
  ✦ {cat_data['優勢話術']}

【強大物流支援 — 告別配送煩惱】
  ✦ 7大倉庫覆蓋全港：{kb['物流網絡']['倉庫數量']}
  ✦ 350+輛貨車 + 75家O2O門市：全港最大住宅配送網絡
  ✦ 每日處理{kb['物流網絡']['每日訂單處理量']}，德國自動化系統高效可靠
  ✦ 專業冷鏈配送，生鮮冷藏無客訴

【全方位營銷工具 — 提升銷售額】
  ✦ HKTVmall全額承擔85折推廣，單日銷售額可突破4,000萬
  ✦ 個人專屬價：AI個性化優惠，轉化率高達20%
  ✦ 名人代言（鄭裕玲、黃子華）+ 58個MTR站全渠道廣告
  ✦ 直播頻道：單場直播可售出477+件商品
  ✦ GREEN LAB臨期百貨：有效清理庫存，降低損耗

{fee_section}

【成功案例】
  ✦ 蘇寧：直播單晚GMV接近90萬，產品數量1,700+件
  ✦ 中小品牌：2025年Q1已售出1,000+件
  ✦ 小品牌：業績年增長3倍，5星評價

{custom_section}

{'─'*60}
{ctas.get(tone, ctas["誠懇專業型"])}
{'─'*60}

📞 WhatsApp查詢：+852 5283 4138
📧 電郵：business@hktvmall.com
🌐 網站：https://business.hktvmall.com
"""

    return content.strip()


# ──────────────────────────────────────────────
# Streamlit UI（HKTVmall 綠 #106946 + 橙 #E47D21）
# ──────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="HKTVmall B2B 智慧招商工具",
        page_icon="🛒",
        layout="wide",
    )

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&display=swap');
    * { font-family: 'Noto Sans TC', 'PingFang TC', 'Microsoft JhengHei', sans-serif; }

    .hero-banner {
        background: linear-gradient(135deg, #106946 0%, #0D5238 45%, #E47D21 100%);
        padding: 1.5rem 2rem 1.5rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        margin-top: 0.5rem;
        color: white;
        overflow: hidden;
        position: relative;
        z-index: 0;
    }
    .hero-title { font-size: 1.8rem; font-weight: 700; margin-bottom: 0.3rem; }
    .hero-sub { font-size: 0.95rem; opacity: 0.88; }

    .metric-card {
        background: #f8faf8;
        border-left: 4px solid #106946;
        padding: 0.8rem 1rem;
        border-radius: 8px;
        margin-bottom: 0.5rem;
        font-size: 0.88rem;
    }
    .section-header {
        font-size: 0.75rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #106946;
        margin: 1rem 0 0.5rem;
        border-bottom: 2px solid #E47D21;
        padding-bottom: 4px;
    }
    .output-box {
        background: #fafafa;
        border: 1.5px solid #e0e0e0;
        border-radius: 12px;
        padding: 1.5rem;
        font-size: 0.9rem;
        line-height: 1.7;
        white-space: pre-wrap;
        min-height: 400px;
    }
    div[data-testid="stMainBlockContainer"] { padding-top: 1rem; }
    .stButton>button {
        background: linear-gradient(135deg, #106946, #0D5238) !important;
        color: white !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        padding: 0.6rem 2rem !important;
        border: 2px solid #E47D21 !important;
        border-radius: 10px !important;
        width: 100% !important;
        transition: all 0.3s !important;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #E47D21, #F5A623) !important;
        box-shadow: 0 4px 15px rgba(228,125,33,0.4) !important;
    }
    div[data-testid="stExpander"] summary { font-weight: 600; color: #106946; }
    </style>
    """, unsafe_allow_html=True)

    # Hero
    st.markdown("""
    <div class="hero-banner">
        <div class="hero-title">🛒 HKTVmall B2B 智慧招商動態生成工具</div>
        <div class="hero-sub">基於「2026 HKTVmall New Merchant Acquisition Deck [CHI]」| 內置310萬+用戶數據 | 即時生成招商提案</div>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 1.4], gap="large")

    with col_left:
        st.markdown('<div class="section-header">📋 商戶資料輸入</div>', unsafe_allow_html=True)

        merchant_name = st.text_input(
            "商戶名稱",
            placeholder="例如：永明凍肉有限公司",
            help="輸入目標招商商戶的全稱或簡稱"
        )

        category = st.selectbox(
            "產品品類",
            options=list(CATEGORY_DATA.keys()),
            help="選擇商戶主要銷售的產品品類"
        )

        st.markdown("#### 痛點標籤（可多選）")
        pain_options = list(PAIN_POINT_MAPPING.keys())
        selected_pains = []
        for i in range(0, len(pain_options), 2):
            cols = st.columns(2)
            for j, opt in enumerate(pain_options[i:i+2]):
                with cols[j]:
                    if st.checkbox(opt, value=False, key=f"pain_{opt}"):
                        selected_pains.append(opt)

        tone = st.selectbox("語氣風格", options=TONE_OPTIONS, index=0)
        output_format = st.selectbox("輸出格式", options=FORMAT_OPTIONS, index=0)

        st.markdown("#### ✍️ 自訂 Prompt 指令")
        custom_prompt = st.text_area(
            "在這裡輸入您的特殊指示",
            placeholder="例如：請用非常誠懇但強勢的語氣撰寫、強調我們能幫他清臨期百貨庫存的優勢、生成一封適合 WhatsApp 傳送的 200 字短訊息...",
            height=120,
            help="這段文字會附加到生成Prompt中，用於特殊風格或格式要求"
        )

    with col_right:
        st.markdown('<div class="section-header">📊 平台核心數據速覽</div>', unsafe_allow_html=True)

        m1, m2 = st.columns(2)
        m1.markdown(f"""<div class="metric-card"><b>獨立用戶</b><br>{KNOWLEDGE_BASE['平台概覽']['獨立用戶']}</div>""", unsafe_allow_html=True)
        m2.markdown(f"""<div class="metric-card"><b>月活躍設備</b><br>{KNOWLEDGE_BASE['平台概覽']['月活躍設備']}</div>""", unsafe_allow_html=True)
        m3, m4 = st.columns(2)
        m3.markdown(f"""<div class="metric-card"><b>2025 GMV</b><br>{KNOWLEDGE_BASE['平台概覽']['2025總GMV']}</div>""", unsafe_allow_html=True)
        m4.markdown(f"""<div class="metric-card"><b>每季購買頻率</b><br>{KNOWLEDGE_BASE['平台概覽']['每季購買頻率']}</div>""", unsafe_allow_html=True)
        m5, m6 = st.columns(2)
        m5.markdown(f"""<div class="metric-card"><b>倉庫網絡</b><br>{KNOWLEDGE_BASE['物流網絡']['倉庫數量']}</div>""", unsafe_allow_html=True)
        m6.markdown(f"""<div class="metric-card"><b>配送車隊</b><br>{KNOWLEDGE_BASE['物流網絡']['貨車數量']} + O2O門市{KNOWLEDGE_BASE['物流網絡']['O2O門市']}</div>""", unsafe_allow_html=True)
        st.markdown(f"""<div class="metric-card"><b>加盟年費</b><br>{KNOWLEDGE_BASE['收費模式']['常規商戶加盟年費']}（含廣告金回贈）</div>""", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown('<div class="section-header">🤖 AI 招商提案生成區</div>', unsafe_allow_html=True)

        generate_btn = st.button("🚀  生成招商提案")

        if generate_btn:
            if not merchant_name.strip():
                st.warning("⚠️ 請先輸入商戶名稱")
            else:
                with st.spinner("正在交叉匹配數據並生成提案..."):
                    result = generate_recruitment_content(
                        merchant_name=merchant_name,
                        category=category,
                        pain_points=selected_pains,
                        custom_prompt=custom_prompt,
                        tone=tone,
                        output_format=output_format,
                    )
                st.markdown(f'<div class="output-box">{result}</div>', unsafe_allow_html=True)
                filename = f"HKTVmall_招商提案_{merchant_name}.txt"
                st.download_button(
                    label="📥 下載提案文字檔",
                    data=result,
                    file_name=filename,
                    mime="text/plain",
                )
        else:
            st.info("👈 請在左側填寫商戶資料，點擊「生成招商提案」按鈕開始。")

    with st.expander("📚 完整內置知識庫（點擊展開）"):
        for section, items in KNOWLEDGE_BASE.items():
            st.markdown(f"**{section}**")
            for k, v in items.items():
                st.markdown(f"  • **{k}**：{v}")
            st.markdown("")


if __name__ == "__main__":
    main()
