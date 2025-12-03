import streamlit as st
import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold
import pandas as pd
import io
import random
import datetime
import traceback

# 預先檢查環境
try:
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("請先安裝 openpyxl: pip install openpyxl")

# --- 1. 頁面設定 ---
st.set_page_config(page_title="專業 Excel 生成器", page_icon="📊", layout="centered")

if 'user_prompt' not in st.session_state:
    st.session_state['user_prompt'] = ''

# --- 2. 標題 ---
st.title("📊 AI Excel 專業生成器")
st.markdown("專為 Excel 小白設計的救星！內建 **AI 自我修復機制**，大幅降低出錯率。")

# --- 3. 側邊欄 ---
with st.sidebar:
    st.header("🔑 啟動金鑰")
    api_key = None
    try:
        if "GEMINI_API_KEY" in st.secrets:
            api_key = st.secrets["GEMINI_API_KEY"]
            st.success("✅ 已連接內建金鑰")
    except: pass

    if not api_key:
        api_key = st.text_input("請在此輸入 Gemini API Key", type="password", placeholder="AIzaSy...")
        with st.expander("❓ 如何免費獲取 API Key？"):
            st.markdown("[👉 點此前往 Google AI Studio](https://aistudio.google.com/app/apikey)")
    
    st.divider()
    
    # ⚡ 懶人樣板按鈕 (都在！)
    st.write("⚡ **快速樣板 (點擊自動填寫)：**")
    if st.button("💰 個人記帳表"): st.session_state['user_prompt'] = "幫我做一個2025年個人記帳表。欄位：日期、類別、項目、金額、付款方式。請生成10筆範例。公式要求：計算本月總支出、分類小計。美化：標題深藍底白字，金額加$符號。"
    if st.button("📦 商品庫存表"): st.session_state['user_prompt'] = "幫我做一個庫存管理表。欄位：商品編號、名稱、進貨價、售價、庫存量、庫存總值(公式：進貨價*庫存量)。請生成10筆範例。美化：標題深綠底，金額加千分位。"
    if st.button("🛒 網拍訂單表"): st.session_state['user_prompt'] = "幫我做一個電商訂單管理表。欄位：訂單編號、平台(蝦皮/官網)、商品、單價、數量、手續費(蝦皮8%/官網2%)、實收金額。請生成10筆。公式要求：用IF判斷手續費，退貨實收為0。美化：標題亮橘底。"

    st.divider()
    model_choice = st.selectbox("模型選擇", ["gemini-2.5-flash", "gemini-2.5-pro"])

# --- 4. 核心邏輯：記憶體直出 + 安全解鎖 + 自我修復 ---
def generate_excel_buffer(user_prompt, key, model_name):
    try:
        genai.configure(api_key=key)
        
        # 安全設定解鎖
        safety_settings = {
            HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
            HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
        }

        model = genai.GenerativeModel(
            model_name,
            generation_config=genai.types.GenerationConfig(
                max_output_tokens=8000,
                temperature=0.0 
            )
        ) 
        
        base_prompt = f"""
        你是一位 Python Excel 自動化專家。需求："{user_prompt}"
        
        請寫一段 **完整且可執行** 的 Python 代碼。
        
        【嚴格代碼規範】：
        1. **Imports**：務必包含 `io`, `random`, `datetime`, `pandas` 以及 `openpyxl` 相關模組。
        2. **核心邏輯**：
           - 建立 `wb = Workbook()`
           - 填入資料與公式。
           - 進行美化 (樣式定義)。
           - **關鍵步驟 (OUTPUT)**：
             最後請將檔案儲存到變數 `output_buffer` 中，不要存成檔案！
             範例：
             ```python
             output_buffer = io.BytesIO()
             wb.save(output_buffer)
             output_buffer.seek(0)
             ```
        3. **禁止事項**：不要使用 `wb.save('file.xlsx')`。只輸出 Python 代碼。不要使用 openpyxl.formatting。
        """
        
        current_prompt = base_prompt
        max_retries = 3 
        
        for attempt in range(max_retries):
            response = model.generate_content(current_prompt, safety_settings=safety_settings)
            
            if not response.parts:
                return None, f"AI 拒絕生成 (Finish Reason: {response.candidates[0].finish_reason})。"
                
            raw_code = response.text
            clean_code = raw_code.replace("```python", "").replace("```", "").strip()
            if not clean_code.startswith("import") and not clean_code.startswith("from"):
                 import_pos = clean_code.find("import")
                 if import_pos != -1: clean_code = clean_code[import_pos:]
            
            try:
                # 測試執行
                test_vars = {}
                exec(clean_code, globals(), test_vars)
                
                # 確認有產出 buffer
                if 'output_buffer' in test_vars:
                    return clean_code, None
                else:
                    raise Exception("代碼執行成功但未產生 output_buffer 變數")
                    
            except Exception as e:
                error_msg = str(e)
                print(f"第 {attempt+1} 次嘗試失敗: {error_msg}")
                current_prompt += f"\n\n\n【系統回報】：錯誤訊息為：{error_msg}。\n請修正代碼，確保最後將結果存入 `output_buffer = io.BytesIO()`。"
                
        return None, "AI 嘗試修復了 3 次但仍然失敗，請嘗試簡化您的需求。"
        
    except Exception as e:
        return None, str(e)

# --- 5. 主介面 ---

# 🔥 V4.6 保證：好壞範例教學完整保留！
with st.expander("💡 怎麼樣才能做出完美的表格？ (點我看秘訣)"):
    st.markdown("""
    **黃金許願公式：**
    > **我要做什麼表 + 有哪些欄位 + 資料邏輯/公式 + 美化要求**
    
    **❌ 壞範例：**
    "幫我做一個記帳表。" (AI 不知道你要記什麼，可能會做得很簡陋)
    
    **✅ 好範例：**
    "幫我做一個**家庭收支表**。
    欄位要有：**日期、項目、金額、類別**。
    請幫我造 **10 筆** 隨機資料。
    最下面要用公式幫我算**總金額**。
    標題請用**深綠色底**，金額要有**錢字號**。"
    """)

# 使用 session_state 綁定輸入框
user_input = st.text_area("請輸入需求 (或點擊左側快速樣板)：", value=st.session_state['user_prompt'], height=150, placeholder="例如：幫我做一個房東收租表...")

if st.button("✨ 生成專業表格 (自動修復模式)", type="primary"):
    if not api_key:
        st.error("❌ 請先輸入 API Key")
    elif not user_input:
        st.warning("⚠️ 請輸入需求")
    else:
        spinner_text = f"🤖 AI 正在製作中 (已解除安全限制)..."
        with st.spinner(spinner_text):
            
            # 1. 獲取代碼
            code, error_msg = generate_excel_buffer(user_input, api_key, model_choice)
            
            if code:
                try:
                    # 2. 正式執行
                    local_vars = {}
                    exec(code, globals(), local_vars)
                    
                    if 'output_buffer' in local_vars:
                        excel_data = local_vars['output_buffer']
                        
                        st.download_button(
                            label="📥 下載 Excel (.xlsx)",
                            data=excel_data,
                            file_name=f"excel_gen_{datetime.datetime.now().strftime('%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("🎉 完成！這是全新生成的資料。")
                    
                except Exception as e:
                    st.error(f"未知錯誤：{e}")
                    with st.expander("查看代碼"):
                        st.code(code, language='python')
            else:
                st.error("連線或修復失敗！")
                st.error(error_msg)

# --- 6. 頁尾 ---
st.divider()
st.caption("Excel Generator V4.6 (The Real Ultimate Ver.)")
