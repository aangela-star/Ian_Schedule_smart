import streamlit as st

# 1. 設定頁面 (這行一定要在最上面)
st.set_page_config(
    page_title="晉安毅安聯合排班系統",
    page_icon="🏥",
    layout="wide"
)

# ==========================================
# 🔒 安全守門員：登入檢查系統
# ==========================================
def check_password():
    """如果使用者輸入正確密碼，回傳 True，否則回傳 False"""

    def password_entered():
        """檢查使用者輸入的密碼是否與 secrets 中的設定相符"""
        if st.session_state["password"] == st.secrets["LOGIN_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 驗證後刪除輸入框的暫存，保持乾淨
        else:
            st.session_state["password_correct"] = False

    # 初始化 session state
    if "password_correct" not in st.session_state:
        # 第一次進入，顯示輸入框
        st.text_input(
            "請輸入系統密碼 / Password", type="password", on_change=password_entered, key="password"
        )
        return False
    
    elif not st.session_state["password_correct"]:
        # 密碼錯誤，再次顯示輸入框
        st.text_input(
            "❌ 密碼錯誤，請重試 / Password", type="password", on_change=password_entered, key="password"
        )
        return False
    
    else:
        # 密碼正確
        return True

# 🚨 執行檢查：如果沒通過，程式就停在這裡 (st.stop)
if not check_password():
    st.stop()

# ==========================================
# 👇 只有登入成功後，才會執行下面的程式碼
# ==========================================

st.title("🏥 晉安毅安 聯合智慧排班入口")
st.markdown("---")

st.info("👋 歡迎回來！身分驗證成功，請從左側選單開始作業。")

col1, col2 = st.columns(2)

with col1:
    st.header("🏥 復健部")
    st.write("包含：PT/OT 排班、瀑布流運算、三診制支援")
    st.write("👉 請點擊左側 **復健部排班**")

with col2:
    st.header("💉 護理部")
    st.write("包含：N1/N2/N3 輪替、行政優先權、護理長與PT支援")
    st.write("👉 請點擊左側 **護理部排班**")

st.markdown("---")
st.caption("© 2026 晉安毅安醫療體系 | IT 部門開發")