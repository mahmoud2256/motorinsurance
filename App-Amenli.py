# Amenli-style Quote + Myshurance Admin - Motor / Medical / Home / Business
# ملف واحد Streamlit - بيشتغل لوحده من غير Supabase
import io
import uuid
from datetime import date, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Myshurance - Get a Quote", page_icon="🛡", layout="wide", initial_sidebar_state="collapsed")

# ── ترجمة كاملة: عربي أو English ──
T = {
 "ar": {
  "tagline": "اشتري تأمينك في 3 خطوات", "start": "يلا نبدأ", "first": "أولاً،",
  "q_type": "بتدور على تأمين إيه؟", "motor": "تأمين العربيات", "medical": "تأمين طبي للشركات",
  "home": "تأمين المنزل", "business": "تأمين الشركات", "life": "تأمين الحياة",
  "how": "إزاي تشتري؟", "s1": "الخطوة 1", "s1_t": "قولنا بتدور على إيه..", "s1_d": "هنجاوبك على شوية أسئلة بسيطة عشان نرشحلك أحسن وثيقة تناسبك",
  "s2": "الخطوة 2", "s2_t": "قارن بين اختياراتك..", "s2_d": "هنشرحلك الوثائق بوضوح - ومصلحتك دايماً الأول",
  "s3": "الخطوة 3", "s3_t": "اختار وأمّن", "s3_d": "اختار وثيقتك وطريقة الدفع والخطة اللي تناسبك",
  "car_title": "تأمين العربيات", "car_sub": "تغطية لخسايرك لو عربيتك - جديدة أو مستعملة - عملت حادثة. وثائق من أحسن شركات التأمين في مصر.",
  "cov_title": "التغطية:", "cov1": "دفع مباشر", "cov2": "تحمل 0% في التوكيل", "cov3": "تحمل 0% في مراكز الصيانة",
  "cov4": "سرقة أو هلاك كلي", "cov5": "مساعدة على الطريق لو عطلت", "cov_more": "ومزايا تانية زي:",
  "cov6": "خصومات على قطع الغيار", "cov7": "خصومات على الصيانة",
  "did": "عارف؟", "did_d": "32,000 شخص بيتصابوا في حوادث الطرق في مصر كل سنة.",
  "claim_help": "تقدر تقدم مطالبتك في أي وقت وأي يوم، وهنساعدك ونتكلم مع شركة التأمين بالنيابة عنك.",
  "get_policy": "خد وثيقتك دلوقتي", "back": "رجوع", "next": "كمل", "compare": "قارن الأسعار",
  "car_brand": "ماركة العربية", "car_model": "الموديل", "car_year": "سنة الصنع", "car_value": "قيمة العربية (جنيه)",
  "car_new": "العربية جديدة ولا مستعملة؟", "new": "جديدة", "used": "مستعملة",
  "car_fuel": "عربيتك كهربا ولا بنزين؟", "electric": "كهربا", "fuel": "بنزين", "hybrid": "هايبرد (كهربا + بنزين)",
  "name": "اسمك", "phone": "رقم التليفون", "nid": "الرقم القومي", "plate": "رقم اللوحة (اختياري)",
  "your_quote": "عرض السعر بتاعك", "provider": "شركة التأمين", "premium": "القسط السنوي", "deduct": "التحمل",
  "choose": "اختار الوثيقة دي", "chosen": "تم اختيار الوثيقة: {p}", "pay": "طريقة الدفع",
  "pay_full": "كاش - سنة كاملة", "pay_half": "نص سنوي", "pay_month": "شهري", "confirm": "أكّد واطلب الوثيقة",
  "done": "تم استلام طلبك! هنكلمك في خلال يوم عمل واحد.", "ref": "رقم طلبك",
  "admin": "دخول الإدارة", "admin_title": "طلبات الأسعار - الإدارة", "all_quotes": "كل طلبات الأسعار",
  "export": "تصدير Excel", "search": "بحث", "status": "الحالة", "new_q": "جديد", "contacted": "تم التواصل", "issued": "تم الإصدار",
  "move_to": "نقل إلى", "readonly": "حسابك للعرض فقط.", "logout": "خروج", "login": "دخول", "pw": "كلمة المرور",
  "role": "الصلاحية", "admin_r": "مدير", "viewer_r": "عرض فقط", "bad_pw": "كلمة المرور غير صحيحة",
  "user_name": "اسمك", "demo": "تجريبي: مدير / myshurance123",
 },
 "en": {
  "tagline": "Buy insurance in 3 steps", "start": "Let's get you started.", "first": "First off,",
  "q_type": "What type of insurance are you looking for?", "motor": "Motor Insurance", "medical": "Business Medical Insurance",
  "home": "Home Insurance", "business": "Business Insurance", "life": "Life Insurance",
  "how": "How it works", "s1": "STEP 1", "s1_t": "Tell us what you're looking for..", "s1_d": "We'll ask a few simple questions to recommend the best policies that suit your needs",
  "s2": "STEP 2", "s2_t": "Compare your options..", "s2_d": "We'll guide and explain the policies clearly - always with your best interest in mind",
  "s3": "STEP 3", "s3_t": "Now choose and insure", "s3_d": "Pick your policy and select the payment method and plan that suits you most",
  "car_title": "Car Insurance", "car_sub": "Car insurance provides coverage for losses if your car, new or used, is in an accident. Policies from the best insurance companies in Egypt.",
  "cov_title": "Insurance Coverage:", "cov1": "Direct billing options", "cov2": "Up to 0% deductible at Agency", "cov3": "Up to 0% deductible at Service Centers",
  "cov4": "Theft or total loss", "cov5": "Roadside assistance if you breakdown on the road", "cov_more": "In addition to further benefits like:",
  "cov6": "Discounts on spare parts", "cov7": "Discounts on car maintenance",
  "did": "Did you know..", "did_d": "32,000 get hurt in Egyptian road accidents each year.",
  "claim_help": "You can make your claim any time any day and we'll support you and handle all talks with the insurance providers on your behalf.",
  "get_policy": "Get your policy now", "back": "Back", "next": "Continue", "compare": "Compare prices",
  "car_brand": "Car brand", "car_model": "Model", "car_year": "Year", "car_value": "Car value (EGP)",
  "car_new": "Is the car new or used?", "new": "New", "used": "Used",
  "car_fuel": "Is your car electric or fuel?", "electric": "Electric", "fuel": "Fuel", "hybrid": "Hybrid (Electric + Fuel)",
  "name": "Your name", "phone": "Phone number", "nid": "National ID", "plate": "Plate number (optional)",
  "your_quote": "Your quote", "provider": "Insurance company", "premium": "Yearly premium", "deduct": "Deductible",
  "choose": "Choose this policy", "chosen": "Policy chosen: {p}", "pay": "Payment method",
  "pay_full": "Cash - full year", "pay_half": "Semi-annual", "pay_month": "Monthly", "confirm": "Confirm and request policy",
  "done": "Your request is received! We will call you within 1 working day.", "ref": "Your request no",
  "admin": "Admin login", "admin_title": "Quote requests - Admin", "all_quotes": "All quote requests",
  "export": "Export Excel", "search": "Search", "status": "Status", "new_q": "New", "contacted": "Contacted", "issued": "Issued",
  "move_to": "Move to", "readonly": "Your account is read-only.", "logout": "Logout", "login": "Login", "pw": "Password",
  "role": "Role", "admin_r": "Admin", "viewer_r": "Viewer", "bad_pw": "Wrong password",
  "user_name": "Your name", "demo": "Demo: admin / myshurance123",
 },
}
PROVIDERS = [
    {"name": "Allianz", "rate": 0.032, "deduct": "0%", "note": "0% at Agency"},
    {"name": "AXA", "rate": 0.029, "deduct": "0%", "note": "0% at Service Centers"},
    {"name": "Orient", "rate": 0.027, "deduct": "500 EGP", "note": "Direct billing"},
    {"name": "Sarwa", "rate": 0.025, "deduct": "1,000 EGP", "note": "Roadside assistance"},
    {"name": "GIG", "rate": 0.024, "deduct": "1,500 EGP", "note": "Spare parts discount"},
]
BRANDS = ["Toyota", "Hyundai", "Kia", "Nissan", "MG", "Chery", "Suzuki", "Renault", "Peugeot", "Chevrolet", "Other"]

def lang():
    return "ar" if st.session_state.get("lang_label", "العربية") == "العربية" else "en"
def t(k):
    return T[lang()].get(k, k)
def secret(n):
    try: return st.secrets.get(n, "")
    except Exception: return ""
def to_excel(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w: df.to_excel(w, index=False)
    buf.seek(0); return buf.getvalue()

TEAL, DARK, SAND, CREAM = "#0B4F55", "#07363B", "#E8C9A0", "#F6F1E8"
def css():
    c = f"""
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
html, body, [class*="css"], .stApp, button, input, textarea, select {{ font-family: 'Cairo', sans-serif !important; }}
.stApp {{ background: {CREAM}; {'direction: rtl; text-align: right;' if lang()=='ar' else ''} }}
.block-container {{ max-width: 900px; padding-top: 1.2rem; }}
.hero {{ background: linear-gradient(135deg, {TEAL} 0%, #226068 100%); color:#fff; border-radius: 26px; padding: 28px 30px; }}
.hero h1 {{ color:#fff; font-size: 34px; font-weight:800; margin:0; }}
.hero p {{ opacity:.9; margin: 8px 0 0 0; }}
.card {{ background:#fff; border:1px solid #E8DCC8; border-radius:20px; padding:20px; margin:10px 0; }}
.card h3 {{ color:{TEAL}; margin:0 0 6px 0; }}
.opt {{ background:#fff; border:2px solid #E8DCC8; border-radius:18px; padding:18px; text-align:center; cursor:pointer; }}
.opt:hover {{ border-color:{TEAL}; }}
.stButton > button, .stFormSubmitButton > button, .stDownloadButton > button {{ background:{TEAL}; color:#fff; border:0; border-radius:12px; padding:.6rem 1.4rem; font-weight:700; }}
[data-testid="stForm"] {{ background:#fff; border:1px solid #E8DCC8; border-radius:20px; }}
"""
    st.markdown(f"<style>{c}</style>", unsafe_allow_html=True)

def init():
    ss = st.session_state
    ss.setdefault("step", "type")
    ss.setdefault("quotes", pd.DataFrame(columns=["ref","type","name","phone","nid","car","value","provider","premium","pay","status","date"]))

def header():
    c1, c2 = st.columns([5, 1])
    with c1: st.markdown("**Myshurance**  ·  🛡")
    with c2: st.radio("lang", ["English", "العربية"], horizontal=True, key="lang_label", label_visibility="collapsed")

# ── الصفحة 1: اختار نوع التأمين - زي qe.amenli.com ──
def page_type():
    st.markdown(f'<div class="hero"><h1>{t("start")}</h1><p>{t("first")}</p><h1 style="font-size:28px;margin-top:10px">{t("q_type")}</h1></div>', unsafe_allow_html=True)
    st.write("")
    opts = [("motor", "🚗", t("motor")), ("medical", "🏥", t("medical")), ("home", "🏠", t("home")), ("business", "🏢", t("business"))]
    cols = st.columns(2)
    for i, (key, icon, label) in enumerate(opts):
        with cols[i % 2]:
            if st.button(f"{icon}\n\n{label}", key=f"type_{key}", width="stretch"):
                st.session_state.q_type = key
                st.session_state.step = "detail" if key == "motor" else "coming"
                st.rerun()
    st.write("")
    st.subheader(t("how"))
    s1, s2, s3 = st.columns(3)
    for col, sk, tk, dk in [(s1, "s1", "s1_t", "s1_d"), (s2, "s2", "s2_t", "s2_d"), (s3, "s3", "s3_t", "s3_d")]:
        with col: st.markdown(f'<div class="card"><small>{t(sk)}</small><h3>{t(tk)}</h3><p>{t(dk)}</p></div>', unsafe_allow_html=True)
    st.caption("—")
    if st.button(f"🔐 {t('admin')}"):
        st.session_state.step = "admin_login"; st.rerun()

# ── صفحة العربيات: تفاصيل + تغطية - زي amenli.com/en/services/motorinsurance ──
def page_motor_detail():
    st.markdown(f'<div class="hero"><h1>🚗 {t("car_title")}</h1><p>{t("car_sub")}</p></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="card"><h3>{t("cov_title")}</h3><p>✓ {t("cov1")}<br>✓ {t("cov2")}<br>✓ {t("cov3")}<br>✓ {t("cov4")}<br>✓ {t("cov5")}</p><p><b>{t("cov_more")}</b><br>✓ {t("cov6")}<br>✓ {t("cov7")}</p></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="card"><h3>💡 {t("did")}</h3><p>{t("did_d")}</p><p>{t("claim_help")}</p></div>', unsafe_allow_html=True)
    if st.button(f"{t('back')}"):
        st.session_state.step = "type"; st.rerun()
    with st.form("motor_form"):
        r1 = st.columns(2)
        brand = r1[0].selectbox(t("car_brand"), BRANDS); model = r1[1].text_input(t("car_model"))
        r2 = st.columns(2)
        year = r2[0].number_input(t("car_year"), min_value=2000, max_value=2027, value=2023, step=1)
        value = r2[1].number_input(t("car_value"), min_value=50000, max_value=10000000, value=800000, step=10000)
        cond = st.radio(t("car_new"), [t("new"), t("used")], horizontal=True)
        fuel = st.radio(t("car_fuel"), [t("electric"), t("fuel"), t("hybrid")], horizontal=True)
        r3 = st.columns(2)
        name = r3[0].text_input(t("name")); phone = r3[1].text_input(t("phone"))
        r4 = st.columns(2)
        nid = r4[0].text_input(t("nid")); plate = r4[1].text_input(t("plate"))
        if st.form_submit_button(t("compare")):
            if not (name.strip() and phone.strip()):
                st.error("—")
            else:
                st.session_state.quote = {"type": "motor", "name": name, "phone": phone, "nid": nid, "plate": plate,
                                          "car": f"{brand} {model} {year}", "value": float(value), "cond": cond, "fuel": fuel, "date": str(date.today())}
                st.session_state.step = "compare"; st.rerun()

# ── صفحة المقارنة - زي Amenli Compare ──
def page_compare():
    q = st.session_state.quote
    st.markdown(f'<div class="hero"><h1>{t("your_quote")}</h1><p>{q["car"]} - {q["value"]:,.0f} - {q.get("fuel","")}</p></div>', unsafe_allow_html=True)
    if st.button(t("back")):
        st.session_state.step = "detail"; st.rerun()
    st.caption(f"{q['name']} · {q['phone']}")
    for p in PROVIDERS:
        prem = round(q["value"] * p["rate"])
        # خصم لو العربية جديدة
        if q["cond"] == t("new"): prem = round(prem * 0.95)
        # تعديل حسب نوع الوقود: كهربا أغلى 10% (بطارية)، هايبرد أغلى 5%
        if q.get("fuel") == t("electric"): prem = round(prem * 1.10)
        elif q.get("fuel") == t("hybrid"): prem = round(prem * 1.05)
        with st.container():
            st.markdown(f'<div class="card"><h3>{p["name"]}</h3><p>{p["note"]} · {t("deduct")}: {p["deduct"]}</p><h2 style="color:{TEAL}">{prem:,.0f} جنيه / سنة</h2><p>✓ {t("cov1")} · ✓ {t("cov4")} · ✓ {t("cov5")}</p></div>', unsafe_allow_html=True)
            if st.button(f"{t('choose')} - {p['name']}", key=f"ch_{p['name']}"):
                st.session_state.quote["provider"] = p["name"]
                st.session_state.quote["premium"] = prem
                st.session_state.step = "checkout"; st.rerun()

# ── صفحة الاختيار والدفع ──
def page_checkout():
    q = st.session_state.quote
    st.markdown(f'<div class="hero"><h1>{t("chosen").format(p=q["provider"])}</h1><p>{q["car"]} - {q.get("fuel","")} - {q["premium"]:,.0f}</p></div>', unsafe_allow_html=True)
    pay = st.radio(t("pay"), [t("pay_full"), t("pay_half"), t("pay_month")])
    if pay == t("pay_half"): total = q["premium"]
    elif pay == t("pay_month"): total = round(q["premium"] / 12)
    else: total = q["premium"]
    st.markdown(f'<div class="card"><h3>{pay}</h3><h2 style="color:{TEAL}">{total:,.0f}</h2></div>', unsafe_allow_html=True)
    if st.button(t("back")):
        st.session_state.step = "compare"; st.rerun()
    if st.button(t("confirm")):
        ref = f"Q-{len(st.session_state.quotes)+1:04d}"
        row = {"ref": ref, "type": q["type"], "name": q["name"], "phone": q["phone"], "nid": q["nid"], "fuel": q.get("fuel",""),
               "car": q["car"], "value": q["value"], "provider": q["provider"], "premium": q["premium"],
               "pay": pay, "status": "new", "date": q["date"]}
        st.session_state.quotes = pd.concat([st.session_state.quotes, pd.DataFrame([row])], ignore_index=True)
        st.session_state.last_ref = ref
        st.session_state.step = "done"; st.rerun()

def page_done():
    st.markdown(f'<div class="hero"><h1>✅ {t("done")}</h1><p>{t("ref")}: {st.session_state.get("last_ref","")}</p></div>', unsafe_allow_html=True)
    if st.button(t("start")):
        st.session_state.step = "type"; st.rerun()

def page_coming():
    st.markdown(f'<div class="hero"><h1>{t("q_type")}</h1><p>{t("s1_d")}</p></div>', unsafe_allow_html=True)
    st.info("— " + (st.session_state.get("q_type","")))
    if st.button(t("back")):
        st.session_state.step = "type"; st.rerun()

# ── إدارة ──
def page_admin_login():
    st.markdown(f'<div class="hero"><h1>🔐 {t("admin")}</h1></div>', unsafe_allow_html=True)
    st.caption(t("demo"))
    with st.form("alogin"):
        role = st.selectbox(t("role"), ["admin", "viewer"], format_func=lambda r: t("admin_r") if r=="admin" else t("viewer_r"))
        name = st.text_input(t("user_name")); typed = st.text_input(t("pw"), type="password")
        if st.form_submit_button(t("login")):
            pw = {"admin": secret("ADMIN_PASSWORD") or "myshurance123", "viewer": secret("VIEWER_PASSWORD") or "viewer123"}
            import hmac as _hm
            if _hm.compare_digest(typed.encode(), pw[role].encode()):
                st.session_state.role, st.session_state.user_name = role, name.strip()
                st.session_state.step = "admin"; st.rerun()
            else: st.error(t("bad_pw"))
    if st.button(t("back")):
        st.session_state.step = "type"; st.rerun()

def page_admin():
    st.markdown(f'<div class="hero"><h1>{t("admin_title")}</h1><p>{t("user_name")}: {st.session_state.get("user_name","")}</p></div>', unsafe_allow_html=True)
    df = st.session_state.quotes
    if df.empty: st.info("—")
    else:
        q = st.text_input(t("search"))
        v = df
        if q:
            hay = v.astype(str).agg(" ".join, axis=1)
            v = v[hay.str.contains(q, case=False, na=False, regex=False)]
        st.dataframe(v, hide_index=True, width="stretch")
        st.download_button(t("export"), to_excel(v), "quotes.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if st.session_state.get("role") == "admin" and not v.empty:
            lab = dict(zip(v["ref"], v["ref"].astype(str)+" | "+v["name"].astype(str)))
            ref = st.selectbox(t("ref"), list(lab), format_func=lab.get)
            cols = st.columns(3)
            for col, s in zip(cols, ["new", "contacted", "issued"]):
                if col.button(f"{t('move_to')}: {t('new_q') if s=='new' else t('contacted') if s=='contacted' else t('issued')}", key=f"mv_{ref}_{s}"):
                    idx = df[df["ref"]==ref].index[0]
                    st.session_state.quotes.at[idx, "status"] = s; st.rerun()
    if st.button(t("logout")):
        for k in ("role","user_name"): st.session_state.pop(k, None)
        st.session_state.step = "type"; st.rerun()

def main():
    init(); css(); header()
    step = st.session_state.step
    {"type": page_type, "detail": page_motor_detail, "compare": page_compare, "checkout": page_checkout,
     "done": page_done, "coming": page_coming, "admin_login": page_admin_login, "admin": page_admin}[step]()

main()
