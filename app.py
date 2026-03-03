import streamlit as st
import sqlite3
import pandas as pd
from openai import OpenAI

st.set_page_config(
    page_title="Ozon — Юнит-экономика FBO/FBS",
    layout="wide",
    page_icon="📦"
)

DB_PATH = "products_storage.db"

def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sku TEXT UNIQUE,
            name TEXT,
            length_cm REAL,
            width_cm REAL,
            height_cm REAL,
            weight_kg REAL,
            cost REAL DEFAULT 0
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS ai_cache (
            name TEXT,
            client TEXT,
            category TEXT,
            PRIMARY KEY (name, client)
        )
    """)
    conn.commit()
    return conn

def normalize_value(raw, unit):
    try:
        v = float(str(raw).replace(",", ".").strip())
    except (ValueError, TypeError):
        return 0.0
    u = str(unit).strip().lower()
    if u in ("мм", "mm"): return v / 10.0
    if u in ("г", "g", "гр", "gr"): return v / 1000.0
    return v

def get_ai_category(name: str, categories: list, conn, client_key: str) -> str:
    c = conn.cursor()
    row = c.execute(
        "SELECT category FROM ai_cache WHERE name=? AND client=?",
        (name, client_key)
    ).fetchone()
    if row: return row[0]
    api_key = st.session_state.get("openai_key", "")
    if not api_key or not categories: return categories[0] if categories else "Неизвестно"
    try:
        client = OpenAI(api_key=api_key)
        cats_str = "\n".join(f"- {cat}" for cat in categories)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": f"Ты классификатор товаров для маркетплейса {client_key}. Выбери ОДНУ категорию из списка. Ответь ТОЛЬКО её названием."},
                {"role": "user", "content": f"Товар: {name}\nКатегории:\n{cats_str}"}
            ],
            max_tokens=60,
            temperature=0
        )
        category = resp.choices[0].message.content.strip()
        if category not in categories: category = categories[0]
    except Exception:
        category = categories[0] if categories else "Неизвестно"
    c.execute("INSERT OR REPLACE INTO ai_cache (name, client, category) VALUES (?,?,?)", (name, client_key, category))
    conn.commit()
    return category

def calc_tax(revenue: float, cost_total: float, regime: str):
    profit_before = revenue - cost_total
    rates = {
        "ОСНО (25% от прибыли)": ("profit", 0.25),
        "УСН Доходы (6%)": ("revenue", 0.06),
        "УСН Доходы-Расходы (15%)": ("profit", 0.15),
        "АУСН (8% от дохода)": ("revenue", 0.08),
        "УСН с НДС 5%": ("revenue", 0.05),
        "УСН с НДС 7%": ("revenue", 0.07),
    }
    mode, rate = rates.get(regime, ("profit", 0.0))
    if mode == "revenue": tax = revenue * rate
    else: tax = max(profit_before * rate, 0)
    profit_after = profit_before - tax
    margin_after = (profit_after / revenue * 100) if revenue > 0 else 0
    return round(tax, 2), round(profit_after, 2), round(margin_after, 1)

# --- Ozon Logistics ---
CATEGORY_COMMISSIONS = {
    "Электроника": 15.0,
    "Смартфоны": 10.0,
    "Одежда и обувь": 25.0,
    "Спорт и отдых": 15.0,
    "Товары для детей": 18.0,
    "Прочее": 20.0,
}

def get_logistics_fbo(vol_liters, weight_kg, days=45):
    processing = 30.0
    delivery = 50.0 + max(0, weight_kg - 1) * 5.0 + max(0, vol_liters - 5) * 10.0
    storage_days = max(0, days - 14)
    storage_cost = storage_days * vol_liters * 5.0
    return processing, delivery, storage_cost

def get_logistics_fbs(vol_liters, weight_kg):
    processing = 25.0
    delivery = 60.0 + max(0, weight_kg - 1) * 7.0 + max(0, vol_liters - 5) * 12.0
    return processing, delivery, 0.0

# --- Main App ---
conn = init_db()

if "openai_key" not in st.session_state:
    st.session_state["openai_key"] = st.secrets.get("OPENAI_API_KEY", "")

st.header("Ozon — Расчёт целевой цены (Юнит-экономика)")

with st.sidebar:
    st.subheader("⚙️ Параметры расчёта")
    tax_regime = st.selectbox("Система налогообложения", [
        "ОСНО (25% от прибыли)", "УСН Доходы (6%)", "УСН Доходы-Расходы (15%)",
        "АУСН (8% от дохода)", "УСН с НДС 5%", "УСН с НДС 7%"
    ])
    st.divider()
    st.subheader("📊 Параметры менеджера")
    buyout = st.slider("Выкупаемость, %", 10, 100, 85)
    defect = st.slider("Брак, %", 0, 20, 2)
    ad = st.slider("Реклама, %", 0, 50, 10)
    boost = st.slider("Буст продаж, %", 0, 20, 5)
    target_margin = st.slider("Целевая маржа, %", 0, 50, 20)

# Catalog Management
with st.expander("Блок 1. Каталог товаров", expanded=True):
    col1, col2 = st.columns(2)
    with col1: dim_unit = st.selectbox("Размеры", ["см", "мм"])
    with col2: wt_unit = st.selectbox("Вес", ["кг", "г"])
    uploaded = st.file_uploader("Загрузить Excel (SKU, Название, Длина, Ширина, Высота, Вес, Себестоимость)", type=["xlsx"])
    if uploaded:
        df = pd.read_excel(uploaded)
        if st.button("Сохранить в базу"):
            for _, row in df.iterrows():
                try:
                    sku = str(row.get('SKU', row.get('Артикул', ''))).strip()
                    name = str(row.get('Название', row.get('Наименование', ''))).strip()
                    l = normalize_value(row.get('Длина', 0), dim_unit)
                    w = normalize_value(row.get('Ширина', 0), dim_unit)
                    h = normalize_value(row.get('Высота', 0), dim_unit)
                    wt = normalize_value(row.get('Вес', 0), wt_unit)
                    cost = float(str(row.get('Себестоимость', 0)).replace(',', '.'))
                    conn.execute("INSERT OR REPLACE INTO products (sku, name, length_cm, width_cm, height_cm, weight_kg, cost) VALUES (?,?,?,?,?,?,?)",
                                 (sku, name, l, w, h, wt, cost))
                except: continue
            conn.commit()
            st.success("Каталог обновлен")
    all_p = pd.read_sql("SELECT * FROM products", conn)
    st.dataframe(all_p)

# Calculation
st.subheader("Блок 2. Расчёт юнит-экономики")
if not all_p.empty:
    col_a, col_b = st.columns(2)
    with col_a:
        model = st.radio("Модель работы", ["FBO", "FBS"], horizontal=True)
    with col_b:
        storage_days = st.number_input("Срок хранения (дней, для FBO)", 0, 365, 45)

    if st.button("🚀 Рассчитать для всего каталога"):
        results = []
        cat_list = list(CATEGORY_COMMISSIONS.keys())
        buyout_rate = buyout / 100.0
        defect_rate = defect / 100.0
        adv_rate = ad / 100.0
        boost_rate = boost / 100.0
        target_m = target_margin / 100.0

        for _, p in all_p.iterrows():
            vol_l = (p['length_cm'] * p['width_cm'] * p['height_cm']) / 1000.0
            if model == "FBO":
                proc, deliv, stor = get_logistics_fbo(vol_l, p['weight_kg'], storage_days)
            else:
                proc, deliv, stor = get_logistics_fbs(vol_l, p['weight_kg'])
            logistics_fix = proc + deliv + stor

            cat = get_ai_category(p['name'], cat_list, conn, "ozon")
            comm_rate = CATEGORY_COMMISSIONS.get(cat, 20.0) / 100.0
            total_pct = comm_rate + adv_rate + boost_rate + target_m
            denom = (1 - total_pct) * buyout_rate * (1 - defect_rate)
            price = (p['cost'] + logistics_fix) / denom if denom > 0 else 0
            tax, profit, marg = calc_tax(price, p['cost'] + logistics_fix + (price * (total_pct - target_m)), tax_regime)

            results.append({
                "SKU": p['sku'], "Название": p['name'], "Модель": model,
                "Категория": cat, "Комиссия %": int(comm_rate * 100),
                "Логистика, руб": round(logistics_fix, 0),
                "Цена на полке, руб": round(price, 0),
                "Прибыль, руб": profit, "Маржа %": marg
            })
        res_df = pd.DataFrame(results)
        st.dataframe(res_df, use_container_width=True)
        st.download_button("Скачать (CSV)", res_df.to_csv(index=False).encode("utf-8"), "ozon_results.csv", mime="text/csv")
else:
    st.info("Загрузите каталог товаров для расчёта.")
