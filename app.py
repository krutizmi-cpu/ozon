import math
import sqlite3
from io import BytesIO

import pandas as pd
import streamlit as st
from openai import OpenAI

st.set_page_config(
    page_title="Ozon — Юнит-экономика FBO/FBS",
    layout="wide",
    page_icon="📦"
)

DB_PATH = "products_storage.db"


# =========================
# DB
# =========================
def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sku TEXT UNIQUE,
            name TEXT,
            category_manual TEXT,
            length_cm REAL,
            width_cm REAL,
            height_cm REAL,
            weight_kg REAL,
            cost REAL DEFAULT 0,
            price_regular REAL DEFAULT 0,
            price_promo REAL DEFAULT 0
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


# =========================
# Utils
# =========================
def clean_num(raw, default=0.0):
    if raw is None:
        return default
    try:
        if pd.isna(raw):
            return default
    except Exception:
        pass

    try:
        return float(str(raw).replace(" ", "").replace(",", ".").strip())
    except Exception:
        return default


def normalize_dimension(raw, unit):
    value = clean_num(raw, 0.0)
    if str(unit).strip().lower() in ("мм", "mm"):
        return value / 10.0
    return value


def normalize_weight(raw, unit):
    value = clean_num(raw, 0.0)
    if str(unit).strip().lower() in ("г", "гр", "g", "gr"):
        return value / 1000.0
    return value


def safe_round(value, ndigits=2):
    try:
        if value is None:
            return 0
        value = float(value)
        if math.isnan(value) or math.isinf(value):
            return 0
        return round(value, ndigits)
    except Exception:
        return 0


def to_excel_bytes(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    return output.getvalue()


# =========================
# Templates
# =========================
DEFAULT_CATEGORY_RATES = {
    "Смартфоны": 10.0,
    "Электроника": 15.0,
    "Бытовая техника": 15.0,
    "Компьютеры и комплектующие": 12.0,
    "Одежда и обувь": 22.0,
    "Красота и здоровье": 12.0,
    "Дом и сад": 16.0,
    "Детские товары": 18.0,
    "Спорт и отдых": 15.0,
    "Автотовары": 15.0,
    "Канцтовары": 17.0,
    "Прочее": 20.0,
}


def build_catalog_template():
    return pd.DataFrame([
        {
            "SKU": "SKU-001",
            "Название": "Смартфон 128GB черный",
            "Категория Ozon": "Смартфоны",
            "Длина, см": 18,
            "Ширина, см": 9,
            "Высота, см": 5,
            "Вес, кг": 0.45,
            "Себестоимость, ₽": 12000,
            "Цена без акции, ₽": 18990,
            "Цена акции, ₽": 16990,
        },
        {
            "SKU": "SKU-002",
            "Название": "Футболка хлопковая мужская",
            "Категория Ozon": "Одежда и обувь",
            "Длина, см": 30,
            "Ширина, см": 25,
            "Высота, см": 3,
            "Вес, кг": 0.20,
            "Себестоимость, ₽": 450,
            "Цена без акции, ₽": 1490,
            "Цена акции, ₽": 1190,
        }
    ])


def build_category_rates_template():
    return pd.DataFrame(
        [{"Категория Ozon": k, "Комиссия, %": v} for k, v in DEFAULT_CATEGORY_RATES.items()]
    )


def build_instructions_template():
    return pd.DataFrame([
        {
            "Поле": "SKU",
            "Описание": "Артикул товара. Обязательное поле.",
            "Пример": "SKU-001"
        },
        {
            "Поле": "Название",
            "Описание": "Название товара. Используется для автоопределения категории.",
            "Пример": "Смартфон 128GB черный"
        },
        {
            "Поле": "Категория Ozon",
            "Описание": "Можно заполнить вручную. Если пусто — приложение попробует определить категорию автоматически.",
            "Пример": "Смартфоны"
        },
        {
            "Поле": "Длина, см / Ширина, см / Высота, см",
            "Описание": "Габариты товара. Нужны для логистики.",
            "Пример": "18 / 9 / 5"
        },
        {
            "Поле": "Вес, кг",
            "Описание": "Фактический вес товара.",
            "Пример": "0.45"
        },
        {
            "Поле": "Себестоимость, ₽",
            "Описание": "Полная закупочная / производственная себестоимость одной единицы.",
            "Пример": "12000"
        },
        {
            "Поле": "Цена без акции, ₽",
            "Описание": "Базовая цена до скидок / акций.",
            "Пример": "18990"
        },
        {
            "Поле": "Цена акции, ₽",
            "Описание": "Фактическая цена продажи при акции. Если пусто — берётся цена без акции.",
            "Пример": "16990"
        },
        {
            "Поле": "Итоговый расчёт",
            "Описание": "В отчёте будут текущая цена, рекомендованная цена, прибыль, маржинальность, наценка, комиссии, логистика, возвраты и т.д.",
            "Пример": "-"
        }
    ])


# =========================
# Categories
# =========================
CATEGORY_KEYWORDS = {
    "Смартфоны": [
        "смартфон", "iphone", "xiaomi", "redmi", "realme", "galaxy", "mobile phone"
    ],
    "Электроника": [
        "наушник", "гарнитур", "колонка", "power bank", "пауэрбанк", "зарядк",
        "кабель", "видеорегистратор", "роутер", "bluetooth"
    ],
    "Бытовая техника": [
        "чайник", "пылесос", "блендер", "утюг", "кофевар", "кофемаш", "мультиварк",
        "микроволнов", "увлажнитель", "обогреватель"
    ],
    "Компьютеры и комплектующие": [
        "ноутбук", "монитор", "ssd", "hdd", "клавиатур", "мышь", "видеокарта",
        "оперативная память", "ram", "материнская плата", "процессор"
    ],
    "Одежда и обувь": [
        "футболка", "худи", "толстовка", "джинсы", "куртка", "пальто", "кроссовки",
        "ботинки", "носки", "майка", "рубашка", "брюки"
    ],
    "Красота и здоровье": [
        "крем", "шампунь", "бальзам", "сыворотка", "маска", "парфюм",
        "духи", "гель для душа", "витамин", "массажер"
    ],
    "Дом и сад": [
        "контейнер", "посуда", "сковород", "кастрюл", "подушка", "одеяло",
        "простын", "полотенце", "штора", "органайзер", "лампа"
    ],
    "Детские товары": [
        "детск", "игрушк", "коляска", "подгуз", "пеленк", "конструктор", "кукла"
    ],
    "Спорт и отдых": [
        "гантел", "коврик", "фитнес", "палатка", "рюкзак", "велосипед", "эспандер"
    ],
    "Автотовары": [
        "авто", "машин", "держатель", "щетка", "чехол на руль", "коврик в авто"
    ],
    "Канцтовары": [
        "тетрад", "ручка", "карандаш", "маркер", "ежедневник", "бумага"
    ],
}


def get_keyword_category(name: str):
    text = str(name).lower().strip()
    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw in text:
                return category
    return None


def get_ai_category(name: str, categories: list, conn, client_key: str) -> str:
    c = conn.cursor()
    row = c.execute(
        "SELECT category FROM ai_cache WHERE name=? AND client=?",
        (name, client_key)
    ).fetchone()
    if row:
        return row[0]

    api_key = st.session_state.get("openai_key", "")
    if not api_key or not categories:
        return "Прочее" if "Прочее" in categories else (categories[0] if categories else "Прочее")

    try:
        client = OpenAI(api_key=api_key)
        cats_str = "\n".join(f"- {cat}" for cat in categories)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "Ты классификатор товаров для маркетплейса Ozon. "
                        "Выбери ровно одну категорию из списка. "
                        "Ответь только названием категории."
                    ),
                },
                {
                    "role": "user",
                    "content": f"Товар: {name}\nКатегории:\n{cats_str}"
                }
            ],
            max_tokens=40,
            temperature=0
        )
        category = resp.choices[0].message.content.strip()
        if category not in categories:
            category = "Прочее" if "Прочее" in categories else categories[0]
    except Exception:
        category = "Прочее" if "Прочее" in categories else (categories[0] if categories else "Прочее")

    c.execute(
        "INSERT OR REPLACE INTO ai_cache (name, client, category) VALUES (?, ?, ?)",
        (name, client_key, category)
    )
    conn.commit()
    return category


def resolve_category(name, manual_category, available_categories, conn):
    if str(manual_category).strip():
        manual = str(manual_category).strip()
        if manual in available_categories:
            return manual, "Из файла"

    by_kw = get_keyword_category(name)
    if by_kw and by_kw in available_categories:
        return by_kw, "По словарю"

    by_ai = get_ai_category(name, available_categories, conn, "ozon")
    if by_ai in available_categories:
        return by_ai, "AI"

    return ("Прочее" if "Прочее" in available_categories else available_categories[0]), "По умолчанию"


# =========================
# Tax
# =========================
def calc_tax(revenue: float, total_cost_before_tax: float, regime: str):
    profit_before_tax = revenue - total_cost_before_tax

    rates = {
        "ОСНО (22% от прибыли)": ("profit", 0.22),
        "УСН Доходы (6%)": ("revenue", 0.06),
        "УСН Доходы-Расходы (15%)": ("profit", 0.15),
        "АУСН Доходы (8%)": ("revenue", 0.08),
        "УСН с НДС 5%": ("revenue", 0.05),
        "УСН с НДС 7%": ("revenue", 0.07),
    }

    mode, rate = rates.get(regime, ("profit", 0.0))

    if mode == "revenue":
        tax = max(revenue, 0) * rate
    else:
        tax = max(profit_before_tax, 0) * rate

    profit_after_tax = profit_before_tax - tax
    profit_pct_of_price = (profit_after_tax / revenue * 100) if revenue > 0 else 0

    return (
        safe_round(tax, 2),
        safe_round(profit_before_tax, 2),
        safe_round(profit_after_tax, 2),
        safe_round(profit_pct_of_price, 2)
    )


# =========================
# Commission rules
# =========================
def get_commission_rate(price_for_commission: float, category: str, category_rates: dict):
    """
    Логика как в предыдущей версии:
    - до 100 ₽: 14%
    - 100+ до 300 ₽: 20%
    - выше 300 ₽: ставка категории
    """
    if price_for_commission <= 100:
        return 0.14
    if price_for_commission <= 300:
        return 0.20
    return category_rates.get(category, category_rates.get("Прочее", 20.0)) / 100.0


# =========================
# Logistics / Returns
# =========================
def calc_logistics(model, volume_liters, weight_kg, storage_days, params):
    if model == "FBO":
        processing = params["fbo_processing_rub"]
        base_delivery = params["fbo_base_delivery_rub"]
        extra_kg_rate = params["fbo_extra_kg_rub"]
        extra_liter_rate = params["fbo_extra_liter_rub"]
    else:
        processing = params["fbs_processing_rub"]
        base_delivery = params["fbs_base_delivery_rub"]
        extra_kg_rate = params["fbs_extra_kg_rub"]
        extra_liter_rate = params["fbs_extra_liter_rub"]

    overweight = max(0.0, weight_kg - params["included_weight_kg"])
    overvolume = max(0.0, volume_liters - params["included_volume_l"])

    delivery = base_delivery + overweight * extra_kg_rate + overvolume * extra_liter_rate

    storage = 0.0
    if model == "FBO":
        paid_days = max(0, storage_days - params["storage_grace_days"])
        storage = paid_days * volume_liters * params["storage_rub_per_liter_day"]

    direct_logistics = processing + delivery
    return (
        safe_round(processing, 2),
        safe_round(delivery, 2),
        safe_round(storage, 2),
        safe_round(direct_logistics, 2)
    )


def calc_returns_and_cancellations(
    direct_logistics_rub: float,
    price_to_customer: float,
    cost: float,
    buyout_rate: float,
    cancellation_rate: float,
    return_logistics_coef: float,
    cancellation_logistics_coef: float,
    return_processing_rub: float,
    defect_on_return_rate: float,
):
    """
    Более точная модель:
    1) cancellation_rate — доля заказов, отменённых до выкупа/получения
    2) buyout_rate — доля выкупленных из неотменённых заказов
    3) обратная логистика отдельно на возвраты и отдельно на отмены
    4) резерв на повреждение / потерю товарного вида при возврате
    """

    cancellation_share = max(0.0, min(1.0, cancellation_rate))
    post_cancel_orders_share = max(0.0, 1.0 - cancellation_share)

    buyout_share_total = post_cancel_orders_share * max(0.0, min(1.0, buyout_rate))
    non_buyout_after_dispatch_share = max(0.0, post_cancel_orders_share - buyout_share_total)

    return_logistics_rub = direct_logistics_rub * non_buyout_after_dispatch_share * return_logistics_coef
    cancellation_logistics_rub = direct_logistics_rub * cancellation_share * cancellation_logistics_coef
    return_processing_total_rub = (non_buyout_after_dispatch_share + cancellation_share) * return_processing_rub
    damage_reserve_on_returns_rub = cost * (non_buyout_after_dispatch_share + cancellation_share) * defect_on_return_rate

    total_reverse_cost = (
        return_logistics_rub
        + cancellation_logistics_rub
        + return_processing_total_rub
        + damage_reserve_on_returns_rub
    )

    effective_sale_probability = buyout_share_total

    return {
        "cancellation_share_pct": safe_round(cancellation_share * 100, 2),
        "effective_buyout_share_pct": safe_round(buyout_share_total * 100, 2),
        "non_buyout_after_dispatch_share_pct": safe_round(non_buyout_after_dispatch_share * 100, 2),
        "return_logistics_rub": safe_round(return_logistics_rub, 2),
        "cancellation_logistics_rub": safe_round(cancellation_logistics_rub, 2),
        "return_processing_rub": safe_round(return_processing_total_rub, 2),
        "damage_reserve_on_returns_rub": safe_round(damage_reserve_on_returns_rub, 2),
        "total_reverse_cost_rub": safe_round(total_reverse_cost, 2),
        "effective_sale_probability": safe_round(effective_sale_probability, 6),
    }


# =========================
# Unit economics
# =========================
def calc_price_metrics(
    base_regular_price: float,
    promo_price: float,
    spp_discount_pct: float,
    cost: float,
    category: str,
    model: str,
    volume_liters: float,
    weight_kg: float,
    storage_days: int,
    category_rates: dict,
    tax_regime: str,
    adv_rate: float,
    boost_rate: float,
    acquiring_rate: float,
    defect_base_rate: float,
    buyout_rate: float,
    cancellation_rate: float,
    other_fixed_rub: float,
    logistics_params: dict,
):
    """
    base_regular_price — цена без акции
    promo_price — цена акции
    spp_discount_pct — скидка маркетплейса / СПП
    Выручка продавца считаем по цене после СПП.
    """
    if promo_price <= 0:
        promo_price = base_regular_price

    if base_regular_price <= 0:
        base_regular_price = promo_price

    spp_rate = max(0.0, min(1.0, spp_discount_pct))
    customer_price = promo_price
    seller_revenue_price = promo_price * (1.0 - spp_rate)

    commission_rate = get_commission_rate(seller_revenue_price, category, category_rates)

    processing_rub, delivery_rub, storage_rub, direct_logistics_rub = calc_logistics(
        model=model,
        volume_liters=volume_liters,
        weight_kg=weight_kg,
        storage_days=storage_days,
        params=logistics_params
    )

    returns_block = calc_returns_and_cancellations(
        direct_logistics_rub=direct_logistics_rub,
        price_to_customer=customer_price,
        cost=cost,
        buyout_rate=buyout_rate,
        cancellation_rate=cancellation_rate,
        return_logistics_coef=logistics_params["return_logistics_coef"],
        cancellation_logistics_coef=logistics_params["cancellation_logistics_coef"],
        return_processing_rub=logistics_params["return_processing_rub"],
        defect_on_return_rate=logistics_params["defect_on_return_rate"],
    )

    commission_rub = seller_revenue_price * commission_rate
    advertising_rub = seller_revenue_price * adv_rate
    boost_rub = seller_revenue_price * boost_rate
    acquiring_rub = seller_revenue_price * acquiring_rate
    base_defect_reserve_rub = cost * defect_base_rate
    marketplace_discount_rub = customer_price - seller_revenue_price

    full_cost_before_tax = (
        cost
        + commission_rub
        + direct_logistics_rub
        + storage_rub
        + returns_block["total_reverse_cost_rub"]
        + advertising_rub
        + boost_rub
        + acquiring_rub
        + base_defect_reserve_rub
        + other_fixed_rub
    )

    tax_rub, profit_before_tax_rub, profit_after_tax_rub, profit_pct_of_revenue = calc_tax(
        revenue=seller_revenue_price,
        total_cost_before_tax=full_cost_before_tax,
        regime=tax_regime
    )

    margin_pct = ((seller_revenue_price / full_cost_before_tax - 1) * 100) if full_cost_before_tax > 0 else 0.0
    markup_to_cost_pct = ((seller_revenue_price / cost - 1) * 100) if cost > 0 else 0.0

    return {
        "regular_price_rub": safe_round(base_regular_price, 2),
        "promo_price_rub": safe_round(promo_price, 2),
        "customer_price_rub": safe_round(customer_price, 2),
        "seller_revenue_price_rub": safe_round(seller_revenue_price, 2),
        "marketplace_discount_rub": safe_round(marketplace_discount_rub, 2),

        "commission_rate_pct": safe_round(commission_rate * 100, 2),
        "commission_rub": safe_round(commission_rub, 2),

        "processing_rub": safe_round(processing_rub, 2),
        "delivery_rub": safe_round(delivery_rub, 2),
        "direct_logistics_rub": safe_round(direct_logistics_rub, 2),
        "storage_rub": safe_round(storage_rub, 2),

        "advertising_rub": safe_round(advertising_rub, 2),
        "boost_rub": safe_round(boost_rub, 2),
        "acquiring_rub": safe_round(acquiring_rub, 2),
        "base_defect_reserve_rub": safe_round(base_defect_reserve_rub, 2),

        "returns_total_rub": safe_round(returns_block["total_reverse_cost_rub"], 2),
        "return_logistics_rub": safe_round(returns_block["return_logistics_rub"], 2),
        "cancellation_logistics_rub": safe_round(returns_block["cancellation_logistics_rub"], 2),
        "return_processing_rub": safe_round(returns_block["return_processing_rub"], 2),
        "damage_reserve_on_returns_rub": safe_round(returns_block["damage_reserve_on_returns_rub"], 2),

        "cancellation_share_pct": returns_block["cancellation_share_pct"],
        "effective_buyout_share_pct": returns_block["effective_buyout_share_pct"],
        "non_buyout_after_dispatch_share_pct": returns_block["non_buyout_after_dispatch_share_pct"],

        "other_fixed_rub": safe_round(other_fixed_rub, 2),
        "full_cost_before_tax": safe_round(full_cost_before_tax, 2),
        "tax_rub": safe_round(tax_rub, 2),
        "profit_before_tax_rub": safe_round(profit_before_tax_rub, 2),
        "profit_after_tax_rub": safe_round(profit_after_tax_rub, 2),
        "profit_pct_of_revenue": safe_round(profit_pct_of_revenue, 2),
        "margin_pct": safe_round(margin_pct, 2),
        "markup_to_cost_pct": safe_round(markup_to_cost_pct, 2),
    }


def find_recommended_price(
    target_margin_pct: float,
    regular_price_reference: float,
    spp_discount_pct: float,
    cost: float,
    category: str,
    model: str,
    volume_liters: float,
    weight_kg: float,
    storage_days: int,
    category_rates: dict,
    tax_regime: str,
    adv_rate: float,
    boost_rate: float,
    acquiring_rate: float,
    defect_base_rate: float,
    buyout_rate: float,
    cancellation_rate: float,
    other_fixed_rub: float,
    logistics_params: dict,
    promo_discount_from_regular_pct: float,
):
    """
    Ищем рекомендованную цену БЕЗ АКЦИИ.
    Акционная цена = regular_price * (1 - promo_discount_from_regular_pct)
    Потом применяем СПП / скидку маркетплейса.
    """

    promo_discount_from_regular_pct = max(0.0, min(1.0, promo_discount_from_regular_pct))

    def get_metrics_for_regular_price(regular_price):
        promo_price = regular_price * (1.0 - promo_discount_from_regular_pct)
        return calc_price_metrics(
            base_regular_price=regular_price,
            promo_price=promo_price,
            spp_discount_pct=spp_discount_pct,
            cost=cost,
            category=category,
            model=model,
            volume_liters=volume_liters,
            weight_kg=weight_kg,
            storage_days=storage_days,
            category_rates=category_rates,
            tax_regime=tax_regime,
            adv_rate=adv_rate,
            boost_rate=boost_rate,
            acquiring_rate=acquiring_rate,
            defect_base_rate=defect_base_rate,
            buyout_rate=buyout_rate,
            cancellation_rate=cancellation_rate,
            other_fixed_rub=other_fixed_rub,
            logistics_params=logistics_params,
        )

    low = max(cost * 0.5, 1.0)
    high = max(regular_price_reference if regular_price_reference > 0 else cost * 3, cost * 10, 1000.0)

    for _ in range(25):
        m = get_metrics_for_regular_price(high)
        if m["margin_pct"] >= target_margin_pct:
            break
        high *= 1.5

    for _ in range(60):
        mid = (low + high) / 2
        m = get_metrics_for_regular_price(mid)
        if m["margin_pct"] >= target_margin_pct:
            high = mid
        else:
            low = mid

    recommended_regular = safe_round(high, 2)
    recommended_promo = safe_round(recommended_regular * (1.0 - promo_discount_from_regular_pct), 2)
    recommended_metrics = get_metrics_for_regular_price(recommended_regular)

    return recommended_regular, recommended_promo, recommended_metrics


def classify_sku_status(margin_pct, profit_rub):
    if profit_rub < 0 or margin_pct < 0:
        return "Критично"
    if margin_pct < 10:
        return "Риск"
    if margin_pct < 20:
        return "Норма"
    return "Хорошо"


# =========================
# Styling
# =========================
def highlight_status(row):
    status = row.get("Статус SKU", "")
    if status == "Критично":
        return ["background-color: #ffd6d6"] * len(row)
    if status == "Риск":
        return ["background-color: #fff0c7"] * len(row)
    if status == "Хорошо":
        return ["background-color: #d9f7df"] * len(row)
    return [""] * len(row)


# =========================
# App
# =========================
conn = init_db()

if "openai_key" not in st.session_state:
    st.session_state["openai_key"] = st.secrets.get("OPENAI_API_KEY", "")

st.title("Ozon — Юнит-экономика FBO / FBS")
st.caption("Расчёт текущей и рекомендованной цены, прибыли, маржинальности, наценки, логистики, возвратов и скидок маркетплейса")


# =========================
# Sidebar
# =========================
with st.sidebar:
    st.subheader("⚙️ Налоги")
    tax_regime = st.selectbox(
        "Система налогообложения",
        [
            "ОСНО (22% от прибыли)",
            "УСН Доходы (6%)",
            "УСН Доходы-Расходы (15%)",
            "АУСН Доходы (8%)",
            "УСН с НДС 5%",
            "УСН с НДС 7%",
        ]
    )

    st.divider()
    st.subheader("📊 Управленческие параметры")

    buyout = st.slider("Выкупаемость после доставки, %", 10, 100, 85)
    cancellation = st.slider("Отмены до выкупа / получения, %", 0, 50, 5)
    defect = st.slider("Базовый резерв на брак / списание, %", 0, 20, 2)
    ad = st.slider("Реклама, % от выручки продавца", 0, 50, 10)
    boost = st.slider("Буст / продвижение, % от выручки продавца", 0, 20, 5)
    acquiring = st.slider("Эквайринг, % от выручки продавца", 0.0, 10.0, 1.5, 0.1)
    spp_discount = st.slider("СПП / скидка маркетплейса, %", 0, 50, 0)
    target_margin = st.slider("Целевая маржинальность, %", 0, 100, 20)
    other_fixed_rub = st.number_input("Прочие расходы на 1 шт., ₽", min_value=0.0, value=0.0, step=10.0)

    st.divider()
    st.subheader("💸 Модель цен")
    promo_discount_from_regular_pct = st.slider(
        "Скидка акции от цены без акции, %",
        0, 80, 10
    )

    st.divider()
    st.subheader("🚚 Параметры логистики")

    included_weight_kg = st.number_input("Вес без доплаты, кг", min_value=0.0, value=1.0, step=0.1)
    included_volume_l = st.number_input("Объём без доплаты, л", min_value=0.0, value=5.0, step=0.5)

    st.markdown("**FBO**")
    fbo_processing_rub = st.number_input("FBO: обработка, ₽", min_value=0.0, value=20.0, step=1.0)
    fbo_base_delivery_rub = st.number_input("FBO: базовая доставка, ₽", min_value=0.0, value=83.0, step=1.0)
    fbo_extra_kg_rub = st.number_input("FBO: доплата за 1 кг сверх порога, ₽", min_value=0.0, value=8.0, step=1.0)
    fbo_extra_liter_rub = st.number_input("FBO: доплата за 1 л сверх порога, ₽", min_value=0.0, value=8.0, step=1.0)

    st.markdown("**FBS**")
    fbs_processing_rub = st.number_input("FBS: обработка, ₽", min_value=0.0, value=20.0, step=1.0)
    fbs_base_delivery_rub = st.number_input("FBS: базовая доставка, ₽", min_value=0.0, value=83.0, step=1.0)
    fbs_extra_kg_rub = st.number_input("FBS: доплата за 1 кг сверх порога, ₽", min_value=0.0, value=8.0, step=1.0)
    fbs_extra_liter_rub = st.number_input("FBS: доплата за 1 л сверх порога, ₽", min_value=0.0, value=8.0, step=1.0)

    storage_grace_days = st.number_input("Льготный срок хранения, дней", min_value=0, value=14, step=1)
    storage_rub_per_liter_day = st.number_input("Хранение, ₽ / 1 л / день", min_value=0.0, value=0.25, step=0.05)

    st.markdown("**Возвраты / отмены**")
    return_logistics_coef = st.number_input("Коэффициент обратной логистики по возврату", min_value=0.0, value=1.0, step=0.1)
    cancellation_logistics_coef = st.number_input("Коэффициент логистики по отмене", min_value=0.0, value=0.5, step=0.1)
    return_processing_rub = st.number_input("Обработка возврата / отмены, ₽", min_value=0.0, value=15.0, step=1.0)
    defect_on_return_rate = st.number_input("Потеря товарного вида на возвратах, доля", min_value=0.0, value=0.05, step=0.01)

    st.divider()
    st.subheader("🤖 AI")
    st.text_input("OpenAI API key", type="password", key="openai_key")


logistics_params = {
    "included_weight_kg": included_weight_kg,
    "included_volume_l": included_volume_l,
    "fbo_processing_rub": fbo_processing_rub,
    "fbo_base_delivery_rub": fbo_base_delivery_rub,
    "fbo_extra_kg_rub": fbo_extra_kg_rub,
    "fbo_extra_liter_rub": fbo_extra_liter_rub,
    "fbs_processing_rub": fbs_processing_rub,
    "fbs_base_delivery_rub": fbs_base_delivery_rub,
    "fbs_extra_kg_rub": fbs_extra_kg_rub,
    "fbs_extra_liter_rub": fbs_extra_liter_rub,
    "storage_grace_days": storage_grace_days,
    "storage_rub_per_liter_day": storage_rub_per_liter_day,
    "return_logistics_coef": return_logistics_coef,
    "cancellation_logistics_coef": cancellation_logistics_coef,
    "return_processing_rub": return_processing_rub,
    "defect_on_return_rate": defect_on_return_rate,
}


# =========================
# Block 1. Templates and uploads
# =========================
with st.expander("Блок 1. Шаблоны, ставки категорий и загрузка каталога", expanded=True):
    st.markdown("### 1.1 Скачать шаблоны")

    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        st.download_button(
            "Шаблон каталога (Excel)",
            data=to_excel_bytes({"catalog_template": build_catalog_template()}),
            file_name="ozon_catalog_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_t2:
        st.download_button(
            "Шаблон ставок категорий (Excel)",
            data=to_excel_bytes({"category_rates_template": build_category_rates_template()}),
            file_name="ozon_category_rates_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col_t3:
        st.download_button(
            "Шаблон с примечаниями для сотрудников (Excel)",
            data=to_excel_bytes({
                "catalog_template": build_catalog_template(),
                "instructions": build_instructions_template(),
                "category_rates_template": build_category_rates_template(),
            }),
            file_name="ozon_template_with_instructions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown("### 1.2 Ручное редактирование ставок категорий")

    if "category_rates_df" not in st.session_state:
        st.session_state["category_rates_df"] = pd.DataFrame(
            [{"Категория Ozon": k, "Комиссия, %": v} for k, v in DEFAULT_CATEGORY_RATES.items()]
        )

    rates_upload = st.file_uploader(
        "Загрузить Excel со ставками категорий",
        type=["xlsx"],
        key="rates_upload"
    )

    if rates_upload is not None:
        try:
            rates_df_loaded = pd.read_excel(rates_upload)
            rates_df_loaded.columns = [str(c).strip() for c in rates_df_loaded.columns]
            if "Категория Ozon" in rates_df_loaded.columns and "Комиссия, %" in rates_df_loaded.columns:
                st.session_state["category_rates_df"] = rates_df_loaded[["Категория Ozon", "Комиссия, %"]].copy()
                st.success("Ставки категорий загружены")
            else:
                st.warning("В файле ставок должны быть колонки 'Категория Ozon' и 'Комиссия, %'")
        except Exception as e:
            st.error(f"Ошибка чтения файла ставок: {e}")

    edited_rates_df = st.data_editor(
        st.session_state["category_rates_df"],
        num_rows="dynamic",
        use_container_width=True,
        key="category_rates_editor"
    )

    st.session_state["category_rates_df"] = edited_rates_df.copy()

    category_rates = {}
    for _, row in edited_rates_df.iterrows():
        cat = str(row.get("Категория Ozon", "")).strip()
        rate = clean_num(row.get("Комиссия, %", 0), 0.0)
        if cat:
            category_rates[cat] = rate

    if "Прочее" not in category_rates:
        category_rates["Прочее"] = DEFAULT_CATEGORY_RATES["Прочее"]

    st.markdown("### 1.3 Загрузка каталога")

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        dim_unit = st.selectbox("Единица размеров в исходном файле", ["см", "мм"])
    with col_u2:
        wt_unit = st.selectbox("Единица веса в исходном файле", ["кг", "г"])

    uploaded_catalog = st.file_uploader(
        "Excel: SKU / Название / Категория Ozon / Длина / Ширина / Высота / Вес / Себестоимость / Цена без акции / Цена акции",
        type=["xlsx"],
        key="catalog_upload"
    )

    if uploaded_catalog is not None:
        raw_df = pd.read_excel(uploaded_catalog)
        st.write("Предпросмотр загруженного файла:")
        st.dataframe(raw_df.head(20), use_container_width=True)

        if st.button("Сохранить каталог в базу", type="primary"):
            inserted = 0
            for _, row in raw_df.iterrows():
                try:
                    sku = str(row.get("SKU", row.get("Артикул", ""))).strip()
                    name = str(row.get("Название", row.get("Наименование", ""))).strip()
                    category_manual = str(row.get("Категория Ozon", "")).strip()

                    l = normalize_dimension(row.get("Длина, см", row.get("Длина", 0)), dim_unit)
                    w = normalize_dimension(row.get("Ширина, см", row.get("Ширина", 0)), dim_unit)
                    h = normalize_dimension(row.get("Высота, см", row.get("Высота", 0)), dim_unit)
                    wt = normalize_weight(row.get("Вес, кг", row.get("Вес", 0)), wt_unit)

                    cost = clean_num(row.get("Себестоимость, ₽", row.get("Себестоимость", 0)), 0.0)
                    price_regular = clean_num(
                        row.get("Цена без акции, ₽", row.get("Цена без акции", row.get("Цена", 0))), 0.0
                    )
                    price_promo = clean_num(
                        row.get("Цена акции, ₽", row.get("Цена акции", 0)), 0.0
                    )

                    if not sku or not name:
                        continue

                    conn.execute("""
                        INSERT OR REPLACE INTO products
                        (sku, name, category_manual, length_cm, width_cm, height_cm, weight_kg, cost, price_regular, price_promo)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        sku, name, category_manual, l, w, h, wt, cost, price_regular, price_promo
                    ))
                    inserted += 1
                except Exception:
                    continue

            conn.commit()
            st.success(f"Каталог обновлён. Загружено строк: {inserted}")

    all_products = pd.read_sql("SELECT * FROM products ORDER BY id DESC", conn)
    st.markdown("### 1.4 Каталог в базе")
    st.dataframe(all_products, use_container_width=True)


# =========================
# Block 2. Calculation
# =========================
st.subheader("Блок 2. Расчёт юнит-экономики")

if all_products.empty:
    st.info("Сначала загрузите каталог товаров.")
else:
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        model = st.radio("Модель работы", ["FBO", "FBS"], horizontal=True)
    with col_c2:
        storage_days = st.number_input("Срок хранения, дней", min_value=0, max_value=365, value=45, step=1)

    if st.button("🚀 Рассчитать для всего каталога"):
        buyout_rate = buyout / 100.0
        cancellation_rate = cancellation / 100.0
        defect_base_rate = defect / 100.0
        adv_rate = ad / 100.0
        boost_rate = boost / 100.0
        acquiring_rate = acquiring / 100.0
        spp_discount_pct = spp_discount / 100.0
        promo_discount_from_regular = promo_discount_from_regular_pct / 100.0

        results = []
        available_categories = list(category_rates.keys())

        for _, p in all_products.iterrows():
            sku = str(p.get("sku", "")).strip()
            name = str(p.get("name", "")).strip()
            manual_category = str(p.get("category_manual", "") or "").strip()

            category, category_source = resolve_category(
                name=name,
                manual_category=manual_category,
                available_categories=available_categories,
                conn=conn
            )

            length_cm = clean_num(p.get("length_cm", 0), 0.0)
            width_cm = clean_num(p.get("width_cm", 0), 0.0)
            height_cm = clean_num(p.get("height_cm", 0), 0.0)
            weight_kg = clean_num(p.get("weight_kg", 0), 0.0)
            cost = clean_num(p.get("cost", 0), 0.0)
            price_regular = clean_num(p.get("price_regular", 0), 0.0)
            price_promo = clean_num(p.get("price_promo", 0), 0.0)

            if price_regular <= 0 and price_promo > 0:
                price_regular = price_promo / (1.0 - promo_discount_from_regular) if promo_discount_from_regular < 1 else price_promo

            if price_promo <= 0 and price_regular > 0:
                price_promo = price_regular * (1.0 - promo_discount_from_regular)

            volume_liters = 0.0
            if length_cm > 0 and width_cm > 0 and height_cm > 0:
                volume_liters = (length_cm * width_cm * height_cm) / 1000.0

            # Если цены вообще нет — строим от рекомендованной
            if price_regular <= 0 and price_promo <= 0:
                rec_regular, rec_promo, rec_metrics = find_recommended_price(
                    target_margin_pct=target_margin,
                    regular_price_reference=max(cost * 3, 1000),
                    spp_discount_pct=spp_discount_pct,
                    cost=cost,
                    category=category,
                    model=model,
                    volume_liters=volume_liters,
                    weight_kg=weight_kg,
                    storage_days=storage_days,
                    category_rates=category_rates,
                    tax_regime=tax_regime,
                    adv_rate=adv_rate,
                    boost_rate=boost_rate,
                    acquiring_rate=acquiring_rate,
                    defect_base_rate=defect_base_rate,
                    buyout_rate=buyout_rate,
                    cancellation_rate=cancellation_rate,
                    other_fixed_rub=other_fixed_rub,
                    logistics_params=logistics_params,
                    promo_discount_from_regular_pct=promo_discount_from_regular,
                )
                price_regular = rec_regular
                price_promo = rec_promo

            current_metrics = calc_price_metrics(
                base_regular_price=price_regular,
                promo_price=price_promo,
                spp_discount_pct=spp_discount_pct,
                cost=cost,
                category=category,
                model=model,
                volume_liters=volume_liters,
                weight_kg=weight_kg,
                storage_days=storage_days,
                category_rates=category_rates,
                tax_regime=tax_regime,
                adv_rate=adv_rate,
                boost_rate=boost_rate,
                acquiring_rate=acquiring_rate,
                defect_base_rate=defect_base_rate,
                buyout_rate=buyout_rate,
                cancellation_rate=cancellation_rate,
                other_fixed_rub=other_fixed_rub,
                logistics_params=logistics_params,
            )

            recommended_regular_price, recommended_promo_price, recommended_metrics = find_recommended_price(
                target_margin_pct=target_margin,
                regular_price_reference=price_regular,
                spp_discount_pct=spp_discount_pct,
                cost=cost,
                category=category,
                model=model,
                volume_liters=volume_liters,
                weight_kg=weight_kg,
                storage_days=storage_days,
                category_rates=category_rates,
                tax_regime=tax_regime,
                adv_rate=adv_rate,
                boost_rate=boost_rate,
                acquiring_rate=acquiring_rate,
                defect_base_rate=defect_base_rate,
                buyout_rate=buyout_rate,
                cancellation_rate=cancellation_rate,
                other_fixed_rub=other_fixed_rub,
                logistics_params=logistics_params,
                promo_discount_from_regular_pct=promo_discount_from_regular,
            )

            status = classify_sku_status(
                current_metrics["margin_pct"],
                current_metrics["profit_after_tax_rub"]
            )

            results.append({
                "SKU": sku,
                "Название": name,
                "Модель": model,
                "Категория Ozon": category,
                "Источник категории": category_source,
                "Статус SKU": status,

                "Длина, см": safe_round(length_cm, 2),
                "Ширина, см": safe_round(width_cm, 2),
                "Высота, см": safe_round(height_cm, 2),
                "Вес, кг": safe_round(weight_kg, 3),
                "Объём, л": safe_round(volume_liters, 3),
                "Себестоимость, ₽": safe_round(cost, 2),

                "Цена без акции, ₽": current_metrics["regular_price_rub"],
                "Цена акции, ₽": current_metrics["promo_price_rub"],
                "СПП / скидка маркетплейса, %": safe_round(spp_discount, 2),
                "Цена покупателя, ₽": current_metrics["customer_price_rub"],
                "Выручка продавца после СПП, ₽": current_metrics["seller_revenue_price_rub"],
                "Скидка маркетплейса, ₽": current_metrics["marketplace_discount_rub"],

                "Комиссия от текущей цены, %": current_metrics["commission_rate_pct"],
                "Комиссия от текущей цены, ₽": current_metrics["commission_rub"],

                "Логистика прямая, ₽": current_metrics["direct_logistics_rub"],
                "в т.ч. обработка, ₽": current_metrics["processing_rub"],
                "в т.ч. доставка, ₽": current_metrics["delivery_rub"],
                "Хранение, ₽": current_metrics["storage_rub"],

                "Отмены, %": current_metrics["cancellation_share_pct"],
                "Эффективная выкупаемость, %": current_metrics["effective_buyout_share_pct"],
                "Невыкупы после доставки, %": current_metrics["non_buyout_after_dispatch_share_pct"],

                "Возвраты / отмены всего, ₽": current_metrics["returns_total_rub"],
                "Обратная логистика возвратов, ₽": current_metrics["return_logistics_rub"],
                "Логистика отмен, ₽": current_metrics["cancellation_logistics_rub"],
                "Обработка возвратов / отмен, ₽": current_metrics["return_processing_rub"],
                "Потеря товарного вида на возвратах, ₽": current_metrics["damage_reserve_on_returns_rub"],

                "Эквайринг, %": safe_round(acquiring, 2),
                "Эквайринг от текущей цены, ₽": current_metrics["acquiring_rub"],
                "Реклама, %": safe_round(ad, 2),
                "Реклама от текущей цены, ₽": current_metrics["advertising_rub"],
                "Буст, %": safe_round(boost, 2),
                "Буст от текущей цены, ₽": current_metrics["boost_rub"],
                "Базовый резерв на брак, ₽": current_metrics["base_defect_reserve_rub"],
                "Прочие расходы, ₽": current_metrics["other_fixed_rub"],

                "Полная себестоимость от текущей цены, ₽": current_metrics["full_cost_before_tax"],
                "Налог от текущей цены, ₽": current_metrics["tax_rub"],
                "Прибыль от текущей цены, ₽": current_metrics["profit_after_tax_rub"],
                "Прибыль от текущей цены, % от цены": current_metrics["profit_pct_of_revenue"],
                "Маржинальность от текущей цены, %": current_metrics["margin_pct"],

                "Рекомендованная цена без акции, ₽": safe_round(recommended_regular_price, 2),
                "Рекомендованная цена акции, ₽": safe_round(recommended_promo_price, 2),
                "Выручка продавца от рекомендованной цены, ₽": recommended_metrics["seller_revenue_price_rub"],
                "Комиссия от рекомендованной цены, %": recommended_metrics["commission_rate_pct"],
                "Комиссия от рекомендованной цены, ₽": recommended_metrics["commission_rub"],
                "Полная себестоимость от рекомендованной цены, ₽": recommended_metrics["full_cost_before_tax"],
                "Налог от рекомендованной цены, ₽": recommended_metrics["tax_rub"],
                "Прибыль от рекомендованной цены, ₽": recommended_metrics["profit_after_tax_rub"],
                "Прибыль от рекомендованной цены, % от цены": recommended_metrics["profit_pct_of_revenue"],
                "Маржинальность от рекомендованной цены, %": recommended_metrics["margin_pct"],
                "Наценка к себестоимости от рекомендованной цены, %": recommended_metrics["markup_to_cost_pct"],
            })

        res_df = pd.DataFrame(results)

        # =========================
        # KPI Dashboard
        # =========================
        st.markdown("## KPI дашборд")

        total_sku = len(res_df)
        bad_sku = int((res_df["Статус SKU"] == "Критично").sum())
        risk_sku = int((res_df["Статус SKU"] == "Риск").sum())
        avg_current_margin = safe_round(res_df["Маржинальность от текущей цены, %"].mean(), 2)
        avg_recommended_margin = safe_round(res_df["Маржинальность от рекомендованной цены, %"].mean(), 2)
        avg_current_profit = safe_round(res_df["Прибыль от текущей цены, ₽"].mean(), 2)
        avg_recommended_price = safe_round(res_df["Рекомендованная цена акции, ₽"].mean(), 2)

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.metric("SKU в расчёте", total_sku)
        with k2:
            st.metric("Критично", bad_sku)
        with k3:
            st.metric("Риск", risk_sku)
        with k4:
            st.metric("Средняя маржинальность текущая, %", avg_current_margin)

        k5, k6, k7 = st.columns(3)
        with k5:
            st.metric("Средняя маржинальность рекомендованная, %", avg_recommended_margin)
        with k6:
            st.metric("Средняя прибыль текущая, ₽", avg_current_profit)
        with k7:
            st.metric("Средняя рекомендованная цена акции, ₽", avg_recommended_price)

        st.markdown("## Результат расчёта")

        styled_df = res_df.style.apply(highlight_status, axis=1)
        st.dataframe(styled_df, use_container_width=True)

        st.markdown("## Выгрузка результата")

        csv_data = res_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Скачать результат (CSV)",
            data=csv_data,
            file_name="ozon_unit_economics_results.csv",
            mime="text/csv"
        )

        export_rates_df = pd.DataFrame(
            [{"Категория Ozon": k, "Комиссия, %": v} for k, v in category_rates.items()]
        )

        xlsx_data = to_excel_bytes({
            "unit_economics": res_df,
            "category_rates_used": export_rates_df,
            "instructions": build_instructions_template(),
        })
        st.download_button(
            "Скачать результат (Excel)",
            data=xlsx_data,
            file_name="ozon_unit_economics_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
