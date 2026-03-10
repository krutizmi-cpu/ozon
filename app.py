import json
import math
import sqlite3
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Ozon — Юнит-экономика",
    layout="wide",
    page_icon="📦"
)

DB_PATH = "products_storage.db"
DATA_DIR = Path("data")
COMMISSIONS_PATH = DATA_DIR / "ozon_commissions.xlsx"
LOGISTICS_CONFIG_PATH = DATA_DIR / "ozon_logistics_config.json"
OZON_API_BASE = "https://api-seller.ozon.ru"


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
            length_cm REAL,
            width_cm REAL,
            height_cm REAL,
            weight_kg REAL,
            cost REAL DEFAULT 0,
            current_price REAL DEFAULT 0
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


def safe_round(value, ndigits=2):
    try:
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


def load_json(path: Path, default_value: dict):
    if not path.exists():
        return default_value
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default_value


def save_json(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================
# Templates / files
# =========================
def build_catalog_template():
    return pd.DataFrame([
        {
            "Артикул, SKU": "SKU-001",
            "Название товара": "Дрель ударная электрическая 650 Вт",
            "Длина, см": 32,
            "Ширина, см": 8,
            "Высота, см": 26,
            "Вес, кг": 2.2,
            "Себестоимость, ₽": 2800,
            "Текущая цена, ₽": 4290,
        },
        {
            "Артикул, SKU": "SKU-002",
            "Название товара": "Электровелосипед городской взрослый складной",
            "Длина, см": 155,
            "Ширина, см": 28,
            "Высота, см": 80,
            "Вес, кг": 32,
            "Себестоимость, ₽": 30167,
            "Текущая цена, ₽": 68000,
        }
    ])


def build_template_notes():
    return pd.DataFrame([
        {"Поле": "Артикул, SKU", "Описание": "Артикул продавца. По нему система пытается найти товар в Ozon.", "Пример": "SKU-001"},
        {"Поле": "Название товара", "Описание": "Название товара. Используется как резерв для определения категории, если Ozon API не вернул категорию.", "Пример": "Дрель ударная электрическая 650 Вт"},
        {"Поле": "Длина, см", "Описание": "Длина товара в сантиметрах.", "Пример": "32"},
        {"Поле": "Ширина, см", "Описание": "Ширина товара в сантиметрах.", "Пример": "8"},
        {"Поле": "Высота, см", "Описание": "Высота товара в сантиметрах.", "Пример": "26"},
        {"Поле": "Вес, кг", "Описание": "Вес товара в килограммах.", "Пример": "2.2"},
        {"Поле": "Себестоимость, ₽", "Описание": "Полная себестоимость одной единицы товара.", "Пример": "2800"},
        {"Поле": "Текущая цена, ₽", "Описание": "Текущая цена продажи, от которой считается текущая юнит-экономика.", "Пример": "4290"},
    ])


def build_default_commissions_df():
    return pd.DataFrame([
        {"category_id": 0, "Категория Ozon": "Электровелосипеды", "Вознаграждение Ozon, %": 35},
        {"category_id": 0, "Категория Ozon": "Электроинструменты", "Вознаграждение Ozon, %": 15},
        {"category_id": 0, "Категория Ozon": "Велосипеды", "Вознаграждение Ozon, %": 15},
        {"category_id": 0, "Категория Ozon": "Спорт и отдых", "Вознаграждение Ozon, %": 15},
        {"category_id": 0, "Категория Ozon": "Смартфоны", "Вознаграждение Ozon, %": 10},
        {"category_id": 0, "Категория Ozon": "Электроника", "Вознаграждение Ozon, %": 15},
        {"category_id": 0, "Категория Ozon": "Компьютеры и комплектующие", "Вознаграждение Ozon, %": 12},
        {"category_id": 0, "Категория Ozon": "Одежда и обувь", "Вознаграждение Ozon, %": 22},
        {"category_id": 0, "Категория Ozon": "Красота и здоровье", "Вознаграждение Ozon, %": 12},
        {"category_id": 0, "Категория Ozon": "Дом и сад", "Вознаграждение Ozon, %": 16},
        {"category_id": 0, "Категория Ozon": "Детские товары", "Вознаграждение Ozon, %": 18},
        {"category_id": 0, "Категория Ozon": "Автотовары", "Вознаграждение Ozon, %": 15},
        {"category_id": 0, "Категория Ozon": "Канцтовары", "Вознаграждение Ozon, %": 17},
        {"category_id": 0, "Категория Ozon": "Прочее", "Вознаграждение Ozon, %": 20},
    ])


def ensure_data_files():
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if not COMMISSIONS_PATH.exists():
        with pd.ExcelWriter(COMMISSIONS_PATH, engine="openpyxl") as writer:
            build_default_commissions_df().to_excel(writer, index=False, sheet_name="commissions")

    if not LOGISTICS_CONFIG_PATH.exists():
        save_json(LOGISTICS_CONFIG_PATH, {
            "fbo": {
                "processing_rub": 20.0,
                "base_logistics_rub": 83.0,
                "extra_kg_rub": 8.0,
                "extra_liter_rub": 8.0,
                "last_mile_percent": 5.5,
                "last_mile_min_rub": 20.0,
                "last_mile_max_rub": 500.0
            },
            "fbs": {
                "processing_rub_min_rub": 10.0,
                "processing_rub_max_rub": 30.0,
                "base_logistics_rub": 83.0,
                "extra_kg_rub": 8.0,
                "extra_liter_rub": 8.0,
                "last_mile_rub": 25.0
            },
            "common": {
                "included_weight_kg": 1.0,
                "included_volume_l": 5.0,
                "storage_grace_days": 14,
                "storage_rub_per_liter_day": 0.25,
                "return_logistics_coef": 1.0,
                "return_processing_rub": 15.0,
                "defect_on_return_rate": 0.05
            }
        })


def load_commissions_df():
    ensure_data_files()
    try:
        df = pd.read_excel(COMMISSIONS_PATH)
    except Exception:
        df = build_default_commissions_df()

    for col in ["category_id", "Категория Ozon", "Вознаграждение Ozon, %"]:
        if col not in df.columns:
            df[col] = None

    df["category_id"] = df["category_id"].fillna(0)
    df["Категория Ozon"] = df["Категория Ozon"].fillna("").astype(str)
    df["Вознаграждение Ozon, %"] = df["Вознаграждение Ozon, %"].apply(lambda x: clean_num(x, 0.0))
    return df


# =========================
# Ozon API
# =========================
def get_ozon_credentials():
    client_id = st.secrets.get("OZON_CLIENT_ID", "")
    api_key = st.secrets.get("OZON_API_KEY", "")
    return str(client_id).strip(), str(api_key).strip()


def has_ozon_credentials():
    client_id, api_key = get_ozon_credentials()
    return bool(client_id and api_key)


def ozon_headers():
    client_id, api_key = get_ozon_credentials()
    return {
        "Client-Id": client_id,
        "Api-Key": api_key,
        "Content-Type": "application/json"
    }


def ozon_post(path: str, payload: dict, timeout=30):
    url = f"{OZON_API_BASE}{path}"
    resp = requests.post(url, headers=ozon_headers(), json=payload, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


def fetch_ozon_products_by_offer_ids(offer_ids):
    result = {}
    offer_ids = [str(x).strip() for x in offer_ids if str(x).strip()]
    if not offer_ids or not has_ozon_credentials():
        return result

    product_map = {}

    try:
        list_payload = {"filter": {"offer_id": offer_ids}, "limit": len(offer_ids)}
        list_resp = ozon_post("/v3/product/list", list_payload)
        items = list_resp.get("result", {}).get("items", []) if isinstance(list_resp, dict) else []

        for item in items:
            offer_id = str(item.get("offer_id", "")).strip()
            product_id = item.get("product_id")
            if offer_id:
                product_map[offer_id] = {"offer_id": offer_id, "product_id": product_id}
    except Exception:
        return result

    if product_map:
        try:
            info_payload = {"product_id": [x["product_id"] for x in product_map.values() if x.get("product_id")]}
            info_resp = ozon_post("/v2/product/info/list", info_payload)
            items = info_resp.get("result", {}).get("items", []) if isinstance(info_resp, dict) else []

            by_product_id = {}
            for item in items:
                by_product_id[item.get("id")] = item

            for offer_id, meta in product_map.items():
                product_id = meta.get("product_id")
                info = by_product_id.get(product_id, {})
                result[offer_id] = {
                    "offer_id": offer_id,
                    "product_id": product_id,
                    "sku_ozon": info.get("sku") or info.get("fbo_sku") or info.get("fbs_sku"),
                    "category_id": info.get("description_category_id") or info.get("category_id") or info.get("type_id"),
                    "category_name": info.get("category_name") or info.get("description_category_name") or "",
                    "source": "Ozon API"
                }
        except Exception:
            for offer_id, meta in product_map.items():
                result[offer_id] = {
                    "offer_id": offer_id,
                    "product_id": meta.get("product_id"),
                    "sku_ozon": None,
                    "category_id": None,
                    "category_name": "",
                    "source": "Ozon API partial"
                }

    return result


# =========================
# Category normalization / fallback
# =========================
CATEGORY_KEYWORDS = {
    "Электровелосипеды": [
        "электровелосипед", "e-bike", "ebike", "электро велосипед"
    ],
    "Электроинструменты": [
        "дрель", "перфоратор", "шуруповерт", "шуруповёрт", "болгарка", "лобзик",
        "шлифмашина", "гайковерт", "гайковёрт", "электроинструмент"
    ],
    "Велосипеды": [
        "горный велосипед", "велосипед", "bmx", "двухколесный велосипед", "двухколёсный велосипед"
    ],
    "Спорт и отдых": [
        "самокат", "тренажер", "тренажёр", "беговая дорожка", "эллипс", "гантел", "коврик", "фитнес", "рюкзак"
    ],
    "Смартфоны": ["смартфон", "iphone", "xiaomi", "redmi", "realme", "galaxy"],
    "Электроника": ["наушник", "гарнитур", "колонка", "power bank", "пауэрбанк", "кабель", "зарядк", "bluetooth"],
    "Компьютеры и комплектующие": ["ноутбук", "монитор", "ssd", "hdd", "клавиатур", "мышь", "видеокарта", "процессор"],
    "Одежда и обувь": ["футболка", "худи", "джинсы", "куртка", "кроссовки", "ботинки", "рубашка"],
    "Красота и здоровье": ["крем", "шампунь", "сыворотка", "маска", "парфюм", "духи"],
    "Дом и сад": ["контейнер", "посуда", "сковород", "подушка", "одеяло", "лампа", "стеллаж"],
    "Детские товары": ["детск", "игрушк", "коляска", "подгуз", "конструктор"],
    "Автотовары": ["авто", "машин", "держатель", "щетка", "чехол на руль"],
    "Канцтовары": ["тетрад", "ручка", "карандаш", "маркер", "ежедневник"],
}


def normalize_category_name(raw_name: str):
    text = str(raw_name).strip().lower()

    if not text:
        return ""

    if "электровелосип" in text:
        return "Электровелосипеды"
    if "велосип" in text:
        return "Велосипеды"
    if any(x in text for x in ["дрель", "шурупов", "перфоратор", "болгарка", "лобзик", "гайков"]):
        return "Электроинструменты"
    if any(x in text for x in ["беговая дорожка", "тренаж", "самокат", "фитнес"]):
        return "Спорт и отдых"

    return raw_name.strip()


def get_keyword_category(name: str):
    text = str(name).lower().strip()
    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw in text:
                return category
    return None


def fallback_category(name):
    by_kw = get_keyword_category(name)
    if by_kw:
        return by_kw, "Словарь"
    return "Прочее", "Fallback"


# =========================
# Commissions
# =========================
def get_reward_percent(price_for_commission, category_id, category_name, commissions_df):
    if price_for_commission <= 100:
        return 14.0
    if price_for_commission <= 300:
        return 20.0

    normalized_name = normalize_category_name(category_name)

    if category_id not in (None, "", 0):
        matched = commissions_df[commissions_df["category_id"].astype(str) == str(category_id)]
        if not matched.empty:
            return clean_num(matched.iloc[0]["Вознаграждение Ozon, %"], 20.0)

    if str(normalized_name).strip():
        matched = commissions_df[
            commissions_df["Категория Ozon"].str.strip().str.lower() == str(normalized_name).strip().lower()
        ]
        if not matched.empty:
            return clean_num(matched.iloc[0]["Вознаграждение Ozon, %"], 20.0)

    matched = commissions_df[
        commissions_df["Категория Ozon"].str.strip().str.lower() == "прочее"
    ]
    if not matched.empty:
        return clean_num(matched.iloc[0]["Вознаграждение Ozon, %"], 20.0)

    return 20.0


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
    tax = max(revenue, 0) * rate if mode == "revenue" else max(profit_before_tax, 0) * rate
    profit_after_tax = profit_before_tax - tax
    profit_pct_of_revenue = (profit_after_tax / revenue * 100) if revenue > 0 else 0

    return (
        safe_round(tax, 2),
        safe_round(profit_before_tax, 2),
        safe_round(profit_after_tax, 2),
        safe_round(profit_pct_of_revenue, 2)
    )


# =========================
# Logistics / returns
# =========================
def calc_shipping_blocks(
    model: str,
    price: float,
    volume_liters: float,
    weight_kg: float,
    storage_days: int,
    logistics_cfg: dict,
):
    common = logistics_cfg["common"]
    overweight = max(0.0, weight_kg - common["included_weight_kg"])
    overvolume = max(0.0, volume_liters - common["included_volume_l"])

    if model == "FBO":
        cfg = logistics_cfg["fbo"]
        processing_rub = cfg["processing_rub"]
        logistics_rub = cfg["base_logistics_rub"] + overweight * cfg["extra_kg_rub"] + overvolume * cfg["extra_liter_rub"]
        last_mile_rub = price * (cfg["last_mile_percent"] / 100.0)
        last_mile_rub = max(cfg["last_mile_min_rub"], min(last_mile_rub, cfg["last_mile_max_rub"]))
    else:
        cfg = logistics_cfg["fbs"]
        processing_rub = cfg["processing_rub_min_rub"] if price <= 5000 else cfg["processing_rub_max_rub"]
        logistics_rub = cfg["base_logistics_rub"] + overweight * cfg["extra_kg_rub"] + overvolume * cfg["extra_liter_rub"]
        last_mile_rub = cfg["last_mile_rub"]

    storage_rub = 0.0
    if model == "FBO":
        paid_days = max(0, storage_days - common["storage_grace_days"])
        storage_rub = paid_days * volume_liters * common["storage_rub_per_liter_day"]

    return {
        "Обработка отправления, ₽": safe_round(processing_rub, 2),
        "Логистика, ₽": safe_round(logistics_rub, 2),
        "Последняя миля, ₽": safe_round(last_mile_rub, 2),
        "Хранение, ₽": safe_round(storage_rub, 2),
    }


def calc_returns_block(
    processing_rub: float,
    logistics_rub: float,
    last_mile_rub: float,
    cost: float,
    buyout_rate: float,
    logistics_cfg: dict,
):
    common = logistics_cfg["common"]

    buyout_share = max(0.0, min(1.0, buyout_rate))
    non_buyout_share = max(0.0, 1.0 - buyout_share)

    return_or_cancel_rub = (logistics_rub + last_mile_rub) * non_buyout_share
    reverse_logistics_rub = logistics_rub * non_buyout_share * common["return_logistics_coef"]
    return_processing_rub = processing_rub * non_buyout_share + common["return_processing_rub"] * non_buyout_share
    defect_on_return_rub = cost * non_buyout_share * common["defect_on_return_rate"]

    return {
        "Выкупаемость, %": safe_round(buyout_share * 100, 2),
        "Невыкуп, %": safe_round(non_buyout_share * 100, 2),
        "Возврат / отмена, ₽": safe_round(return_or_cancel_rub, 2),
        "Обратная логистика, ₽": safe_round(reverse_logistics_rub, 2),
        "Обработка возврата, ₽": safe_round(return_processing_rub, 2),
        "Потери на возврате / брак, ₽": safe_round(defect_on_return_rub, 2),
    }


# =========================
# Unit economics
# =========================
def calc_price_metrics(
    price: float,
    cost: float,
    reward_percent: float,
    model: str,
    volume_liters: float,
    weight_kg: float,
    storage_days: int,
    spp_discount_pct: float,
    acquiring_rate: float,
    ad_rate: float,
    defect_base_rate: float,
    buyout_rate: float,
    tax_regime: str,
    logistics_cfg: dict,
):
    seller_revenue = price * (1 - max(0.0, min(1.0, spp_discount_pct)))

    reward_rub = seller_revenue * (reward_percent / 100.0)
    acquiring_rub = seller_revenue * acquiring_rate
    ad_rub = seller_revenue * ad_rate
    base_defect_rub = cost * defect_base_rate

    ship = calc_shipping_blocks(
        model=model,
        price=price,
        volume_liters=volume_liters,
        weight_kg=weight_kg,
        storage_days=storage_days,
        logistics_cfg=logistics_cfg,
    )

    returns = calc_returns_block(
        processing_rub=ship["Обработка отправления, ₽"],
        logistics_rub=ship["Логистика, ₽"],
        last_mile_rub=ship["Последняя миля, ₽"],
        cost=cost,
        buyout_rate=buyout_rate,
        logistics_cfg=logistics_cfg,
    )

    total_cost_before_tax = (
        cost
        + reward_rub
        + acquiring_rub
        + ship["Обработка отправления, ₽"]
        + ship["Логистика, ₽"]
        + ship["Последняя миля, ₽"]
        + ship["Хранение, ₽"]
        + returns["Возврат / отмена, ₽"]
        + returns["Обратная логистика, ₽"]
        + returns["Обработка возврата, ₽"]
        + returns["Потери на возврате / брак, ₽"]
        + ad_rub
        + base_defect_rub
    )

    tax_rub, profit_before_tax_rub, profit_after_tax_rub, profit_pct = calc_tax(
        revenue=seller_revenue,
        total_cost_before_tax=total_cost_before_tax,
        regime=tax_regime
    )

    margin_pct = ((seller_revenue / total_cost_before_tax - 1) * 100) if total_cost_before_tax > 0 else 0.0
    markup_pct = ((price / cost - 1) * 100) if cost > 0 else 0.0

    return {
        "Цена, ₽": safe_round(price, 2),
        "Выручка продавца, ₽": safe_round(seller_revenue, 2),
        "Вознаграждение Ozon, %": safe_round(reward_percent, 2),
        "Вознаграждение Ozon, ₽": safe_round(reward_rub, 2),
        "Эквайринг, ₽": safe_round(acquiring_rub, 2),
        "Реклама, ₽": safe_round(ad_rub, 2),
        "Базовый брак / списание, ₽": safe_round(base_defect_rub, 2),
        **ship,
        **returns,
        "Полная себестоимость, ₽": safe_round(total_cost_before_tax, 2),
        "Налог, ₽": safe_round(tax_rub, 2),
        "Прибыль, ₽": safe_round(profit_after_tax_rub, 2),
        "Прибыль, %": safe_round(profit_pct, 2),
        "Маржинальность, %": safe_round(margin_pct, 2),
        "Наценка к себестоимости, %": safe_round(markup_pct, 2),
    }


def find_recommended_price(
    target_margin_pct: float,
    current_price_reference: float,
    cost: float,
    reward_percent: float,
    model: str,
    volume_liters: float,
    weight_kg: float,
    storage_days: int,
    spp_discount_pct: float,
    acquiring_rate: float,
    ad_rate: float,
    defect_base_rate: float,
    buyout_rate: float,
    tax_regime: str,
    logistics_cfg: dict,
):
    def get_metrics(candidate_price):
        return calc_price_metrics(
            price=candidate_price,
            cost=cost,
            reward_percent=reward_percent,
            model=model,
            volume_liters=volume_liters,
            weight_kg=weight_kg,
            storage_days=storage_days,
            spp_discount_pct=spp_discount_pct,
            acquiring_rate=acquiring_rate,
            ad_rate=ad_rate,
            defect_base_rate=defect_base_rate,
            buyout_rate=buyout_rate,
            tax_regime=tax_regime,
            logistics_cfg=logistics_cfg,
        )

    low = max(cost * 0.5, 1.0)
    high = max(current_price_reference if current_price_reference > 0 else cost * 3, cost * 10, 1000.0)

    for _ in range(25):
        m = get_metrics(high)
        if m["Маржинальность, %"] >= target_margin_pct:
            break
        high *= 1.5

    for _ in range(60):
        mid = (low + high) / 2
        m = get_metrics(mid)
        if m["Маржинальность, %"] >= target_margin_pct:
            high = mid
        else:
            low = mid

    recommended_price = safe_round(high, 2)
    recommended_metrics = get_metrics(recommended_price)
    return recommended_price, recommended_metrics


def classify_sku_status(margin_pct, profit_rub):
    if profit_rub < 0 or margin_pct < 0:
        return "Критично"
    if margin_pct < 10:
        return "Риск"
    if margin_pct < 20:
        return "Норма"
    return "Хорошо"


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
# App init
# =========================
ensure_data_files()
conn = init_db()
commissions_df = load_commissions_df()
logistics_cfg = load_json(LOGISTICS_CONFIG_PATH, {})

st.title("Ozon — Юнит-экономика")
st.caption("Загрузите один Excel-шаблон, система сама определит категорию, подберёт вознаграждение Ozon и посчитает юнит-экономику")

st.markdown("## 1. Загрузите файл")
st.download_button(
    "Скачать шаблон (Excel)",
    data=to_excel_bytes({
        "Товары": build_catalog_template(),
        "Инструкция": build_template_notes()
    }),
    file_name="ozon_шаблон.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

uploaded_catalog = st.file_uploader(
    "Загрузите Excel-файл с товарами",
    type=["xlsx"],
    key="catalog_upload"
)

st.markdown("## 2. Параметры расчёта")

col_p1, col_p2, col_p3 = st.columns(3)
with col_p1:
    model = st.radio("Модель работы", ["FBO", "FBS"], horizontal=True)
with col_p2:
    tax_regime = st.selectbox(
        "Налогообложение",
        [
            "ОСНО (22% от прибыли)",
            "УСН Доходы (6%)",
            "УСН Доходы-Расходы (15%)",
            "АУСН Доходы (8%)",
            "УСН с НДС 5%",
            "УСН с НДС 7%",
        ]
    )
with col_p3:
    target_margin = st.slider("Целевая маржинальность, %", 0, 100, 20)

col_p4, col_p5, col_p6, col_p7 = st.columns(4)
with col_p4:
    spp_discount = st.slider("СПП, %", 0, 50, 0)
with col_p5:
    ad = st.slider("Реклама, %", 0, 50, 5)
with col_p6:
    buyout = st.slider("Выкупаемость, %", 10, 100, 90)
with col_p7:
    acquiring = st.slider("Эквайринг, %", 0.0, 10.0, 1.0, 0.1)

col_p8, col_p9 = st.columns(2)
with col_p8:
    defect = st.slider("Брак / списание, %", 0, 20, 0)
with col_p9:
    storage_days = st.number_input("Срок хранения, дней", min_value=0, max_value=365, value=45, step=1)

st.markdown("## 3. Расчёт")
calculate = st.button("🚀 Рассчитать", type="primary", use_container_width=True)

catalog_df = pd.DataFrame()

if uploaded_catalog is not None:
    try:
        catalog_df = pd.read_excel(uploaded_catalog)
        catalog_df.columns = [str(c).strip() for c in catalog_df.columns]
        st.success(f"Файл загружен. Строк: {len(catalog_df)}")
        with st.expander("Предпросмотр файла"):
            st.dataframe(catalog_df.head(20), use_container_width=True)
    except Exception as e:
        st.error(f"Не удалось прочитать Excel: {e}")

if calculate:
    if uploaded_catalog is None or catalog_df.empty:
        st.error("Сначала загрузите Excel-файл с товарами.")
    else:
        conn.execute("DELETE FROM products")
        conn.commit()

        for _, row in catalog_df.iterrows():
            sku = str(row.get("Артикул, SKU", row.get("SKU", row.get("Артикул", "")))).strip()
            name = str(row.get("Название товара", row.get("Название", row.get("Наименование", "")))).strip()

            length_cm = clean_num(row.get("Длина, см", row.get("Длина", 0)), 0.0)
            width_cm = clean_num(row.get("Ширина, см", row.get("Ширина", 0)), 0.0)
            height_cm = clean_num(row.get("Высота, см", row.get("Высота", 0)), 0.0)
            weight_kg = clean_num(row.get("Вес, кг", row.get("Вес", 0)), 0.0)

            cost = clean_num(row.get("Себестоимость, ₽", row.get("Себестоимость", 0)), 0.0)
            current_price = clean_num(row.get("Текущая цена, ₽", row.get("Цена", 0)), 0.0)

            if not sku or not name:
                continue

            conn.execute("""
                INSERT OR REPLACE INTO products
                (sku, name, length_cm, width_cm, height_cm, weight_kg, cost, current_price)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (sku, name, length_cm, width_cm, height_cm, weight_kg, cost, current_price))

        conn.commit()
        all_products = pd.read_sql("SELECT * FROM products ORDER BY id DESC", conn)

        buyout_rate = buyout / 100.0
        defect_base_rate = defect / 100.0
        ad_rate = ad / 100.0
        acquiring_rate = acquiring / 100.0
        spp_discount_pct = spp_discount / 100.0

        with st.spinner("Определяем категории и считаем юнит-экономику..."):
            skus = all_products["sku"].astype(str).tolist()
            ozon_map = {}
            api_error = None

            try:
                ozon_map = fetch_ozon_products_by_offer_ids(skus)
            except Exception as e:
                api_error = str(e)

            results = []

            for _, p in all_products.iterrows():
                sku = str(p.get("sku", "")).strip()
                name = str(p.get("name", "")).strip()
                length_cm = clean_num(p.get("length_cm", 0), 0.0)
                width_cm = clean_num(p.get("width_cm", 0), 0.0)
                height_cm = clean_num(p.get("height_cm", 0), 0.0)
                weight_kg = clean_num(p.get("weight_kg", 0), 0.0)
                cost = clean_num(p.get("cost", 0), 0.0)
                current_price = clean_num(p.get("current_price", 0), 0.0)

                if current_price <= 0:
                    current_price = max(cost * 3, 1000)

                volume_liters = (length_cm * width_cm * height_cm) / 1000.0 if length_cm and width_cm and height_cm else 0.0

                ozon_info = ozon_map.get(sku, {})
                found_in_api = "Да" if bool(ozon_info) else "Нет"

                api_category_id = ozon_info.get("category_id")
                api_category_name = normalize_category_name(ozon_info.get("category_name", ""))

                if api_category_id or str(api_category_name).strip():
                    resolved_category_name = api_category_name if str(api_category_name).strip() else "Прочее"
                    category_method = "Ozon API"
                else:
                    resolved_category_name, category_method = fallback_category(name)
                    api_category_name = resolved_category_name

                reward_percent = get_reward_percent(
                    price_for_commission=current_price * (1 - spp_discount_pct),
                    category_id=api_category_id,
                    category_name=api_category_name,
                    commissions_df=commissions_df
                )

                current_metrics = calc_price_metrics(
                    price=current_price,
                    cost=cost,
                    reward_percent=reward_percent,
                    model=model,
                    volume_liters=volume_liters,
                    weight_kg=weight_kg,
                    storage_days=storage_days,
                    spp_discount_pct=spp_discount_pct,
                    acquiring_rate=acquiring_rate,
                    ad_rate=ad_rate,
                    defect_base_rate=defect_base_rate,
                    buyout_rate=buyout_rate,
                    tax_regime=tax_regime,
                    logistics_cfg=logistics_cfg,
                )

                recommended_price, recommended_metrics = find_recommended_price(
                    target_margin_pct=target_margin,
                    current_price_reference=current_price,
                    cost=cost,
                    reward_percent=reward_percent,
                    model=model,
                    volume_liters=volume_liters,
                    weight_kg=weight_kg,
                    storage_days=storage_days,
                    spp_discount_pct=spp_discount_pct,
                    acquiring_rate=acquiring_rate,
                    ad_rate=ad_rate,
                    defect_base_rate=defect_base_rate,
                    buyout_rate=buyout_rate,
                    tax_regime=tax_regime,
                    logistics_cfg=logistics_cfg,
                )

                status = classify_sku_status(
                    current_metrics["Маржинальность, %"],
                    current_metrics["Прибыль, ₽"]
                )

                results.append({
                    "Артикул, SKU": sku,
                    "Название товара": name,
                    "Товар найден по Ozon API": found_in_api,
                    "Категория определена через": category_method,
                    "Категория Ozon": api_category_name if str(api_category_name).strip() else resolved_category_name,
                    "Ozon product_id": ozon_info.get("product_id"),
                    "Ozon sku": ozon_info.get("sku_ozon"),
                    "category_id": api_category_id,

                    "Вознаграждение Ozon, %": current_metrics["Вознаграждение Ozon, %"],
                    "Вознаграждение Ozon, ₽": current_metrics["Вознаграждение Ozon, ₽"],
                    "Эквайринг, ₽": current_metrics["Эквайринг, ₽"],
                    "Обработка отправления, ₽": current_metrics["Обработка отправления, ₽"],
                    "Логистика, ₽": current_metrics["Логистика, ₽"],
                    "Последняя миля, ₽": current_metrics["Последняя миля, ₽"],
                    "Возврат / отмена, ₽": current_metrics["Возврат / отмена, ₽"],
                    "Обратная логистика, ₽": current_metrics["Обратная логистика, ₽"],
                    "Обработка возврата, ₽": current_metrics["Обработка возврата, ₽"],
                    "Потери на возврате / брак, ₽": current_metrics["Потери на возврате / брак, ₽"],
                    "Хранение, ₽": current_metrics["Хранение, ₽"],
                    "Реклама, ₽": current_metrics["Реклама, ₽"],
                    "Налог, ₽": current_metrics["Налог, ₽"],

                    "Себестоимость, ₽": safe_round(cost, 2),
                    "Текущая цена, ₽": current_metrics["Цена, ₽"],
                    "Выручка продавца текущая, ₽": current_metrics["Выручка продавца, ₽"],
                    "Полная себестоимость по текущей цене, ₽": current_metrics["Полная себестоимость, ₽"],
                    "Прибыль от текущей цены, ₽": current_metrics["Прибыль, ₽"],
                    "Маржинальность от текущей цены, %": current_metrics["Маржинальность, %"],

                    "Рекомендованная цена, ₽": safe_round(recommended_price, 2),
                    "Выручка продавца рекомендованная, ₽": recommended_metrics["Выручка продавца, ₽"],
                    "Полная себестоимость по рекомендованной цене, ₽": recommended_metrics["Полная себестоимость, ₽"],
                    "Прибыль от рекомендованной цены, ₽": recommended_metrics["Прибыль, ₽"],
                    "Маржинальность от рекомендованной цены, %": recommended_metrics["Маржинальность, %"],
                    "Наценка к себестоимости от рекомендованной цены, %": recommended_metrics["Наценка к себестоимости, %"],

                    "Вес, кг": safe_round(weight_kg, 3),
                    "Объём, л": safe_round(volume_liters, 3),
                    "Выкупаемость, %": current_metrics["Выкупаемость, %"],
                    "Невыкуп, %": current_metrics["Невыкуп, %"],
                    "Статус SKU": status,
                })

        res_df = pd.DataFrame(results)

        st.markdown("## 4. KPI")
        total_sku = len(res_df)
        bad_sku = int((res_df["Статус SKU"] == "Критично").sum())
        risk_sku = int((res_df["Статус SKU"] == "Риск").sum())
        found_api = int((res_df["Товар найден по Ozon API"] == "Да").sum())
        avg_current_margin = safe_round(res_df["Маржинальность от текущей цены, %"].mean(), 2)
        avg_recommended_margin = safe_round(res_df["Маржинальность от рекомендованной цены, %"].mean(), 2)

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("SKU в расчёте", total_sku)
        k2.metric("Найдено по Ozon API", found_api)
        k3.metric("Критично", bad_sku)
        k4.metric("Риск", risk_sku)
        k5.metric("Средняя маржинальность текущая, %", avg_current_margin)
        st.metric("Средняя маржинальность рекомендованная, %", avg_recommended_margin)

        st.markdown("## 5. Результат")
        visible_cols = [
            "Артикул, SKU",
            "Название товара",
            "Товар найден по Ozon API",
            "Категория определена через",
            "Категория Ozon",
            "Вознаграждение Ozon, %",
            "Логистика, ₽",
            "Последняя миля, ₽",
            "Эквайринг, ₽",
            "Полная себестоимость по текущей цене, ₽",
            "Текущая цена, ₽",
            "Прибыль от текущей цены, ₽",
            "Маржинальность от текущей цены, %",
            "Рекомендованная цена, ₽",
            "Прибыль от рекомендованной цены, ₽",
            "Маржинальность от рекомендованной цены, %",
            "Статус SKU",
        ]
        shown_df = res_df[visible_cols].copy()
        st.dataframe(shown_df.style.apply(highlight_status, axis=1), use_container_width=True)

        st.markdown("## 6. Выгрузка")
        result_meta = pd.DataFrame([
            {"Параметр": "Модель работы", "Значение": model},
            {"Параметр": "Налогообложение", "Значение": tax_regime},
            {"Параметр": "Целевая маржинальность, %", "Значение": target_margin},
            {"Параметр": "СПП, %", "Значение": spp_discount},
            {"Параметр": "Реклама, %", "Значение": ad},
            {"Параметр": "Выкупаемость, %", "Значение": buyout},
            {"Параметр": "Брак / списание, %", "Значение": defect},
            {"Параметр": "Эквайринг, %", "Значение": acquiring},
            {"Параметр": "Срок хранения, дней", "Значение": storage_days},
            {"Параметр": "Загружено SKU", "Значение": total_sku},
            {"Параметр": "Найдено по Ozon API", "Значение": found_api},
        ])

        st.download_button(
            "Скачать краткий результат (Excel)",
            data=to_excel_bytes({
                "Результат": shown_df,
                "Параметры": result_meta
            }),
            file_name="ozon_краткий_результат.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.download_button(
            "Скачать полный результат (Excel)",
            data=to_excel_bytes({
                "Результат полный": res_df,
                "Параметры": result_meta,
                "Инструкция": build_template_notes()
            }),
            file_name="ozon_полный_результат.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if api_error:
            st.warning("Часть товаров не была найдена через Ozon API. Для них была использована резервная логика.")
