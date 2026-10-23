import streamlit as st
import pandas as pd
import requests
from datetime import date, datetime
from pathlib import Path

st.set_page_config(layout="wide")

st.markdown("""
<style>
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 0rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    @media (max-width: 768px) {
        html, body { overflow-x: hidden !important; }
        [data-testid="stAppViewContainer"] { overflow-x: hidden !important; }
        [data-testid="stMetric"] { margin-bottom: 0px !important; padding: 5px !important; }
        [data-testid="stMetricValue"] { font-size: 1.1rem !important; }
        [data-testid="stMetricLabel"] { font-size: 0.7rem !important; white-space: normal !important; }
        .block-container { padding-top: 1rem !important; gap: 0.5rem !important; }
        [data-testid="column"] { padding: 0px !important; min-width: 0 !important; }
    }
</style>
""", unsafe_allow_html=True)

PURCHASES_FILE = Path("journal.xlsx")
SALES_FILE = Path("journal_sales.xlsx")
DISTRIBUTION_FILE = Path("sales_distribution.xlsx")
TRANSACTIONS_FILE = Path("transactions.xlsx")
INDICATOR_FILE = Path("indicator.xlsx")
ASSETS_FILE = Path("assets_to_imoex.xlsx")
HISTORY_CACHE_FILE = Path("history_cache.xlsx")

PURCHASE_COLUMNS = ["ID", "Дата", "Тикер", "Наименование", "Цена актива", "Кол-во", "Комиссия", "Инвестировано", "Покупка IMOEX", "Вирт. Кол-во IMOEX", "Тек.цена актива", "Тек. цена IMOEX", "% к индексу", "Статус", "Удалено"]
SALE_COLUMNS = ["Sale_ID", "Purchase_IDs", "Дата продажи", "Тикер", "Наименование", "Цена продажи", "Кол-во", "Комиссия", "Сумма продажи", "Продажа IMOEX", "Вирт. Кол-во IMOEX", "Цена покупки актива", "Покупка IMOEX", "% к индексу", "Примечание", "Удалено"]
DISTRIBUTION_COLUMNS = ["Sale_ID", "Purchase_ID", "Qty"]
TRANSACTION_COLUMNS = ["ID", "Parent_ID", "Дата", "Время", "Тип", "Сумма", "Комментарий", "Связано_ID", "Удалено"]
INDICATOR_COLUMNS = ["Наименование", "Тикер", "Вирт. Инвест", "Начало месяца активы", "Начало недели активы", "IMOEX сначала месяца", "Начало недели IMOEX", "Тек.цена актива", "Тек. цена IMOEX", "Месяц % к IMOEX", "Неделя % к IMOEX", "Обновлено_месяц", "Обновлено_неделя"]
ASSETS_COLUMNS = ["Тикер", "Наименование", "Добавлено"]
HISTORY_COLUMNS = ["Тикер", "Месяц", "% Акции", "% IMOEX", "Разница", "Обновлено"]

@st.cache_data(ttl=300)
def get_imoex():
    url = "https://iss.moex.com/iss/engines/stock/markets/index/securities/IMOEX.json"
    try:
        r = requests.get(url, timeout=10)
        data = r.json()
        cols = data["marketdata"]["columns"]
        rows = data["marketdata"]["data"]
        if not rows: return None
        for field in ("CURRENTVALUE", "LASTVALUE", "MARKETPRICE"):
            if field in cols:
                idx = cols.index(field)
                val = rows[0][idx]
                if val is not None: return round(float(val), 3)
        return None
    except Exception: return None

@st.cache_data(ttl=300)
def get_stock_price(ticker):
    if not ticker: return None
    url = f"https://iss.moex.com/iss/engines/stock/markets/shares/securities/{ticker}.json"
    try:
        r = requests.get(url, timeout=10)
        data = r.json()
        cols = data["marketdata"]["columns"]
        rows = data["marketdata"]["data"]
        if not rows: return None
        for field in ("LAST", "MARKETPRICE", "LEGALCLOSEPRICE"):
            if field in cols:
                idx = cols.index(field)
                for row in rows:
                    val = row[idx]
                    if val is not None: return round(float(val), 3)
        return None
    except Exception: return None

def find_tickers(query):
    if not query or len(query) < 2: return []
    url = "https://iss.moex.com/iss/securities.json"
    params = {"q": query, "iss.meta": "off", "iss.only": "securities"}
    try:
        r = requests.get(url, params=params, timeout=10)
        data = r.json()
        cols = data["securities"]["columns"]
        rows = data["securities"]["data"]
        idx_secid = cols.index("secid")
        idx_name = cols.index("name")
        idx_type = cols.index("type")

        results = []
        for row in rows:
            secid, name, typ = row[idx_secid], row[idx_name], row[idx_type]
            # Оставляем только акции
            if typ not in ("common_share", "preferred_share", "exchange_share"): continue
            
            type_label = {"common_share": "обычка", "preferred_share": "преф", "exchange_share": "акция"}.get(typ, typ)
            results.append({"label": f"{secid} — {name} ({type_label})", "ticker": secid, "name": name})
        return results
    except Exception: return []

def load(file, columns):
    if file.exists():
        try:
            df = pd.read_excel(file, engine='openpyxl')
            if df.empty: return pd.DataFrame(columns=columns)
            for col in columns:
                if col not in df.columns: df[col] = None
            df = df[columns]
            for date_col in ("Дата", "Дата продажи"):
                if date_col in df.columns: df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
            if "Удалено" in df.columns: df["Удалено"] = df["Удалено"].fillna(False).astype(bool)
            else: df["Удалено"] = False
            return df
        except Exception: return pd.DataFrame(columns=columns)
    df = pd.DataFrame(columns=columns)
    df["Удалено"] = df["Удалено"].astype(bool)
    return df

def save(df, file): 
    try: df.to_excel(file, index=False, engine='openpyxl')
    except Exception as e: st.error(f"Ошибка сохранения: {e}")

def load_distribution(): return load(DISTRIBUTION_FILE, DISTRIBUTION_COLUMNS)
def save_distribution(df): save(df, DISTRIBUTION_FILE)
def load_transactions(): return load(TRANSACTIONS_FILE, TRANSACTION_COLUMNS)
def save_transactions(df): save(df, TRANSACTIONS_FILE)
def load_indicator(): return load(INDICATOR_FILE, INDICATOR_COLUMNS)
def save_indicator(df): save(df, INDICATOR_FILE)
def load_assets(): return load(ASSETS_FILE, ASSETS_COLUMNS)
def save_assets(df): save(df, ASSETS_FILE)
def load_history_cache(): return load(HISTORY_CACHE_FILE, HISTORY_COLUMNS)
def save_history_cache(df): save(df, HISTORY_CACHE_FILE)

def get_next_purchase_id(purchases):
    if purchases.empty or "ID" not in purchases.columns: return 1
    ids = purchases["ID"].dropna()
    if ids.empty: return 1
    try: return int(ids.astype(int).max()) + 1
    except Exception: return 1

def get_next_sale_id(sales):
    if sales.empty or "Sale_ID" not in sales.columns: return 1
    ids = sales["Sale_ID"].dropna()
    if ids.empty: return 1
    try: return int(ids.astype(int).max()) + 1
    except Exception: return 1

def get_next_transaction_id(df):
    if df.empty or "ID" not in df.columns: return 1
    ids = df["ID"].dropna()
    if ids.empty: return 1
    try: return int(ids.astype(int).max()) + 1
    except Exception: return 1

def add_transaction(t_type, amount, comment="", linked_id=None, parent_id=None):
    df = st.session_state.transactions.copy()
    new_id = get_next_transaction_id(df)
    now = datetime.now()
    new_row = {"ID": new_id, "Parent_ID": parent_id, "Дата": now.date(), "Время": now.strftime("%H:%M:%S"), "Тип": t_type, "Сумма": amount, "Комментарий": comment, "Связано_ID": linked_id, "Удалено": False}
    st.session_state.transactions = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    save_transactions(st.session_state.transactions)
    return new_id

def get_balance():
    df = st.session_state.transactions
    if df.empty: return 0.0
    active = df[df["Удалено"] == False]
    try: return float(active["Сумма"].sum())
    except Exception: return 0.0

def create_purchase_transaction(purchase_id, invested, name="", parent_id=None):
    return add_transaction("Покупка", -abs(float(invested)), f"Покупка {name} (ID {purchase_id})" if name else f"Покупка ID {purchase_id}", linked_id=purchase_id, parent_id=parent_id)

def create_sale_transaction(sale_id, sale_sum, name="", parent_id=None):
    return add_transaction("Продажа", abs(float(sale_sum)), f"Продажа {name} (Sale_ID {sale_id})" if name else f"Продажа Sale_ID {sale_id}", linked_id=sale_id, parent_id=parent_id)

def mark_transaction_deleted(linked_id):
    if st.session_state.transactions.empty: return
    mask = st.session_state.transactions["Связано_ID"] == linked_id
    if mask.any():
        st.session_state.transactions.loc[mask, "Удалено"] = True
        save_transactions(st.session_state.transactions)

def restore_transaction(linked_id):
    if st.session_state.transactions.empty: return
    mask = st.session_state.transactions["Связано_ID"] == linked_id
    if mask.any():
        st.session_state.transactions.loc[mask, "Удалено"] = False
        save_transactions(st.session_state.transactions)

def _as_int(value):
    if value is None or pd.isna(value): return None
    try: return int(value)
    except (ValueError, TypeError): return None

def _parse_purchase_ids(raw):
    if raw is None or pd.isna(raw) or not str(raw).strip(): return set()
    result = set()
    for part in str(raw).split(","):
        part = part.strip()
        if not part: continue
        try: result.add(int(part))
        except (ValueError, TypeError): continue
    return result

def get_purchases_by_parent(parent_tx_id):
    tx = st.session_state.transactions
    parent_tx_id = _as_int(parent_tx_id)
    if parent_tx_id is None or tx.empty: return []
    mask = (tx["Тип"] == "Покупка") & (tx["Parent_ID"].apply(_as_int) == parent_tx_id) & (tx["Удалено"] == False)
    ids = tx.loc[mask, "Связано_ID"].dropna()
    return [int(x) for x in ids if _as_int(x) is not None]

def get_sales_by_parent(parent_tx_id):
    tx = st.session_state.transactions
    parent_tx_id = _as_int(parent_tx_id)
    if parent_tx_id is None or tx.empty: return []
    mask = (tx["Тип"] == "Продажа") & (tx["Parent_ID"].apply(_as_int) == parent_tx_id) & (tx["Удалено"] == False)
    ids = tx.loc[mask, "Связано_ID"].dropna()
    return [int(x) for x in ids if _as_int(x) is not None]

def get_parent_id_for_purchase(purchase_id):
    tx = st.session_state.transactions
    if tx.empty: return None
    purchase_id = _as_int(purchase_id)
    if purchase_id is None: return None
    mask = (tx["Тип"] == "Покупка") & (tx["Связано_ID"].apply(_as_int) == purchase_id) & (tx["Удалено"] == False)
    match = tx.loc[mask]
    if match.empty: return None
    return _as_int(match.iloc[0]["Parent_ID"])

def cascade_delete_deposit(deposit_tx_id):
    tx = st.session_state.transactions
    deposit_tx_id = _as_int(deposit_tx_id)
    stats = {"deposit": 0, "transactions": 0, "purchases": 0, "sales": 0}
    if deposit_tx_id is None or tx.empty: return stats
    dep_mask = (tx["ID"].apply(_as_int) == deposit_tx_id) & (tx["Удалено"] == False)
    if dep_mask.any():
        tx.loc[dep_mask, "Удалено"] = True
        stats["deposit"] = int(dep_mask.sum())
    child_mask = (tx["Parent_ID"].apply(_as_int) == deposit_tx_id) & (tx["Удалено"] == False)
    child_rows = tx.loc[child_mask]
    purchase_ids, sale_ids = set(), set()
    for _, child in child_rows.iterrows():
        if child["Тип"] == "Покупка":
            pid = _as_int(child["Связано_ID"])
            if pid is not None: purchase_ids.add(pid)
        elif child["Тип"] == "Продажа":
            sid = _as_int(child["Связано_ID"])
            if sid is not None: sale_ids.add(sid)
    if child_mask.any():
        tx.loc[child_mask, "Удалено"] = True
        stats["transactions"] = int(child_mask.sum())
    if purchase_ids and not st.session_state.sales.empty:
        for _, s in st.session_state.sales.iterrows():
            if s.get("Удалено", False): continue
            pids = _parse_purchase_ids(s.get("Purchase_IDs", ""))
            if pids & purchase_ids:
                sid = _as_int(s.get("Sale_ID"))
                if sid is not None: sale_ids.add(sid)
    if purchase_ids and not st.session_state.purchases.empty:
        p_mask = st.session_state.purchases["ID"].apply(_as_int).isin(purchase_ids)
        if p_mask.any():
            st.session_state.purchases.loc[p_mask, "Удалено"] = True
            stats["purchases"] = int(p_mask.sum())
    if sale_ids and not st.session_state.sales.empty:
        s_mask = st.session_state.sales["Sale_ID"].apply(_as_int).isin(sale_ids)
        if s_mask.any():
            st.session_state.sales.loc[s_mask, "Удалено"] = True
            stats["sales"] = int(s_mask.sum())
    st.session_state.transactions = tx
    return stats

def cascade_delete_purchase(purchase_id):
    purchase_id = _as_int(purchase_id)
    stats = {"purchases": 0, "sales": 0, "transactions": 0}
    if purchase_id is None: return stats
    if not st.session_state.purchases.empty:
        p_mask = (st.session_state.purchases["ID"].apply(_as_int) == purchase_id) & (st.session_state.purchases["Удалено"] == False)
        if p_mask.any():
            st.session_state.purchases.loc[p_mask, "Удалено"] = True
            stats["purchases"] = int(p_mask.sum())
    sale_ids = set()
    if not st.session_state.sales.empty:
        for _, s in st.session_state.sales.iterrows():
            if s.get("Удалено", False): continue
            pids = _parse_purchase_ids(s.get("Purchase_IDs", ""))
            if purchase_id in pids:
                sid = _as_int(s.get("Sale_ID"))
                if sid is not None: sale_ids.add(sid)
    if sale_ids:
        s_mask = st.session_state.sales["Sale_ID"].apply(_as_int).isin(sale_ids) & (st.session_state.sales["Удалено"] == False)
        if s_mask.any():
            st.session_state.sales.loc[s_mask, "Удалено"] = True
            stats["sales"] = int(s_mask.sum())
    tx = st.session_state.transactions
    if not tx.empty:
        tx_mask = (tx["Тип"] == "Покупка") & (tx["Связано_ID"].apply(_as_int) == purchase_id) & (tx["Удалено"] == False)
        if sale_ids:
            tx_mask |= (tx["Тип"] == "Продажа") & (tx["Связано_ID"].apply(_as_int).isin(sale_ids)) & (tx["Удалено"] == False)
        if tx_mask.any():
            tx.loc[tx_mask, "Удалено"] = True
            stats["transactions"] = int(tx_mask.sum())
        st.session_state.transactions = tx
    return stats

def restore_deposit_cascade(deposit_tx_id):
    tx = st.session_state.transactions
    deposit_tx_id = _as_int(deposit_tx_id)
    stats = {"deposit": 0, "transactions": 0, "purchases": 0, "sales": 0}
    if deposit_tx_id is None or tx.empty: return stats
    dep_mask = tx["ID"].apply(_as_int) == deposit_tx_id
    if dep_mask.any():
        tx.loc[dep_mask, "Удалено"] = False
        stats["deposit"] = int(dep_mask.sum())
    child_mask = tx["Parent_ID"].apply(_as_int) == deposit_tx_id
    child_rows = tx.loc[child_mask]
    purchase_ids, sale_ids = set(), set()
    for _, child in child_rows.iterrows():
        if child["Тип"] == "Покупка":
            pid = _as_int(child["Связано_ID"])
            if pid is not None: purchase_ids.add(pid)
        elif child["Тип"] == "Продажа":
            sid = _as_int(child["Связано_ID"])
            if sid is not None: sale_ids.add(sid)
    if child_mask.any():
        tx.loc[child_mask, "Удалено"] = False
        stats["transactions"] = int(child_mask.sum())
    if purchase_ids and not st.session_state.purchases.empty:
        p_mask = st.session_state.purchases["ID"].apply(_as_int).isin(purchase_ids)
        st.session_state.purchases.loc[p_mask, "Удалено"] = False
        stats["purchases"] = int(p_mask.sum())
    if sale_ids and not st.session_state.sales.empty:
        s_mask = st.session_state.sales["Sale_ID"].apply(_as_int).isin(sale_ids)
        st.session_state.sales.loc[s_mask, "Удалено"] = False
        stats["sales"] = int(s_mask.sum())
    st.session_state.transactions = tx
    return stats

def find_and_clean_orphan_purchases():
    tx = st.session_state.transactions
    if tx.empty: return 0
    active_deposits = set(tx[(tx["Тип"] == "Пополнение") & (tx["Удалено"] == False)]["ID"].apply(_as_int).dropna().tolist())
    orphan_ids = []
    for _, p in st.session_state.purchases.iterrows():
        if p.get("Удалено", False): continue
        pid = _as_int(p.get("ID"))
        if pid is None: continue
        p_tx = tx[(tx["Тип"] == "Покупка") & (tx["Связано_ID"].apply(_as_int) == pid)]
        if p_tx.empty:
            orphan_ids.append(pid)
            continue
        parent_tx_id = _as_int(p_tx.iloc[0]["Parent_ID"])
        if parent_tx_id not in active_deposits:
            orphan_ids.append(pid)
    deleted_count = 0
    for pid in orphan_ids:
        cascade_delete_purchase(pid)
        deleted_count += 1
    if deleted_count > 0:
        save(st.session_state.purchases, PURCHASES_FILE); save(st.session_state.sales, SALES_FILE); save_transactions(st.session_state.transactions)
    return deleted_count

def recalc_status(purchases, sales, distribution):
    purchases = purchases.copy()
    active_sale_ids = set(sales[sales["Удалено"] == False]["Sale_ID"].dropna().astype(int).tolist())
    if not distribution.empty:
        active_dist = distribution[distribution["Sale_ID"].astype(float).isin(active_sale_ids)].copy()
    else:
        active_dist = pd.DataFrame(columns=DISTRIBUTION_COLUMNS)
    sold_list = []
    for _, row in purchases.iterrows():
        purchase_id = row.get("ID")
        sold = 0
        if pd.notna(purchase_id):
            try:
                pid = int(purchase_id)
                matching = active_dist[active_dist["Purchase_ID"].astype(float) == pid]
                for _, d in matching.iterrows():
                    try: sold += abs(int(d["Qty"]))
                    except Exception: pass
            except Exception: pass
        sold_list.append(sold)
    remaining_list, status_list = [], []
    for i, (idx, row) in enumerate(purchases.iterrows()):
        try: total = int(row["Кол-во"])
        except Exception: total = 0
        sold = sold_list[i]
        remaining = max(total - sold, 0)
        if remaining == total: status = "ОТКРЫТА"
        elif remaining == 0: status = "ЗАКРЫТА"
        else: status = "ЧАСТИЧНО ЗАКРЫТА"
        remaining_list.append(remaining)
        status_list.append(status)
    purchases["Продано"] = sold_list
    purchases["Остаток"] = remaining_list
    purchases["Статус"] = status_list
    return purchases

def format_number(val):
    if val is None or pd.isna(val): return ""
    try:
        num = float(val)
        if num == int(num): return str(int(num))
        return f"{num:.3f}".rstrip("0").rstrip(".")
    except (ValueError, TypeError): return str(val)

def refresh_open_purchases(df, current_imoex):
    df = df.copy().reset_index(drop=True)
    df["Тек.цена актива"] = pd.to_numeric(df["Тек.цена актива"], errors="coerce").astype("float64")
    df["Тек. цена IMOEX"] = pd.to_numeric(df["Тек. цена IMOEX"], errors="coerce").astype("float64")
    for idx, row in df.iterrows():
        if row.get("Статус", "") == "ЗАКРЫТА": continue
        ticker = row.get("Тикер")
        if pd.notna(ticker) and ticker:
            price = get_stock_price(ticker)
            if price is not None: df.loc[idx, "Тек.цена актива"] = float(price)
        if current_imoex is not None: df.loc[idx, "Тек. цена IMOEX"] = float(current_imoex)
        try:
            qty = float(row["Кол-во"])
            cur_price, virt_imoex, cur_imoex = df.at[idx, "Тек.цена актива"], row["Вирт. Кол-во IMOEX"], df.at[idx, "Тек. цена IMOEX"]
            if (pd.notna(cur_price) and pd.notna(virt_imoex) and pd.notna(cur_imoex) and virt_imoex != 0 and cur_imoex != 0):
                df.at[idx, "% к индексу"] = round((qty * cur_price) / (virt_imoex * cur_imoex) - 1, 6)
            else: df.at[idx, "% к индексу"] = None
        except Exception: df.at[idx, "% к индексу"] = None
    return df

def calc_total_invested():
    inv_total, sales_total = 0.0, 0.0
    if not st.session_state.purchases.empty:
        try: inv_total = float(st.session_state.purchases[st.session_state.purchases["Удалено"] == False]["Инвестировано"].sum())
        except Exception: pass
    if not st.session_state.sales.empty:
        try: sales_total = float(st.session_state.sales[st.session_state.sales["Удалено"] == False]["Сумма продажи"].sum())
        except Exception: pass
    return inv_total + sales_total

def calc_current_value(visible_purchases, visible_sales):
    if visible_purchases.empty: return 0.0
    recalced = recalc_status(visible_purchases, visible_sales, st.session_state.distribution)
    total = 0.0
    for _, row in recalced.iterrows():
        try:
            remaining, price = float(row["Остаток"]), row.get("Тек.цена актива")
            if pd.notna(price) and remaining > 0: total += remaining * float(price)
        except Exception: pass
    return total

def calc_indicator_pct(row, current_imoex):
    virt = 1
    cur_price = row.get("Тек.цена актива")
    cur_imoex = current_imoex if current_imoex else row.get("Тек. цена IMOEX")
    start_month_price = row.get("Начало месяца активы")
    start_month_imoex = row.get("IMOEX сначала месяца")
    start_week_price = row.get("Начало недели активы")
    start_week_imoex = row.get("Начало недели IMOEX")
    month_pct, week_pct = None, None
    try:
        if (pd.notna(start_month_price) and float(start_month_price) != 0 and pd.notna(cur_price) and pd.notna(start_month_imoex) and float(start_month_imoex) != 0 and pd.notna(cur_imoex)):
            month_pct = (virt / float(start_month_price) * float(cur_price)) - (virt / float(start_month_imoex) * float(cur_imoex))
    except Exception: pass
    try:
        if (pd.notna(start_week_price) and float(start_week_price) != 0 and pd.notna(cur_price) and pd.notna(start_week_imoex) and float(start_week_imoex) != 0 and pd.notna(cur_imoex)):
            week_pct = (virt / float(start_week_price) * float(cur_price)) - (virt / float(start_week_imoex) * float(cur_imoex))
    except Exception: pass
    return month_pct, week_pct

def refresh_indicator(df, current_imoex):
    df = df.copy().reset_index(drop=True)
    df["Тек.цена актива"] = pd.to_numeric(df["Тек.цена актива"], errors="coerce").astype("float64")
    df["Тек. цена IMOEX"] = pd.to_numeric(df["Тек. цена IMOEX"], errors="coerce").astype("float64")
    numeric_cols = ["Вирт. Инвест", "Начало месяца активы", "Начало недели активы", "IMOEX сначала месяца", "Начало недели IMOEX", "Месяц % к IMOEX", "Неделя % к IMOEX"]
    for col in numeric_cols:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce").astype("float64")
    for idx in df.index:
        ticker = df.at[idx, "Тикер"]
        if pd.notna(ticker) and ticker:
            price = get_stock_price(ticker)
            if price is not None: df.loc[idx, "Тек.цена актива"] = float(price)
        if current_imoex is not None: df.loc[idx, "Тек. цена IMOEX"] = float(current_imoex)
    return df

def recalc_indicator(df, current_imoex):
    df = df.copy().reset_index(drop=True)
    numeric_cols = ["Вирт. Инвест", "Начало месяца активы", "Начало недели активы", "IMOEX сначала месяца", "Начало недели IMOEX", "Тек.цена актива", "Тек. цена IMOEX"]
    for col in numeric_cols:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce").astype(float)
    for idx, row in df.iterrows():
        month_pct, week_pct = calc_indicator_pct(row, current_imoex)
        df.at[idx, "Месяц % к IMOEX"] = month_pct
        df.at[idx, "Неделя % к IMOEX"] = week_pct
    return df

def get_open_price_first_day_of_month(ticker, is_index=False):
    today = date.today()
    start_date = date(today.year, today.month, 1)
    url = "https://iss.moex.com/iss/history/engines/stock/markets/index/securities/IMOEX.json" if is_index else f"https://iss.moex.com/iss/history/engines/stock/markets/shares/securities/{ticker}.json"
    params = {"from": start_date.strftime("%Y-%m-%d"), "till": today.strftime("%Y-%m-%d"), "iss.meta": "off", "start": 0}
    try:
        r = requests.get(url, params=params, timeout=15)
        data = r.json()
        if "history" not in data: return None
        cols, rows = data["history"]["columns"], data["history"]["data"]
        if not rows: return None
        idx_open = cols.index("OPEN") if "OPEN" in cols else None
        if idx_open is None: return None
        for row in rows:
            if row[idx_open] is not None: return round(float(row[idx_open]), 3)
        return None
    except Exception: return None

def get_open_price_first_day_of_week(ticker, is_index=False):
    today = date.today()
    # Находим понедельник текущей недели
    monday = today - pd.Timedelta(days=today.weekday())
    if hasattr(monday, "date"): 
        monday = monday.date()
        
    # Если запрашиваем в понедельник утром, то данных за сегодня может не быть.
    # Начнем с прошлого понедельника? Нет, нам нужна цена на начало ТЕКУЩЕЙ недели.
    
    url = "https://iss.moex.com/iss/history/engines/stock/markets/index/securities/IMOEX.json" if is_index else f"https://iss.moex.com/iss/history/engines/stock/markets/shares/securities/{ticker}.json"
    params = {"from": monday.strftime("%Y-%m-%d"), "till": today.strftime("%Y-%m-%d"), "iss.meta": "off", "start": 0}
    
    try:
        r = requests.get(url, params=params, timeout=15)
        data = r.json()
        if "history" not in data: return None
        
        cols, rows = data["history"]["columns"], data["history"]["data"]
        if not rows: return None
        
        idx_open = cols.index("OPEN") if "OPEN" in cols else None
        idx_trade_date = cols.index("TRADEDATE") if "TRADEDATE" in cols else None
        
        if idx_open is None or idx_trade_date is None: return None
        
        # Перебираем все дни с понедельника. MOEX отдает их по возрастанию.
        for row in rows:
            # Превращаем дату в объект datetime, чтобы узнать день недели
            trade_dt = pd.to_datetime(row[idx_trade_date])
            
            # 5 = Суббота, 6 = Воскресенье. Пропускаем их.
            if trade_dt.weekday() >= 5:
                continue
                
            # Если это будний день и есть цена открытия - берем ее
            if row[idx_open] is not None:
                return round(float(row[idx_open]), 3)
                
        return None
    except Exception: 
        return None

def auto_fill_indicator_fields(ticker):
    try:
        month_asset = get_open_price_first_day_of_month(ticker, is_index=False)
        month_imoex = get_open_price_first_day_of_month("IMOEX", is_index=True)
        week_asset = get_open_price_first_day_of_week(ticker, is_index=False)
        week_imoex = get_open_price_first_day_of_week("IMOEX", is_index=True)
    except Exception: return None
    if (month_asset is None or month_imoex is None or week_asset is None or week_imoex is None): return None
    today = date.today()
    return {"Начало месяца активы": month_asset, "IMOEX сначала месяца": month_imoex, "Начало недели активы": week_asset, "Начало недели IMOEX": week_imoex, "Обновлено_месяц": today, "Обновлено_неделя": today}

def check_and_update_indicator(current_imoex):
    if st.session_state.indicator.empty: return st.session_state.indicator
    df = st.session_state.indicator.copy()
    today = date.today()
    current_monday_dt = today - pd.Timedelta(days=today.weekday())
    current_monday = current_monday_dt.date() if hasattr(current_monday_dt, "date") else current_monday_dt
    imoex_month = get_open_price_first_day_of_month("IMOEX", is_index=True)
    imoex_week = get_open_price_first_day_of_week("IMOEX", is_index=True)
    updated = False
    for idx, row in df.iterrows():
        ticker = row.get("Тикер")
        if pd.isna(ticker) or not ticker: continue
        upd_month = pd.to_datetime(row.get("Обновлено_месяц"), errors="coerce")
        month_changed = (pd.isna(upd_month) or upd_month.date().month != today.month or upd_month.date().year != today.year)
        if month_changed:
            m_asset = get_open_price_first_day_of_month(ticker, is_index=False)
            if m_asset is not None:
                df.at[idx, "Начало месяца активы"] = m_asset
                df.at[idx, "Обновлено_месяц"] = today
                updated = True
            if imoex_month is not None: df.at[idx, "IMOEX сначала месяца"] = imoex_month
        upd_week = pd.to_datetime(row.get("Обновлено_неделя"), errors="coerce")
        week_changed = (pd.isna(upd_week) or upd_week.date() < current_monday)
        if week_changed:
            w_asset = get_open_price_first_day_of_week(ticker, is_index=False)
            if w_asset is not None:
                df.at[idx, "Начало недели активы"] = w_asset
                df.at[idx, "Обновлено_неделя"] = today
                updated = True
            if imoex_week is not None: df.at[idx, "Начало недели IMOEX"] = imoex_week
    if updated:
        st.session_state.indicator = df
        save_indicator(df)
    return df

def is_cache_fresh():
    cache = load_history_cache()
    if cache.empty: return False
    try:
        last = pd.to_datetime(cache["Обновлено"], errors="coerce").max()
        if pd.isna(last): return False
        return last.date() == date.today()
    except Exception: return False

def get_month_last_day(year, month):
    if month == 12: return date(year, 12, 31)
    return date(year, month + 1, 1) - pd.Timedelta(days=1)

def get_last_24_months():
    today = date.today()
    months = []
    y, m = today.year, today.month
    if m == 1: y -= 1; m = 12
    else: m -= 1
    for _ in range(24):
        months.append((y, m))
        if m == 1: y -= 1; m = 12
        else: m -= 1
    months.reverse()
    return months

def get_history_from_moex(ticker, is_index=False):
    months = get_last_24_months()
    if not months: return None
    start_year, start_month = months[0]
    end_year, end_month = months[-1]
    start_date = date(start_year, start_month, 1)
    end_date = get_month_last_day(end_year, end_month)
    url = "https://iss.moex.com/iss/history/engines/stock/markets/index/securities/IMOEX.json" if is_index else f"https://iss.moex.com/iss/history/engines/stock/markets/shares/securities/{ticker}.json"
    all_rows, start = [], 0
    while True:
        params = {"from": start_date.strftime("%Y-%m-%d"), "till": end_date.strftime("%Y-%m-%d"), "iss.meta": "off", "start": start}
        try:
            r = requests.get(url, params=params, timeout=20)
            data = r.json()
            if "history" not in data: break
            cols, rows = data["history"]["columns"], data["history"]["data"]
            if not rows: break
            idx_trade_date = cols.index("TRADEDATE") if "TRADEDATE" in cols else None
            idx_close = cols.index("CLOSE") if "CLOSE" in cols else None
            idx_open = cols.index("OPEN") if "OPEN" in cols else None
            idx_legal = cols.index("LEGALCLOSEPRICE") if "LEGALCLOSEPRICE" in cols else None
            if idx_trade_date is None or (idx_close is None and idx_legal is None): break
            for row in rows:
                trade_date = row[idx_trade_date]
                close_val = None
                open_val = None
                if idx_close is not None and row[idx_close] is not None: close_val = row[idx_close]
                elif idx_legal is not None and row[idx_legal] is not None: close_val = row[idx_legal]
                if idx_open is not None and row[idx_open] is not None: open_val = row[idx_open]
                if trade_date and close_val is not None:
                    all_rows.append({"date": pd.to_datetime(trade_date).date(), "open": float(open_val) if open_val is not None else None, "close": float(close_val)})
            if len(rows) < 100: break
            start += 100
        except Exception: break
    if not all_rows: return None
    return pd.DataFrame(all_rows).sort_values("date").reset_index(drop=True)

def calc_monthly_returns(df_history):
    if df_history is None or df_history.empty: return {}
    df_history = df_history.copy()
    df_history["year_month"] = df_history["date"].apply(lambda d: (d.year, d.month))
    monthly = {}
    for (y, m), group in df_history.groupby("year_month"):
        group = group.sort_values("date")
        first_row, last_row = group.iloc[0], group.iloc[-1]
        open_val, close_val = first_row.get("open"), last_row.get("close")
        if pd.notna(open_val) and pd.notna(close_val) and open_val != 0:
            monthly[(y, m)] = (close_val / open_val) - 1
    return monthly

def get_monthly_returns_for_ticker(ticker):
    df = get_history_from_moex(ticker, is_index=False)
    return calc_monthly_returns(df)

def get_monthly_returns_for_imoex():
    df = get_history_from_moex("IMOEX", is_index=True)
    return calc_monthly_returns(df)

def refresh_history_cache(tickers, force=False):
    if not force and is_cache_fresh(): return load_history_cache()
    imoex_returns = get_monthly_returns_for_imoex()
    all_rows = []
    for ticker in tickers:
        try: asset_returns = get_monthly_returns_for_ticker(ticker)
        except Exception: asset_returns = {}
        all_months = set(asset_returns.keys()) | set(imoex_returns.keys())
        for ym in sorted(all_months):
            y, m = ym
            month_label = f"{y}-{m:02d}"
            a, i = asset_returns.get(ym), imoex_returns.get(ym)
            diff = a - i if a is not None and i is not None else None
            all_rows.append({"Тикер": ticker, "Месяц": month_label, "% Акции": a, "% IMOEX": i, "Разница": diff, "Обновлено": date.today()})
    cache = pd.DataFrame(all_rows, columns=HISTORY_COLUMNS)
    save_history_cache(cache)
    return cache

def calc_averages_for_ticker(cache_df, ticker):
    df = cache_df[cache_df["Тикер"] == ticker].dropna(subset=["Разница"])
    if df.empty: return None, None, None
    df = df.sort_values("Месяц").reset_index(drop=True)
    values = df["Разница"].tolist()
    avg_2y = sum(values) / len(values) if values else None
    avg_1y = sum(values[-12:]) / len(values[-12:]) if len(values) >= 12 else None
    avg_3m = sum(values[-3:]) / len(values[-3:]) if len(values) >= 3 else None
    return avg_2y, avg_1y, avg_3m

if "purchases" not in st.session_state: st.session_state.purchases = load(PURCHASES_FILE, PURCHASE_COLUMNS)
if "sales" not in st.session_state: st.session_state.sales = load(SALES_FILE, SALE_COLUMNS)
if "distribution" not in st.session_state: st.session_state.distribution = load_distribution()
if "transactions" not in st.session_state: st.session_state.transactions = load_transactions()
if "indicator" not in st.session_state: st.session_state.indicator = load_indicator()
if "assets_to_imoex" not in st.session_state: st.session_state.assets_to_imoex = load_assets()
if "history_cache" not in st.session_state: st.session_state.history_cache = load_history_cache()
if "sell_targets" not in st.session_state: st.session_state.sell_targets = None
if "confirm_delete_purchases" not in st.session_state: st.session_state.confirm_delete_purchases = None
if "show_deposit_form" not in st.session_state: st.session_state.show_deposit_form = False
if "show_withdraw_form" not in st.session_state: st.session_state.show_withdraw_form = False
if "show_aimoex_detail" not in st.session_state: st.session_state.show_aimoex_detail = False
if "tx_to_delete" not in st.session_state: st.session_state.tx_to_delete = None

page = st.sidebar.radio("Меню", ["💼 Портфель", "📈 Индикатор", "📊 Активы к IMOEX", "🟢 Покупки", "📊 История продаж", "💰 Транзакции", "📦 Архив", "🗑 Корзина"], key="page")

if page == "💼 Портфель":
    st.subheader("💼 Портфель")
    visible_purchases = st.session_state.purchases[st.session_state.purchases["Удалено"] == False].copy()
    visible_sales = st.session_state.sales[st.session_state.sales["Удалено"] == False].copy()

    if not visible_purchases.empty:
        current_imoex_val = get_imoex()
        updated = refresh_open_purchases(recalc_status(visible_purchases, visible_sales, st.session_state.distribution), current_imoex_val)
        for idx_vis in updated.index:
            if idx_vis in st.session_state.purchases.index:
                st.session_state.purchases.at[idx_vis, "Тек.цена актива"] = updated.at[idx_vis, "Тек.цена актива"]
                st.session_state.purchases.at[idx_vis, "Тек. цена IMOEX"] = updated.at[idx_vis, "Тек. цена IMOEX"]
                st.session_state.purchases.at[idx_vis, "% к индексу"] = updated.at[idx_vis, "% к индексу"]
        save(st.session_state.purchases, PURCHASES_FILE)
        visible_purchases = st.session_state.purchases[st.session_state.purchases["Удалено"] == False].copy()

    total_invested = calc_total_invested()
    current_value = calc_current_value(visible_purchases, visible_sales)
    balance = get_balance()
    total = current_value + balance

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Инвестировано", f"{total_invested:,.2f} ₽")
    col2.metric("Стоимость", f"{current_value:,.2f} ₽")
    col3.metric("Баланс", f"{balance:,.2f} ₽")
    col4.metric("Итого", f"{total:,.2f} ₽")

    st.markdown("---")
    col_dep, col_wd, col_refresh = st.columns(3)
    with col_dep:
        if st.button("➕ Пополнить", key="deposit_button", use_container_width=True):
            st.session_state.show_deposit_form = True
            st.session_state.show_withdraw_form = False
            st.rerun()
    with col_wd:
        if st.button("➖ Вывести", key="withdraw_button", use_container_width=True):
            st.session_state.show_withdraw_form = True
            st.session_state.show_deposit_form = False
            st.rerun()
    with col_refresh:
        if st.button("🔄 Цены", key="refresh_portfolio", use_container_width=True):
            get_imoex.clear(); get_stock_price.clear(); st.rerun()

    if st.session_state.get("show_deposit_form"):
        with st.form("deposit_form", clear_on_submit=True):
            st.markdown("**➕ Пополнение баланса**")
            amount = st.number_input("Сумма, ₽", value=0.0, step=100.0, format="%.2f")
            comment = st.text_input("Комментарий")
            c1, c2 = st.columns(2)
            with c1:
                if st.form_submit_button("Пополнить", use_container_width=True):
                    if amount <= 0: st.warning("Введите сумму > 0")
                    else:
                        add_transaction("Пополнение", abs(amount), comment)
                        st.session_state.show_deposit_form = False
                        st.success(f"Пополнено: {amount:.2f} ₽")
                        st.rerun()
            with c2:
                if st.form_submit_button("Отмена", use_container_width=True):
                    st.session_state.show_deposit_form = False
                    st.rerun()

    if st.session_state.get("show_withdraw_form"):
        with st.form("withdraw_form", clear_on_submit=True):
            st.markdown("**➖ Вывод средств**")
            amount = st.number_input("Сумма, ₽", value=0.0, step=100.0, format="%.2f")
            comment = st.text_input("Комментарий")
            c1, c2 = st.columns(2)
            with c1:
                if st.form_submit_button("Вывести", use_container_width=True):
                    if amount <= 0: st.warning("Введите сумму > 0")
                    else:
                        add_transaction("Вывод", -abs(amount), comment)
                        st.session_state.show_withdraw_form = False
                        st.success(f"Выведено: {amount:.2f} ₽")
                        st.rerun()
            with c2:
                if st.form_submit_button("Отмена", use_container_width=True):
                    st.session_state.show_withdraw_form = False
                    st.rerun()

    st.markdown("---")
    st.subheader("📊 Мои активы")
    if visible_purchases.empty and visible_sales.empty:
        st.info("Пока нет активов.")
    else:
        all_names = set(visible_purchases["Наименование"].dropna().unique().tolist()) | set(visible_sales["Наименование"].dropna().unique().tolist())
        assets_rows = []
        for name in sorted(all_names):
            p_rows = visible_purchases[visible_purchases["Наименование"] == name]
            s_rows = visible_sales[visible_sales["Наименование"] == name]
            bought, sold, inv_purchases, inv_sales = 0, 0, 0.0, 0.0
            for _, r in p_rows.iterrows():
                try: bought += int(float(r["Кол-во"])); inv_purchases += float(r["Инвестировано"])
                except Exception: pass
            for _, r in s_rows.iterrows():
                try: sold += int(float(r["Кол-во"])); inv_sales += float(r["Сумма продажи"])
                except Exception: pass
            remaining = bought + sold
            invested_net = inv_purchases + inv_sales
            current_val = 0.0
            if remaining > 0 and not p_rows.empty:
                try:
                    price = p_rows.iloc[-1].get("Тек.цена актива")
                    if pd.notna(price): current_val = remaining * float(price)
                except Exception: pass
            profit = current_val - invested_net
            assets_rows.append({"Наименование": name, "Куплено": bought, "Продано": sold, "Остаток": remaining, "Инвестировано": invested_net, "Тек.стоимость": current_val, "Прибыль": profit})

        adf = pd.DataFrame(assets_rows)
        adf_show = adf.copy()
        adf_show["Инвестировано"] = adf_show["Инвестировано"].apply(lambda x: f"{float(x):,.2f} ₽")
        adf_show["Тек.стоимость"] = adf_show["Тек.стоимость"].apply(lambda x: f"{float(x):,.2f} ₽")
        adf_show["Прибыль"] = adf_show["Прибыль"].apply(lambda x: f"{'+' if float(x) >= 0 else ''}{float(x):,.2f} ₽")
        st.dataframe(adf_show, width="stretch", hide_index=True)

if page == "💰 Транзакции":
    st.subheader("💰 Транзакции")
    tx = st.session_state.transactions
    active_tx = tx[tx["Удалено"] == False] if not tx.empty else tx
    deposits = active_tx[active_tx["Тип"] == "Пополнение"]["Сумма"].sum() if not active_tx.empty else 0
    withdrawals = active_tx[active_tx["Тип"] == "Вывод"]["Сумма"].sum() if not active_tx.empty else 0
    balance = get_balance()

    col1, col2, col3 = st.columns(3)
    col1.metric("Всего пополнено", f"{float(deposits):+,.2f} ₽")
    col2.metric("Всего выведено", f"{float(withdrawals):+,.2f} ₽")
    col3.metric("Баланс", f"{float(balance):,.2f} ₽")

    st.markdown("---")
    st.subheader("Все транзакции")
    filter_type = st.selectbox("Фильтр по типу:", ["Пополнение и вывод", "Все", "Пополнение", "Вывод", "Покупка", "Продажа"], key="tx_filter")

    if active_tx.empty:
        st.info("Активных транзакций пока нет.")
    else:
        tx_show = active_tx.copy()
        if filter_type == "Пополнение и вывод": tx_show = tx_show[tx_show["Тип"].isin(["Пополнение", "Вывод"])]
        elif filter_type != "Все": tx_show = tx_show[tx_show["Тип"] == filter_type]

        if tx_show.empty:
            st.info(f"Транзакций типа «{filter_type}» нет.")
        else:
            tx_show["Дата"] = pd.to_datetime(tx_show["Дата"], errors="coerce").dt.strftime("%d.%m.%Y")
            tx_show["Сумма_fmt"] = tx_show["Сумма"].apply(lambda x: format(float(x), '+,.2f') if pd.notna(x) else "")
            tx_show = tx_show.sort_values("ID", ascending=False).reset_index(drop=True)
            tx_display = tx_show[["ID", "Дата", "Время", "Тип", "Сумма_fmt", "Комментарий"]].rename(columns={"Сумма_fmt": "Сумма"})
            tx_display.insert(0, "✓", False)

            edited_tx = st.data_editor(tx_display, width="stretch", hide_index=True, column_config={"✓": st.column_config.CheckboxColumn("✓", default=False, width="small"), "ID": st.column_config.NumberColumn("ID", format="%d", width="small")}, disabled=[c for c in tx_display.columns if c != "✓"], key="tx_editor")
            selected_tx = edited_tx[edited_tx["✓"] == True]

            if st.button("🗑 Удалить выбранные", key="tx_delete_button"):
                if selected_tx.empty: st.warning("Выберите транзакции")
                else:
                    st.session_state.tx_to_delete = selected_tx.to_dict('records')
                    st.rerun()

            if st.session_state.get("tx_to_delete"):
                selected_tx = pd.DataFrame(st.session_state.tx_to_delete)
                to_delete_deposits, to_delete_other = [], []
                total_purchases_to_del, total_sales_to_del = 0, 0
                for _, row in selected_tx.iterrows():
                    tx_id = _as_int(row["ID"])
                    if tx_id is None: continue
                    if row["Тип"] == "Пополнение":
                        to_delete_deposits.append(tx_id)
                        total_purchases_to_del += len(get_purchases_by_parent(tx_id))
                        total_sales_to_del += len(get_sales_by_parent(tx_id))
                    else: to_delete_other.append(tx_id)

                st.markdown("---")
                st.warning(f"⚠️ **Вы уверены, что хотите удалить выбранные транзакции?**\n\n- Пополнений: **{len(to_delete_deposits)}**\n- Прочих транзакций: **{len(to_delete_other)}**\n- Связанных покупок: **{total_purchases_to_del}**\n- Связанных продаж: **{total_sales_to_del}**\n\nВсе эти данные будут перемещены в **Корзину** (раздел «🗑 Корзина»), откуда их можно будет восстановить.")

                col_yes, col_no = st.columns(2)
                with col_yes:
                    if st.button("✅ Да, удалить", key="confirm_tx_delete", type="primary", use_container_width=True):
                        total_deposits, total_transactions, total_purchases, total_sales = 0, 0, 0, 0
                        for tx_id in to_delete_deposits:
                            stats = cascade_delete_deposit(tx_id)
                            total_deposits += stats["deposit"]; total_transactions += stats["transactions"]; total_purchases += stats["purchases"]; total_sales += stats["sales"]
                        for tx_id in to_delete_other:
                            tx_mask = st.session_state.transactions["ID"].apply(_as_int) == tx_id
                            if tx_mask.any():
                                st.session_state.transactions.loc[tx_mask, "Удалено"] = True
                                total_transactions += int(tx_mask.sum())

                        save_transactions(st.session_state.transactions)
                        save(st.session_state.purchases, PURCHASES_FILE)
                        save(st.session_state.sales, SALES_FILE)
                        st.success(f"Удалено в корзину: пополнений — {total_deposits}, транзакций — {total_transactions}, покупок — {total_purchases}, продаж — {total_sales}. Восстановить можно в разделе «🗑 Корзина».")
                        st.session_state.tx_to_delete = None
                        st.rerun()
                with col_no:
                    if st.button("❌ Отмена", key="cancel_tx_delete", use_container_width=True):
                        st.session_state.tx_to_delete = None
                        st.rerun()

if page == "🟢 Покупки":
    col_im, col_balance, col_btn = st.columns([2, 1, 1])
    with col_im:
        current_imoex = get_imoex()
        if current_imoex is not None:
            st.markdown(f"### 📈 IMOEX: **{current_imoex:.3f}**")
            st.caption("Обновляется автоматически раз в 5 минут")
        else:
            st.warning("⚠️ Не удалось получить IMOEX с MOEX")
            current_imoex = None
    with col_balance: st.metric("💰 Баланс", f"{get_balance():,.2f} ₽")
    with col_btn:
        if st.button("🔄 Обновить", key="refresh_purchases", use_container_width=True):
            get_imoex.clear(); get_stock_price.clear(); st.rerun()

    st.markdown("---")
    with st.expander("➕ Добавить покупку", expanded=False):
        with st.form("buy_form", clear_on_submit=True):
            deal_date = st.date_input("Дата покупки", value=date.today())
            query = st.text_input("Введите название (например: Лукойл, Сбер)")
            selected_ticker, selected_name, allow_no_ticker = None, None, False
            if query and len(query) >= 2:
                found = find_tickers(query)
                if found:
                    options = ["— выберите тикер —"] + [f["label"] for f in found]
                    chosen = st.selectbox("Найдено на MOEX:", options)
                    if chosen != "— выберите тикер —":
                        selected = next(f for f in found if f["label"] == chosen)
                        selected_ticker, selected_name = selected["ticker"], selected["name"]
                        st.caption(f"✅ Тикер: `{selected_ticker}` — {selected_name}")
                else:
                    st.warning("⚠️ Бумага не найдена на MOEX.")
                    allow_no_ticker = st.checkbox("Добавить без тикера")

            c1, c2 = st.columns(2)
            with c1:
                price = st.number_input("Цена актива", value=0.0, step=0.001, format="%.3f")
                quantity = st.number_input("Кол-во", value=1, step=1)
            with c2:
                imoex_buy = st.number_input("Покупка IMOEX", value=0.0, step=0.01, format="%.3f")
                commission = st.number_input("Комиссия, ₽", value=0.0, step=1.0, format="%.3f")

            submitted = st.form_submit_button("Добавить покупку")
            if submitted:
                if not query: st.warning("Введите название")
                elif selected_ticker is None and not allow_no_ticker: st.warning("Выберите тикер")
                elif price <= 0: st.warning("Введите цену > 0")
                elif imoex_buy <= 0: st.warning("Введите IMOEX > 0")
                else:
                    invested = price * quantity + commission
                    if get_balance() < invested:
                        st.error(f"❌ Недостаточно средств. Нужно: {invested:,.2f} ₽, на балансе: {get_balance():,.2f} ₽")
                    else:
                        virt_imoex = invested / imoex_buy
                        new_id = get_next_purchase_id(st.session_state.purchases)
                        new_row = {"ID": new_id, "Дата": deal_date, "Тикер": selected_ticker, "Наименование": selected_name if selected_name else query, "Цена актива": price, "Кол-во": quantity, "Комиссия": commission, "Инвестировано": round(invested, 3), "Покупка IMOEX": imoex_buy, "Вирт. Кол-во IMOEX": round(virt_imoex, 6), "Тек.цена актива": None, "Тек. цена IMOEX": None, "% к индексу": None, "Статус": "ОТКРЫТА", "Удалено": False}
                        st.session_state.purchases = pd.concat([st.session_state.purchases, pd.DataFrame([new_row])], ignore_index=True)
                        save(st.session_state.purchases, PURCHASES_FILE)
                        active_dep = st.session_state.transactions[(st.session_state.transactions["Тип"] == "Пополнение") & (st.session_state.transactions["Удалено"] == False)]
                        parent_tx_id = _as_int(active_dep.iloc[-1]["ID"]) if not active_dep.empty else None
                        create_purchase_transaction(new_id, invested, selected_name if selected_name else query, parent_id=parent_tx_id)
                        st.success(f"Добавлено: {selected_name or query}, {quantity} шт. (ID: {new_id})")
                        st.rerun()

    st.subheader("Активные покупки")
    visible_purchases = st.session_state.purchases[st.session_state.purchases["Удалено"] == False].copy()
    visible_sales = st.session_state.sales[st.session_state.sales["Удалено"] == False].copy()

    if visible_purchases.empty: st.info("Пока пусто. Добавьте первую покупку.")
    else:
        all_display = refresh_open_purchases(recalc_status(visible_purchases, visible_sales, st.session_state.distribution), current_imoex)
        for idx_vis in all_display.index:
            if idx_vis in st.session_state.purchases.index:
                st.session_state.purchases.at[idx_vis, "Тек.цена актива"] = all_display.at[idx_vis, "Тек.цена актива"]
                st.session_state.purchases.at[idx_vis, "Тек. цена IMOEX"] = all_display.at[idx_vis, "Тек. цена IMOEX"]
                st.session_state.purchases.at[idx_vis, "% к индексу"] = all_display.at[idx_vis, "% к индексу"]
        save(st.session_state.purchases, PURCHASES_FILE)

        active = all_display[all_display["Статус"] != "ЗАКРЫТА"].copy()
        if active.empty: st.info("Нет активных покупок. Все закрыты — смотри раздел «Архив».")
        else:
            active_show = active.copy()
            active_show["Дата"] = pd.to_datetime(active_show["Дата"], errors="coerce").dt.strftime("%d.%m.%Y")
            for drop_col in ["Комиссия", "Вирт. Кол-во IMOEX", "Удалено"]:
                if drop_col in active_show.columns: active_show = active_show.drop(columns=[drop_col])
            active_show = active_show[["Дата", "Тикер", "Наименование", "Кол-во", "Остаток", "Цена актива", "Покупка IMOEX", "Инвестировано", "Тек.цена актива", "Тек. цена IMOEX", "% к индексу", "Статус", "ID"]].copy()
            for c in ["Кол-во", "Остаток", "Цена актива", "Покупка IMOEX", "Инвестировано", "Тек.цена актива", "Тек. цена IMOEX"]:
                if c in active_show.columns: active_show[c] = active_show[c].apply(format_number)
            if "% к индексу" in active_show.columns: active_show["% к индексу"] = active_show["% к индексу"].apply(lambda x: f"{float(x) * 100:.3f}%" if pd.notna(x) and x is not None else "")
            active_show.insert(0, "✓", False)

            edited = st.data_editor(active_show, width="stretch", hide_index=True, column_config={"✓": st.column_config.CheckboxColumn("✓", default=False, width="small"), "ID": st.column_config.NumberColumn("ID", format="%d", width="small")}, disabled=[c for c in active_show.columns if c != "✓"], key="active_editor")
            selected_rows = edited[edited["✓"] == True]

            col_sell, col_del = st.columns(2)
            with col_sell:
                if st.button("💸 Продать выбранное", key="sell_button", use_container_width=True):
                    if selected_rows.empty: st.warning("Выберите строки (поставьте галочки)")
                    else:
                        tickers = selected_rows["Тикер"].dropna().unique().tolist()
                        if len(tickers) > 1: st.error(f"❌ Выбраны разные тикеры: {tickers}. Продавать можно только один тикер за раз.")
                        else:
                            total_remaining = sum(int(float(r["Остаток"])) for _, r in selected_rows.iterrows() if pd.notna(r["Остаток"]))
                            st.session_state.sell_targets = {"rows": selected_rows.to_dict("records"), "ticker": tickers[0] if tickers else "", "total_remaining": total_remaining}

            with col_del:
                if st.button("🗑 Пометить как удалённое", key="delete_purchase_button", use_container_width=True):
                    if selected_rows.empty: st.warning("Выберите строки (поставьте галочки)")
                    else:
                        to_mark = [all_display[all_display["ID"] == row["ID"]].index[0] for idx, row in selected_rows.iterrows() if pd.notna(row.get("ID")) and not all_display[all_display["ID"] == row["ID"]].empty]
                        if to_mark:
                            st.session_state.confirm_delete_purchases = to_mark
                            st.rerun()

            if st.session_state.get("confirm_delete_purchases"):
                to_mark = st.session_state.confirm_delete_purchases
                st.warning(f"⚠️ Будет помечено как удалённое:\n- Покупок: **{len(to_mark)}**\n\nДанные останутся в корзине, их можно будет восстановить.")
                col_yes, col_no = st.columns(2)
                with col_yes:
                    if st.button("✅ Да, пометить", key="confirm_yes", use_container_width=True):
                        for p_idx in to_mark:
                            if p_idx in st.session_state.purchases.index:
                                pid = st.session_state.purchases.loc[p_idx, "ID"]
                                if pd.notna(pid): cascade_delete_purchase(pid)
                        save(st.session_state.purchases, PURCHASES_FILE); save(st.session_state.sales, SALES_FILE)
                        st.session_state.confirm_delete_purchases = None
                        st.success(f"Помечено: покупок {len(to_mark)}")
                        st.rerun()
                with col_no:
                    if st.button("❌ Отмена", key="confirm_no", use_container_width=True):
                        st.session_state.confirm_delete_purchases = None
                        st.rerun()

            if st.session_state.sell_targets is not None:
                targets = st.session_state.sell_targets
                rows_list = targets["rows"]
                total_remaining = targets["total_remaining"]
                st.markdown("---")
                st.markdown(f"### 💸 Продажа: **{targets['ticker']}** (общий остаток: {total_remaining})")
                with st.form("sell_form"):
                    sc1, sc2 = st.columns(2)
                    with sc1:
                        sell_qty = st.number_input("Кол-во (общее)", value=1, step=1, min_value=1, max_value=int(total_remaining))
                        sell_price = st.number_input("Цена продажи", value=0.0, step=0.001, format="%.3f")
                    with sc2:
                        sell_imoex = st.number_input("IMOEX продажи", value=float(current_imoex) if current_imoex else 0.0, step=0.01, format="%.3f")
                        sell_commission = st.number_input("Комиссия, ₽", value=0.0, step=1.0, format="%.3f")
                    submitted = st.form_submit_button("💸 Продать")
                    if submitted:
                        if sell_price <= 0: st.warning("Введите цену продажи > 0")
                        elif sell_imoex <= 0: st.warning("Введите IMOEX > 0")
                        elif sell_qty <= 0: st.warning("Введите кол-во > 0")
                        elif sell_qty > total_remaining: st.error(f"Общий остаток только {total_remaining}")
                        else:
                            sorted_rows = sorted(rows_list, key=lambda r: int(r.get("ID", 0)))
                            remaining_to_sell, distribution_rows, used_ids, weighted_buy_price_sum, weighted_buy_imoex_sum = int(sell_qty), [], [], 0, 0
                            for r in sorted_rows:
                                if remaining_to_sell <= 0: break
                                avail = int(float(r.get("Остаток", 0)))
                                if avail <= 0: continue
                                take = min(avail, remaining_to_sell)
                                distribution_rows.append({"Purchase_ID": int(r["ID"]), "Qty": take})
                                used_ids.append(int(r["ID"]))
                                weighted_buy_price_sum += float(r["Цена актива"]) * take
                                weighted_buy_imoex_sum += float(r["Покупка IMOEX"]) * take
                                remaining_to_sell -= take
                            if remaining_to_sell > 0: st.error("Не удалось распределить всё количество")
                            else:
                                avg_buy_price, avg_buy_imoex = weighted_buy_price_sum / sell_qty, weighted_buy_imoex_sum / sell_qty
                                sale_sum = sell_price * sell_qty - sell_commission
                                virt_imoex = (avg_buy_price * sell_qty) / avg_buy_imoex if avg_buy_imoex != 0 else 0
                                pct = (virt_imoex * sell_imoex) / sale_sum - 1 if sale_sum != 0 else None
                                oldest_date = min(pd.to_datetime(r["Дата"]) for r in rows_list if pd.notna(r.get("Дата")))
                                new_sale_id = get_next_sale_id(st.session_state.sales)
                                new_sale = {"Sale_ID": new_sale_id, "Purchase_IDs": ",".join(str(x) for x in used_ids), "Дата продажи": date.today(), "Тикер": sorted_rows[0].get("Тикер"), "Наименование": sorted_rows[0]["Наименование"], "Цена продажи": sell_price, "Кол-во": -sell_qty, "Комиссия": sell_commission, "Сумма продажи": -round(sale_sum, 3), "Продажа IMOEX": sell_imoex, "Вирт. Кол-во IMOEX": -round(virt_imoex, 6), "Цена покупки актива": round(avg_buy_price, 3), "Покупка IMOEX": round(avg_buy_imoex, 3), "% к индексу": round(pct, 6) if pct is not None else None, "Примечание": f"закрыта за {oldest_date.strftime('%d.%m.%Y')}", "Удалено": False}
                                st.session_state.sales = pd.concat([st.session_state.sales, pd.DataFrame([new_sale])], ignore_index=True)
                                save(st.session_state.sales, SALES_FILE)
                                for d in distribution_rows: d["Sale_ID"] = new_sale_id
                                st.session_state.distribution = pd.concat([st.session_state.distribution, pd.DataFrame(distribution_rows)], ignore_index=True)
                                save_distribution(st.session_state.distribution)
                                parent_tx_id = next((get_parent_id_for_purchase(pid) for pid in sorted(used_ids) if get_parent_id_for_purchase(pid) is not None), None)
                                create_sale_transaction(new_sale_id, sale_sum, sorted_rows[0]["Наименование"], parent_id=parent_tx_id)
                                st.session_state.sell_targets = None
                                st.success(f"Продано: {sell_qty} шт. (Sale_ID: {new_sale_id})")
                                st.rerun()
                if st.button("❌ Отмена", key="cancel_sell"):
                    st.session_state.sell_targets = None; st.rerun()

if page == "📦 Архив":
    st.subheader("📦 Архив закрытых покупок")
    st.caption("Только просмотр. Чтобы вернуть — удалите продажу в разделе «История продаж».")
    visible_purchases = st.session_state.purchases[st.session_state.purchases["Удалено"] == False].copy()
    visible_sales = st.session_state.sales[st.session_state.sales["Удалено"] == False].copy()

    if visible_purchases.empty: st.info("Пока пусто.")
    else:
        archive = recalc_status(visible_purchases, visible_sales, st.session_state.distribution)
        archive = archive[archive["Статус"] == "ЗАКРЫТА"].copy()
        if archive.empty: st.info("Архив пуст. Закрытых покупок нет.")
        else:
            archive_show = archive.copy()
            archive_show["Дата"] = pd.to_datetime(archive_show["Дата"], errors="coerce").dt.strftime("%d.%m.%Y")
            for drop_col in ["Комиссия", "Вирт. Кол-во IMOEX", "Удалено"]:
                if drop_col in archive_show.columns: archive_show = archive_show.drop(columns=[drop_col])
            archive_show = archive_show[["Дата", "Тикер", "Наименование", "Кол-во", "Остаток", "Цена актива", "Покупка IMOEX", "Инвестировано", "Тек.цена актива", "Тек. цена IMOEX", "% к индексу", "Статус", "ID"]].copy()
            for c in ["Кол-во", "Остаток", "Цена актива", "Покупка IMOEX", "Инвестировано", "Тек.цена актива", "Тек. цена IMOEX"]:
                if c in archive_show.columns: archive_show[c] = archive_show[c].apply(format_number)
            if "% к индексу" in archive_show.columns: archive_show["% к индексу"] = archive_show["% к индексу"].apply(lambda x: f"{float(x) * 100:.3f}%" if pd.notna(x) and x is not None else "")
            st.dataframe(archive_show, width="stretch", hide_index=True, column_config={"ID": st.column_config.NumberColumn("ID", format="%d", width="small")})

if page == "🗑 Корзина":
    st.subheader("🗑 Корзина")
    st.caption("Удалённые транзакции. Можно восстановить или очистить навсегда.")
    deleted_tx = st.session_state.transactions[st.session_state.transactions["Удалено"] == True].copy()

    if deleted_tx.empty:
        st.info("Корзина пуста.")
    else:
        st.markdown(f"**Удалённые транзакции** ({len(deleted_tx)})")
        tx_del_show = deleted_tx.copy()
        tx_del_show["Дата"] = pd.to_datetime(tx_del_show["Дата"], errors="coerce").dt.strftime("%d.%m.%Y")
        tx_del_show["Сумма_fmt"] = tx_del_show["Сумма"].apply(lambda x: format(float(x), '+,.2f') if pd.notna(x) else "")
        tx_del_display = tx_del_show[["ID", "Дата", "Время", "Тип", "Сумма_fmt", "Комментарий"]].rename(columns={"Сумма_fmt": "Сумма"})
        tx_del_display.insert(0, "✓", False)

        edited_tx_del = st.data_editor(tx_del_display, width="stretch", hide_index=True, column_config={"✓": st.column_config.CheckboxColumn("✓", default=False, width="small"), "ID": st.column_config.NumberColumn("ID", format="%d", width="small")}, disabled=[c for c in tx_del_display.columns if c != "✓"], key="deleted_tx_editor")
        selected_tx_del = edited_tx_del[edited_tx_del["✓"] == True]

        if st.button("♻️ Восстановить выбранные транзакции", key="restore_tx"):
            if selected_tx_del.empty: st.warning("Выберите строки")
            else:
                restored_deposits, restored_tx = 0, 0
                for _, row in selected_tx_del.iterrows():
                    tx_id = _as_int(row["ID"]); tx_type = row["Тип"]
                    if tx_id is None: continue
                    if tx_type == "Пополнение":
                        stats = restore_deposit_cascade(tx_id)
                        restored_deposits += stats["deposit"]; restored_tx += stats["transactions"]
                    else:
                        tx_mask = st.session_state.transactions["ID"].apply(_as_int) == tx_id
                        if tx_mask.any():
                            st.session_state.transactions.loc[tx_mask, "Удалено"] = False
                            restored_tx += int(tx_mask.sum())
                save_transactions(st.session_state.transactions)
                save(st.session_state.purchases, PURCHASES_FILE)
                save(st.session_state.sales, SALES_FILE)
                st.success(f"Восстановлено транзакций: {restored_tx}.")
                st.rerun()

        st.markdown("---")
        st.warning("⚠️ **Служебные действия**")
        if st.button("🧹 Удалить покупки без пополнений", key="clean_orphans"):
            count = find_and_clean_orphan_purchases()
            if count > 0: st.success(f"Найдено и удалено «фантомных» покупок: {count}."); st.rerun()
            else: st.info("Фантомных покупок не найдено. Всё чисто!")

        if st.button("🔥 Очистить корзину навсегда", key="clear_trash"):
            st.session_state.transactions = st.session_state.transactions[st.session_state.transactions["Удалено"] == False].reset_index(drop=True)
            st.session_state.purchases = st.session_state.purchases[st.session_state.purchases["Удалено"] == False].reset_index(drop=True)
            st.session_state.sales = st.session_state.sales[st.session_state.sales["Удалено"] == False].reset_index(drop=True)
            save_transactions(st.session_state.transactions); save(st.session_state.purchases, PURCHASES_FILE); save(st.session_state.sales, SALES_FILE)
            st.success("Корзина очищена. Данные удалены навсегда."); st.rerun()

if page == "📊 История продаж":
    st.subheader("📊 История продаж")
    visible_sales = st.session_state.sales[st.session_state.sales["Удалено"] == False].copy()

    if visible_sales.empty: st.info("Пока пусто. Продажи делаются в разделе «Покупки».")
    else:
        sales_display = visible_sales.copy()
        sales_display["Дата продажи"] = pd.to_datetime(sales_display["Дата продажи"], errors="coerce").dt.strftime("%d.%m.%Y")
        for drop_col in ["Комиссия", "Удалено"]:
            if drop_col in sales_display.columns: sales_display = sales_display.drop(columns=[drop_col])
        if "% к индексу" in sales_display.columns: sales_display["% к индексу"] = sales_display["% к индексу"].apply(lambda x: f"{-x * 100:.3f}%" if pd.notna(x) and x is not None else "")
        for col in ["Цена продажи", "Кол-во", "Сумма продажи", "Продажа IMOEX", "Цена покупки актива", "Покупка IMOEX"]:
            if col in sales_display.columns: sales_display[col] = sales_display[col].apply(format_number)
        if "Вирт. Кол-во IMOEX" in sales_display.columns: sales_display["Вирт. Кол-во IMOEX"] = sales_display["Вирт. Кол-во IMOEX"].apply(lambda x: f"{float(x):.2f}" if pd.notna(x) and x is not None else "")
        cols_order = [c for c in sales_display.columns if c != "Sale_ID"]
        if "Sale_ID" in sales_display.columns: cols_order.append("Sale_ID")
        sales_display = sales_display[cols_order]
        sales_display.insert(0, "✓", False)

        edited_sales = st.data_editor(sales_display, width="stretch", hide_index=True, column_config={"✓": st.column_config.CheckboxColumn("✓", default=False, width="small"), "Sale_ID": st.column_config.NumberColumn("Sale_ID", format="%d", width="small")}, disabled=[c for c in sales_display.columns if c != "✓"], key="sales_editor")
        selected_sales = edited_sales[edited_sales["✓"] == True]

        if st.button("🗑 Пометить продажи как удалённые", key="delete_sales"):
            if selected_sales.empty: st.warning("Выберите строки (поставьте галочки)")
            else:
                to_mark = [visible_sales[visible_sales["Sale_ID"] == row.get("Sale_ID")].index[0] for idx, row in selected_sales.iterrows() if pd.notna(row.get("Sale_ID")) and not visible_sales[visible_sales["Sale_ID"] == row.get("Sale_ID")].empty]
                if to_mark:
                    for s_idx in to_mark:
                        s_row = st.session_state.sales.loc[s_idx]
                        sid = _as_int(s_row.get("Sale_ID"))
                        pids = _parse_purchase_ids(s_row.get("Purchase_IDs", ""))
                        st.session_state.sales.loc[s_idx, "Удалено"] = True
                        mark_transaction_deleted(sid)
                        for pid in pids: cascade_delete_purchase(pid)
                    save(st.session_state.sales, SALES_FILE)
                    st.success(f"Помечено продаж: {len(to_mark)}. Можно восстановить в Корзине.")
                    st.rerun()

if page == "📈 Индикатор":
    col_im, col_btn = st.columns([3, 1])
    with col_im:
        current_imoex = get_imoex()
        if current_imoex is not None:
            st.markdown(f"### 📈 IMOEX: **{current_imoex:.3f}**")
            st.caption("Обновляется автоматически раз в 5 минут")
        else:
            st.warning("⚠️ Не удалось получить IMOEX с MOEX")
            current_imoex = None
    with col_btn:
        if st.button("🔄 Обновить", key="refresh_indicator"):
            get_imoex.clear(); get_stock_price.clear()
            if not st.session_state.indicator.empty:
                df_reset = st.session_state.indicator.copy()
                df_reset["Обновлено_месяц"] = None; df_reset["Обновлено_неделя"] = None
                st.session_state.indicator = df_reset; save_indicator(df_reset)
            st.rerun()

    if not st.session_state.indicator.empty: check_and_update_indicator(current_imoex)

    st.markdown("---")
    with st.expander("➕ Добавить актив", expanded=False):
        with st.form("ind_add_form", clear_on_submit=True):
            query = st.text_input("Введите название или тикер (например: Сбер, FESH, ДВМП)")
            
            selected_ticker = None
            selected_name = None
            
            if query and len(query) >= 2:
                found = find_tickers(query)
                if found:
                    options = ["— выберите тикер —"] + [f["label"] for f in found]
                    chosen = st.selectbox("Найдено на MOEX:", options)
                    if chosen != "— выберите тикер —":
                        selected = next(f for f in found if f["label"] == chosen)
                        selected_ticker = selected["ticker"]
                        selected_name = selected["name"]
                        st.caption(f"✅ Тикер: `{selected_ticker}` — {selected_name}")
                else:
                    st.warning("⚠️ Бумага не найдена на MOEX.")

            submitted = st.form_submit_button("Добавить актив")
            
            if submitted:
                if selected_ticker is None:
                    st.warning("Сначала выберите тикер из списка")
                elif not st.session_state.indicator.empty and selected_ticker in st.session_state.indicator["Тикер"].values:
                    st.error(f"❌ Актив {selected_ticker} уже есть в индикаторе!")
                else:
                    with st.spinner("Загрузка данных с MOEX..."):
                        auto_data = auto_fill_indicator_fields(selected_ticker)

                    if auto_data is None:
                        st.error("❌ MOEX недоступен или данные не получены. Попробуйте позже.")
                    else:
                        new_row = {
                            "Наименование": selected_name if selected_name else query,
                            "Тикер": selected_ticker,
                            "Вирт. Инвест": 1,
                            "Начало месяца активы": auto_data["Начало месяца активы"],
                            "Начало недели активы": auto_data["Начало недели активы"],
                            "IMOEX сначала месяца": auto_data["IMOEX сначала месяца"],
                            "Начало недели IMOEX": auto_data["Начало недели IMOEX"],
                            "Тек.цена актива": None,
                            "Тек. цена IMOEX": None,
                            "Месяц % к IMOEX": None,
                            "Неделя % к IMOEX": None,
                            "Обновлено_месяц": auto_data["Обновлено_месяц"],
                            "Обновлено_неделя": auto_data["Обновлено_неделя"],
                        }
                        st.session_state.indicator = pd.concat(
                            [st.session_state.indicator, pd.DataFrame([new_row])],
                            ignore_index=True
                        )
                        save_indicator(st.session_state.indicator)
                        st.success(f"Добавлено: {selected_name or query}")
                        st.rerun()

    st.subheader("Индикатор по IMOEX")
    if st.session_state.indicator.empty: st.info("Пока пусто. Добавьте первый актив.")
    else:
        updated = recalc_indicator(refresh_indicator(st.session_state.indicator, current_imoex), current_imoex)
        st.session_state.indicator = updated; save_indicator(st.session_state.indicator)
        ind_show = updated.copy()
        for c in ["Вирт. Инвест", "Начало месяца активы", "Начало недели активы", "IMOEX сначала месяца", "Начало недели IMOEX", "Тек.цена актива", "Тек. цена IMOEX"]:
            if c in ind_show.columns: ind_show[c] = ind_show[c].apply(format_number)
        if "Месяц % к IMOEX" in ind_show.columns: ind_show["Месяц % к IMOEX"] = ind_show["Месяц % к IMOEX"].apply(lambda x: f"{float(x) * 100:.3f}%" if pd.notna(x) and x is not None else "")
        if "Неделя % к IMOEX" in ind_show.columns: ind_show["Неделя % к IMOEX"] = ind_show["Неделя % к IMOEX"].apply(lambda x: f"{float(x) * 100:.3f}%" if pd.notna(x) and x is not None else "")
        for drop_col in ["Вирт. Инвест", "Тикер"]:
            if drop_col in ind_show.columns: ind_show = ind_show.drop(columns=[drop_col])
        ind_show = ind_show[["Наименование", "Начало месяца активы", "Начало недели активы", "IMOEX сначала месяца", "Начало недели IMOEX", "Тек.цена актива", "Тек. цена IMOEX", "Месяц % к IMOEX", "Неделя % к IMOEX"]].copy()
        ind_show.insert(0, "✓", False)

        edited_ind = st.data_editor(ind_show, width="stretch", hide_index=True, column_config={"✓": st.column_config.CheckboxColumn("✓", default=False, width="small")}, disabled=[c for c in ind_show.columns if c != "✓"], key="indicator_editor")
        selected_ind = edited_ind[edited_ind["✓"] == True]

        if st.button("🗑 Удалить выбранные", key="delete_indicator"):
            if selected_ind.empty: st.warning("Выберите строки")
            else:
                to_delete = [st.session_state.indicator[st.session_state.indicator["Наименование"] == row["Наименование"]].index[0] for idx, row in selected_ind.iterrows() if not st.session_state.indicator[st.session_state.indicator["Наименование"] == row["Наименование"]].empty]
                if to_delete:
                    st.session_state.indicator = st.session_state.indicator.drop(to_delete).reset_index(drop=True)
                    save_indicator(st.session_state.indicator); st.success(f"Удалено активов: {len(to_delete)}"); st.rerun()

if page == "📊 Активы к IMOEX":
    st.subheader("📊 Активы к IMOEX")
    st.caption("Сравнение доходности активов с индексом IMOEX за последние 24 закрытых месяца")

    col_btn1, _ = st.columns([1, 4])
    with col_btn1:
        if st.button("🔄 Обновить историю", key="refresh_history_button", use_container_width=True):
            with st.spinner("Загрузка истории с MOEX... Это может занять 30–60 секунд."):
                tickers = st.session_state.assets_to_imoex["Тикер"].dropna().unique().tolist()
                if tickers:
                    st.session_state.history_cache = refresh_history_cache(tickers, force=True)
                    st.success(f"Обновлено. Активов: {len(tickers)}"); st.rerun()
                else: st.warning("Сначала добавьте активы")

    st.markdown("---")
    with st.expander("➕ Добавить актив", expanded=False):
        with st.form("aimoex_add_form", clear_on_submit=True):
            query = st.text_input("Введите название или тикер (например: Сбер, FESH, ДВМП)")
            
            selected_ticker = None
            selected_name = None
            
            if query and len(query) >= 2:
                found = find_tickers(query)
                if found:
                    options = ["— выберите тикер —"] + [f["label"] for f in found]
                    chosen = st.selectbox("Найдено на MOEX:", options)
                    if chosen != "— выберите тикер —":
                        selected = next(f for f in found if f["label"] == chosen)
                        selected_ticker = selected["ticker"]
                        selected_name = selected["name"]
                        st.caption(f"✅ Тикер: `{selected_ticker}` — {selected_name}")
                else:
                    st.warning("⚠️ Бумага не найдена на MOEX.")

            submitted = st.form_submit_button("Добавить актив")
            
            if submitted:
                if selected_ticker is None:
                    st.warning("Сначала выберите тикер из списка")
                elif selected_ticker in st.session_state.assets_to_imoex["Тикер"].values:
                    st.warning("Этот актив уже добавлен")
                else:
                    new_row = {
                        "Тикер": selected_ticker,
                        "Наименование": selected_name if selected_name else query,
                        "Добавлено": date.today(),
                    }
                    st.session_state.assets_to_imoex = pd.concat(
                        [st.session_state.assets_to_imoex, pd.DataFrame([new_row])],
                        ignore_index=True
                    )
                    save_assets(st.session_state.assets_to_imoex)

                    with st.spinner("Загрузка истории с MOEX..."):
                        tickers = st.session_state.assets_to_imoex["Тикер"].dropna().unique().tolist()
                        cache = refresh_history_cache(tickers, force=True)
                        st.session_state.history_cache = cache

                    st.success(f"Добавлено и загружено: {selected_name or query}")
                    st.rerun()

    st.subheader("Сравнение с IMOEX")
    if st.session_state.assets_to_imoex.empty: st.info("Пока пусто. Добавьте актив.")
    else:
        cache = st.session_state.history_cache
        if cache.empty: st.warning("⚠️ История не загружена. Нажмите «🔄 Обновить историю».")
        else:
            if not is_cache_fresh(): st.info("ℹ️ Кэш истории устарел. Нажмите «🔄 Обновить историю».")
            main_rows = []
            for _, asset in st.session_state.assets_to_imoex.iterrows():
                ticker, name = asset["Тикер"], asset["Наименование"]
                avg_2y, avg_1y, avg_3m = calc_averages_for_ticker(cache, ticker)
                current_month_diff = None
                try:
                    df_cache = cache[cache["Тикер"] == ticker].dropna(subset=["Разница"]).sort_values("Месяц")
                    if not df_cache.empty: current_month_diff = df_cache.iloc[-1]["Разница"]
                except Exception: pass
                main_rows.append({"Наименование": name, "Ср. за 2 года": avg_2y, "Ср. за год": avg_1y, "Ср. за 3 мес": avg_3m, "Последний месяц": current_month_diff})

            main_df = pd.DataFrame(main_rows)
            for col in ["Ср. за 2 года", "Ср. за год", "Ср. за 3 мес", "Последний месяц"]:
                if col in main_df.columns: main_df[col] = main_df[col].apply(lambda x: f"{float(x) * 100:+.3f}%" if pd.notna(x) and x is not None else "")
            main_df.insert(0, "✓", False)

            edited_main = st.data_editor(main_df, width="stretch", hide_index=True, column_config={"✓": st.column_config.CheckboxColumn("✓", default=False, width="small")}, disabled=[c for c in main_df.columns if c != "✓"], key="aimoex_main_editor")
            selected_main = edited_main[edited_main["✓"] == True]

            if st.button("🗑 Удалить выбранные", key="aimoex_delete_inline"):
                if selected_main.empty: st.warning("Выберите строки")
                else:
                    names_to_del = selected_main["Наименование"].tolist()
                    tickers_to_del = [st.session_state.assets_to_imoex[st.session_state.assets_to_imoex["Наименование"] == n].iloc[0]["Тикер"] for n in names_to_del if not st.session_state.assets_to_imoex[st.session_state.assets_to_imoex["Наименование"] == n].empty]
                    if tickers_to_del:
                        st.session_state.assets_to_imoex = st.session_state.assets_to_imoex[~st.session_state.assets_to_imoex["Тикер"].isin(tickers_to_del)].reset_index(drop=True)
                        save_assets(st.session_state.assets_to_imoex)
                        st.session_state.history_cache = st.session_state.history_cache[~st.session_state.history_cache["Тикер"].isin(tickers_to_del)].reset_index(drop=True)
                        save_history_cache(st.session_state.history_cache)
                        st.success(f"Удалено активов: {len(tickers_to_del)}"); st.rerun()

            st.markdown("---")
            if "show_aimoex_detail" not in st.session_state: st.session_state.show_aimoex_detail = False
            if st.button("📂 Показать детализацию" if not st.session_state.show_aimoex_detail else "📁 Скрыть детализацию", key="toggle_aimoex_detail"):
                st.session_state.show_aimoex_detail = not st.session_state.show_aimoex_detail; st.rerun()
            if st.session_state.show_aimoex_detail:
                st.subheader("Детализация по месяцам")
                detail_rows, all_months = [], sorted(cache["Месяц"].dropna().unique().tolist())
                for _, asset in st.session_state.assets_to_imoex.iterrows():
                    ticker, name = asset["Тикер"], asset["Наименование"]
                    asset_cache, row_data = cache[cache["Тикер"] == ticker], {"Наименование": name}
                    for month in all_months:
                        match = asset_cache[asset_cache["Месяц"] == month]
                        row_data[month] = f"{float(match.iloc[0]['Разница']) * 100:+.2f}%" if not match.empty and pd.notna(match.iloc[0]["Разница"]) and match.iloc[0]["Разница"] is not None else ""
                    detail_rows.append(row_data)
                detail_df = pd.DataFrame(detail_rows)
                st.dataframe(detail_df, width="stretch", hide_index=True, height=400)