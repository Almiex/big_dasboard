# -*- coding: utf-8 -*-
"""
Аналитика работы врачей клиники (Streamlit app)

Пользователь загружает от 2 до 12 Excel-отчетов формата
«<Название клиники>. Отчет по докторам, ММ.ГГГГ».

В каждом файле используются:
  - лист «Отчет по докторам»: таблица врачей + блок расчетов
    (Использование мощностей, Достижение потенциала,
    Альтернативный потенциал, Разница с классическим, Вывод);
  - лист «Потенциал»: норма ставки, количество ставок, рабочие часы,
    стоимость часа, целевая загрузка, потенциал по специализациям.

Приложение само:
  - находит листы и строки заголовков в каждом файле;
  - вытаскивает НАЗВАНИЕ КЛИНИКИ и ПЕРИОД (ММ.ГГГГ) из шапки;
  - СОРТИРУЕТ отчеты хронологически, независимо от порядка загрузки;
  - НЕ смешивает клиники между собой (селектор одной клиники,
    однофамильцы разных клиник не склеиваются);
  - Пометку «(кво)» в специализациях НЕ удаляет (визуальный маркер);
  - считает производные метрики и строит дашборды:
    динамика клиники, потенциал и мощности, специализации,
    динамика врачей (выбор 1+ врачей и метрики), сравнение врачей,
    выгрузка данных.
"""

import re
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# --------------------------------------------------------------------------- #
# Настройки
# --------------------------------------------------------------------------- #
MAX_REPORTS = 12
MIN_REPORTS = 2

MONTHS_RU = {
    1: "Янв", 2: "Фев", 3: "Мар", 4: "Апр", 5: "Май", 6: "Июн",
    7: "Июл", 8: "Авг", 9: "Сен", 10: "Окт", 11: "Ноя", 12: "Дек",
}
MONTHS_RU_FULL = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель", 5: "Май", 6: "Июнь",
    7: "Июль", 8: "Август", 9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь",
}

# Алиазы заголовков основного отчета (нормализуются: lower + схлопывание пробелов)
HEADER_ALIASES = {
    "doctor":       ["фио врача", "врач", "фио"],
    "spec":         ["специализация"],
    "hours_plan":   ["рабочих часов по графику", "часы по графику", "рабочие часы"],
    "hours_patient":["время с пациентом", "часов с пациентом"],
    "revenue":      ["стоимость оказанных услуг", "стоимость услуг", "выручка", "оказанных услуг"],
    "visits":       ["кол-во посещений", "количество посещений", "посещения"],
    "patients":     ["кол-во пациентов", "количество пациентов", "пациенты"],
    "first_visits": ["кол-во первичных", "первичные"],
}

# Алиасы заголовков листа «Потенциал»
POT_ALIASES = {
    "spec":        ["специализация"],
    "rate_norm":   ["норма ставки"],
    "rates":       ["количество ставок", "кол-во ставок"],
    "hours_total": ["рабочих часов всего", "часов всего"],
    "cost_hour":   ["стоимость часа"],
    "target_load": ["целевая загрузка"],
    "potential":   ["потенциал"],
}

# Порядок важен: более специфичные специализации проверяются раньше
# generic («хирург», «терапевт»), чтобы «Онколог, маммолог, хирург»
# уходил в Онкологию, а не в Хирургию.
SPEC_RULES = [
    (("травмат", "ортопед"), "Травматология-ортопедия"),
    (("акушер", "гинеколог"), "Гинекология"),
    (("ультразвук", "уздг", "узд", "узи", "функциональн"), "УЗИ / ФД"),
    (("онколог", "маммолог"), "Онкология"),
    (("аллерг",), "Аллергология"),
    (("дерматовенеролог",), "Дерматовенерология"),
    (("дерматолог",), "Дерматология"),
    (("кардиолог",), "Кардиология"),
    (("невролог",), "Неврология"),
    (("эндокринолог",), "Эндокринология"),
    (("гастро",), "Гастроэнтерология"),
    (("уролог",), "Урология"),
    (("отоларинголог", "лор"), "ЛОР"),
    (("колопроктолог", "проктолог"), "Проктология"),
    (("флеболог",), "Флебология"),
    (("пульмонолог", "фтизиатр"), "Пульмонология"),
    (("психиатр", "психотерапевт", "психолог"), "Психология"),
    (("мануальн",), "Мануальная терапия"),
    (("хирург",), "Хирургия"),
    (("терапевт",), "Терапия"),
]

def group_spec(raw: str) -> str:
    s = norm(raw)
    if not s or s == "nan":
        return "Без специализации"
    for keys, group in SPEC_RULES:
        if any(k in s for k in keys):
            return group
    return re.sub(r"\(?кво\)?", "", str(raw)).strip() or "Без специализации"


# --------------------------------------------------------------------------- #
# Утилиты нормализации
# --------------------------------------------------------------------------- #
def norm(x) -> str:
    return re.sub(r"\s+", " ", str(x)).strip().lower()


def parse_period_from_text(text: str):
    """Ищет 'ММ.ГГГГ' (или 'М.ГГГГ') и возвращает (year, month) или None."""
    m = re.search(r"(\d{1,2})\s*[.\-/]\s*(\d{4})", str(text))
    if m:
        month, year = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            return year, month
    return None


def parse_clinic_from_text(text: str):
    """Из шапки «ООО "Клиника в Северном". Отчет по докторам, 09.2026»
    вытаскивает название клиники — всё, что до «Отчет по докторам».
    Регистр и кавычки сохраняются."""
    m = re.search(r"^(.*?)\s*[.,]?\s*отчет\s+по\s+докторам", str(text), re.IGNORECASE)
    if m:
        clinic = m.group(1).strip(" .,")
        return clinic or None
    return None


def open_excel(file_bytes: bytes, file_name: str):
    """Открывает книгу: старый формат .xls — через xlrd, новый .xlsx — через openpyxl."""
    suffix = file_name.lower().rsplit(".", 1)[-1]
    if suffix == "xls":
        try:
            return pd.ExcelFile(BytesIO(file_bytes), engine="xlrd")
        except ImportError:
            raise ValueError(
                f"Файл {file_name} — старый формат .xls: нужен пакет xlrd "
                f"(уже добавлен в requirements.txt, перезапустите приложение)"
            )
    return pd.ExcelFile(BytesIO(file_bytes))


# --------------------------------------------------------------------------- #
# Парсинг основного листа «Отчет по докторам»
# --------------------------------------------------------------------------- #
def find_header_row(raw: pd.DataFrame, aliases: dict):
    """Возвращает (row_idx, {canonical_name: col_idx}) для строки заголовка."""
    for i in range(min(len(raw), 15)):
        row = raw.iloc[i]
        normed = [norm(v) for v in row.tolist()]
        if any(v in ("№", "n", "no") for v in normed) and any("фио" in v for v in normed):
            mapping = {}
            for col, v in enumerate(normed):
                for canon, al in aliases.items():
                    if canon in mapping:
                        continue
                    if v in al or any(v.startswith(a) for a in al):
                        mapping[canon] = col
                        break
            if "doctor" in mapping and "revenue" in mapping:
                return i, mapping
    return None, None


def find_period_and_clinic(raw: pd.DataFrame, fallback_name: str):
    """Ищет период и клинику в первых строках листа; fallback — имя файла."""
    for i in range(min(len(raw), 5)):
        for v in raw.iloc[i].tolist():
            period = parse_period_from_text(v)
            clinic = parse_clinic_from_text(v)
            if period or clinic:
                return period, clinic
    return parse_period_from_text(fallback_name), None


BOTTOM_KEYS = [
    ("capacity_use", ("исп мощност", "использование мощност")),
    ("pot_achieved", ("дост потенциала", "достижение потенциала")),
    ("alt_potential", ("альтернативный потенциал",)),
    ("diff_classic", ("разница с классическим", "разница")),
    ("conclusion", ("вывод",)),
]


def _find_value_nearby(region: pd.DataFrame, r: int, c: int):
    """Ищет число рядом с подписью: справа, сверху или снизу.
    В отчетах КВО значения лежат строкой ВЫШЕ подписи, в других — справа/снизу."""
    row = region.iloc[r]
    for cc in range(c + 1, min(region.shape[1], c + 9)):
        v = pd.to_numeric(row.iloc[cc], errors="coerce")
        if pd.notna(v):
            return float(v)
    if r - 1 >= 0:
        v = pd.to_numeric(region.iloc[r - 1, c], errors="coerce")
        if pd.notna(v):
            return float(v)
    if r + 1 < region.shape[0]:
        v = pd.to_numeric(region.iloc[r + 1, c], errors="coerce")
        if pd.notna(v):
            return float(v)
    return None


def _find_text_nearby(region: pd.DataFrame, r: int, c: int):
    """Ищет текст вывода: непустая текстовая ячейка справа или под меткой."""
    row = region.iloc[r]
    for cc in range(c + 1, min(region.shape[1], c + 9)):
        v = str(row.iloc[cc]).strip()
        if v and v.lower() != "nan":
            return v
    if r + 1 < region.shape[0]:
        v = str(region.iloc[r + 1, c]).strip()
        if v and v.lower() != "nan":
            return v
    return None


def parse_bottom_metrics(raw: pd.DataFrame, start_row: int, meta: dict):
    """Парсит блок расчетов под таблицей врачей: Использование мощностей,
    Достижение потенциала, Альтернативный потенциал, Разница, Вывод."""
    region = raw.iloc[start_row: start_row + 20]
    out = {}
    for r in range(region.shape[0]):
        for c in range(region.shape[1]):
            cell = norm(region.iloc[r, c])
            if not cell:
                continue
            for key, aliases in BOTTOM_KEYS:
                if key in out:
                    continue
                if any(a in cell for a in aliases):
                    if key == "conclusion":
                        val = _find_text_nearby(region, r, c)
                    else:
                        val = _find_value_nearby(region, r, c)
                    if val is not None:
                        out[key] = val
    out.update(meta)
    return out


def parse_report(file_bytes: bytes, file_name: str):
    """
    Парсит один Excel-файл: лист «Отчет по докторам» (таблица врачей +
    блок расчетов внизу) и лист «Потенциал». Возвращает (doctors_df, bottom_dict).
    """
    xls = open_excel(file_bytes, file_name)

    # 1) лист с нужной шапкой, иначе второй лист, иначе первый
    target = None
    for sh in xls.sheet_names:
        head = pd.read_excel(xls, sheet_name=sh, header=None, nrows=5)
        if any("отчет по докторам" in norm(v) for v in head.values.flatten()):
            target = sh
            break
    if target is None:
        target = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]

    raw = pd.read_excel(xls, sheet_name=target, header=None)

    # 2) период и клиника
    period, clinic = find_period_and_clinic(raw, file_name)
    if period is None:
        raise ValueError(f"Не удалось определить период (ММ.ГГГГ) в файле: {file_name}")
    year, month = period
    if not clinic:
        clinic = re.sub(r"\.(xlsx|xls)$", "", file_name, flags=re.IGNORECASE)

    meta = dict(year=year, month=month,
                period=pd.Period(freq="M", year=year, month=month),
                month_label=f"{MONTHS_RU[month]} {year}",
                month_full=f"{MONTHS_RU_FULL[month]} {year}",
                clinic=clinic, source_file=file_name)

    # 3) строка заголовка
    h_row, mapping = find_header_row(raw, HEADER_ALIASES)
    if h_row is None:
        raise ValueError(f"Не найдена строка заголовка (№ / ФИО врача) в файле: {file_name}")

    # 4) строки данных
    records = []
    stop_row = len(raw)
    for i in range(h_row + 1, len(raw)):
        row = raw.iloc[i]
        joined = " ".join(norm(v) for v in row.tolist())
        if "итог" in joined:
            stop_row = i
            break  # «Итого по клинике» и ниже — не данные
        first = norm(row.iloc[0])
        if not re.fullmatch(r"\d+(\.0)?", first):
            continue
        rec = {}
        for canon, col in mapping.items():
            rec[canon] = row.iloc[col]
        records.append(rec)

    if not records:
        raise ValueError(f"Не найдены строки врачей в файле: {file_name}")

    df = pd.DataFrame(records)

    # 5) типы
    num_cols = ["hours_plan", "hours_patient", "revenue", "visits", "patients", "first_visits"]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        else:
            df[c] = np.nan
    df["doctor"] = df["doctor"].astype(str).str.strip()
    # ВАЖНО: пометку «(кво)» НЕ удаляем — по ней врачи визуально отделяются.
    # Только схлопываем переносы строк/лишние пробелы.
    df["spec_raw"] = df["spec"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    df["spec_group"] = df["spec_raw"].map(group_spec)

    # 6) производные метрики
    df["load"] = np.where(df["hours_plan"] > 0, df["hours_patient"] / df["hours_plan"], np.nan)
    df["cost_hour"] = np.where(df["hours_patient"] > 0, df["revenue"] / df["hours_patient"], np.nan)
    df["cost_visit"] = np.where(df["visits"] > 0, df["revenue"] / df["visits"], np.nan)
    df["revenue_per_patient"] = np.where(df["patients"] > 0, df["revenue"] / df["patients"], np.nan)
    df["visits_per_patient"] = np.where(df["patients"] > 0, df["visits"] / df["patients"], np.nan)
    df["first_share"] = np.where(df["patients"] > 0, df["first_visits"] / df["patients"], np.nan)

    # 7) служебные поля: клиника + период
    for k, v in meta.items():
        df[k] = v

    # 8) блок расчетов под таблицей (Исп мощности / Дост потенциала / ...)
    bottom = parse_bottom_metrics(raw, stop_row + 1, meta)

    return df, bottom


# --------------------------------------------------------------------------- #
# Парсинг листа «Потенциал»
# --------------------------------------------------------------------------- #
def parse_potential_sheet(xls, meta: dict):
    """Парсит лист «Потенциал»: ставки, часы, стоимость часа, целевая
    загрузка и потенциал по специализациям. Если листа нет — None."""
    target = None
    for sh in xls.sheet_names:
        if "потенциал" in norm(sh):
            target = sh
            break
    if target is None:
        for sh in xls.sheet_names:
            head = pd.read_excel(xls, sheet_name=sh, header=None, nrows=10)
            if any("норма ставки" in norm(v) for v in head.values.flatten()):
                target = sh
                break
    if target is None:
        return None

    raw = pd.read_excel(xls, sheet_name=target, header=None)
    h_row, mapping = find_header_row_potential(raw)
    if h_row is None:
        return None

    records = []
    for i in range(h_row + 1, len(raw)):
        row = raw.iloc[i]
        joined = " ".join(norm(v) for v in row.tolist())
        if "итог" in joined:
            break
        first = str(row.iloc[mapping["spec"]]).strip()
        if not first or first.lower() == "nan":
            continue
        rec = {}
        for canon, col in mapping.items():
            rec[canon] = row.iloc[col]
        records.append(rec)

    if not records:
        return None

    df = pd.DataFrame(records)
    for c in ["rate_norm", "rates", "hours_total", "cost_hour", "target_load", "potential"]:
        if c not in df.columns:
            df[c] = np.nan
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # потенциал из файла; если пусто — считаем: часы * целевая загрузка * стоимость часа
    df["potential_use"] = df["potential"].where(
        df["potential"].notna() & (df["potential"] > 0),
        df["hours_total"] * df["target_load"] * df["cost_hour"],
    )
    df["spec"] = df["spec"].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()
    df["spec_group"] = df["spec"].map(group_spec)
    for k, v in meta.items():
        df[k] = v
    return df


def find_header_row_potential(raw: pd.DataFrame):
    for i in range(min(len(raw), 15)):
        row = raw.iloc[i]
        normed = [norm(v) for v in row.tolist()]
        if any("специализация" in v for v in normed) and any("норма ставки" in v for v in normed):
            mapping = {}
            for col, v in enumerate(normed):
                for canon, al in POT_ALIASES.items():
                    if canon in mapping:
                        continue
                    if v in al or any(v.startswith(a) for a in al):
                        mapping[canon] = col
                        break
            if "spec" in mapping and "hours_total" in mapping:
                return i, mapping
    return None, None


@st.cache_data(show_spinner=False)
def load_all_reports(files_bytes: list, files_names: list):
    doctor_frames, pot_frames, bottom_rows, errors = [], [], [], []
    for b, n in zip(files_bytes, files_names):
        try:
            xls = open_excel(b, n)
            # мета берем из основного листа через parse_report
            d, bottom = parse_report(b, n)
            doctor_frames.append(d)
            bottom_rows.append(bottom)
            pot = parse_potential_sheet(xls, {k: d[k].iloc[0] for k in
                                              ["year", "month", "period", "month_label",
                                               "month_full", "clinic", "source_file"]})
            if pot is not None:
                pot_frames.append(pot)
        except Exception as e:
            errors.append(str(e))
    if errors:
        for e in errors:
            st.warning(e)
    if not doctor_frames:
        st.error("Не удалось распознать ни одного отчета.")
        st.stop()

    df = pd.concat(doctor_frames, ignore_index=True)
    pot = pd.concat(pot_frames, ignore_index=True) if pot_frames else pd.DataFrame()
    kpi = pd.DataFrame(bottom_rows)
    # ХРОНОЛОГИЯ: сортируем по реальному периоду, а не по порядку загрузки
    df = df.sort_values(["period", "clinic", "doctor"]).reset_index(drop=True)
    if not pot.empty:
        pot = pot.sort_values(["period", "spec"]).reset_index(drop=True)
    if not kpi.empty:
        kpi = kpi.sort_values("period").reset_index(drop=True)
    return df, pot, kpi


# --------------------------------------------------------------------------- #
# Агрегаты для дашбордов
# --------------------------------------------------------------------------- #
def clinic_monthly(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["period", "month_label", "month_full"], observed=True).agg(
        revenue=("revenue", "sum"),
        hours_plan=("hours_plan", "sum"),
        hours_patient=("hours_patient", "sum"),
        visits=("visits", "sum"),
        patients=("patients", "sum"),
        first_visits=("first_visits", "sum"),
        doctors=("doctor", "nunique"),
    ).reset_index().sort_values("period")
    g["load"] = np.where(g["hours_plan"] > 0, g["hours_patient"] / g["hours_plan"], np.nan)
    g["cost_visit"] = np.where(g["visits"] > 0, g["revenue"] / g["visits"], np.nan)
    g["rev_per_patient"] = np.where(g["patients"] > 0, g["revenue"] / g["patients"], np.nan)
    g["first_share"] = np.where(g["patients"] > 0, g["first_visits"] / g["patients"], np.nan)
    g["visits_per_patient"] = np.where(g["patients"] > 0, g["visits"] / g["patients"], np.nan)
    return g


def spec_monthly(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby(["period", "month_label", "spec_group"], observed=True).agg(
        revenue=("revenue", "sum"),
        hours_plan=("hours_plan", "sum"),
        hours_patient=("hours_patient", "sum"),
        visits=("visits", "sum"),
        doctors=("doctor", "nunique"),
    ).reset_index()
    g["load"] = np.where(g["hours_plan"] > 0, g["hours_patient"] / g["hours_plan"], np.nan)
    return g.sort_values(["spec_group", "period"])


def pot_monthly(pot: pd.DataFrame) -> pd.DataFrame:
    """Потенциал клиники по месяцам из листа «Потенциал»."""
    g = pot.groupby(["period", "month_label", "month_full"], observed=True).agg(
        rates=("rates", "sum"),
        hours_total=("hours_total", "sum"),
        potential=("potential_use", "sum"),
    ).reset_index().sort_values("period")
    return g


def month_order(df: pd.DataFrame) -> list:
    return df.sort_values("period")["month_label"].drop_duplicates().tolist()


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
APP_VERSION = "2.6.0 (2026-11-03)"  # смотрите в подписи под заголовком — сверка версии деплоя

st.set_page_config(page_title="Аналитика врачей клиники", page_icon="🏥", layout="wide")

st.title("🏥 Аналитика работы врачей клиники")
st.sidebar.caption(f"🛠 Версия приложения: {APP_VERSION}")
st.caption(
    "Загрузите от 2 до 12 отчетов «<Клиника>. Отчет по докторам, ММ.ГГГГ» (.xlsx/.xls). "
    "Порядок загрузки не важен — период определяется из шапки каждого файла. "
    "Клиники между собой не сравниваются: каждая анализируется отдельно. "
    f"Версия: {APP_VERSION}"
)

with st.sidebar:
    st.header("📤 Загрузка отчетов")
    uploaded = st.file_uploader(
        "Excel-файлы отчетов",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="Ключевой лист — «Отчет по докторам». Можно заливать в любом порядке.",
    )
    if uploaded:
        names = [f.name for f in uploaded]
        dup = [n for n in set(names) if names.count(n) > 1]
        if dup:
            st.error(f"Дубли файлов: {', '.join(dup)}")
            st.stop()
        if not (MIN_REPORTS <= len(uploaded) <= MAX_REPORTS):
            st.warning(f"Загружено {len(uploaded)} файл(ов). Рекомендуется от {MIN_REPORTS} до {MAX_REPORTS}.")
        st.success(f"Файлов: {len(uploaded)}")

if not uploaded:
    st.info("👈 Загрузите отчеты в панели слева, чтобы увидеть дашборды.")
    st.stop()

with st.spinner("Читаю отчеты и выстраиваю хронологию..."):
    df, pot, kpi = load_all_reports([f.read() for f in uploaded], [f.name for f in uploaded])

# ------------------------- КЛИНИКА: без смешивания ------------------------- #
# По умолчанию открываем клинику с наибольшим числом периодов —
# там полноценная динамика, а не одиночный месяц.
clinics = list(df.groupby("clinic")["period"].nunique().sort_values(ascending=False).index)
with st.sidebar:
    st.header("🏥 Клиника")
    if len(clinics) > 1:
        sel_clinic = st.selectbox(
            "Клиника (клиники не сравниваются между собой)",
            clinics,
            help="Выберите, какую клинику анализируем. Разные клиники — разные врачи и разная динамика.",
        )
    else:
        sel_clinic = clinics[0]
        st.success(f"**{sel_clinic}**")

df = df[df["clinic"] == sel_clinic].copy()
pot = pot[pot["clinic"] == sel_clinic].copy() if not pot.empty else pot
kpi = kpi[kpi["clinic"] == sel_clinic].copy() if not kpi.empty else kpi
st.subheader(f"{sel_clinic}")

periods = df.sort_values("period")["period"].unique()
if len(periods) < 2:
    st.warning("По выбранной клинике распознан только один период — динамика будет ограничена.")
else:
    _pmin = df.loc[df["period"].idxmin(), "month_full"]
    _pmax = df.loc[df["period"].idxmax(), "month_full"]
    st.caption(f"Период анализа: **{_pmin} — {_pmax}**")

# ------------------------------ фильтры ------------------------------------ #
with st.sidebar:
    st.header("🔎 Фильтры")
    all_months = month_order(df)
    sel_months = st.multiselect("Периоды", all_months, default=all_months,
                                key=f"f_months_{sel_clinic}")
    sel_months = [m for m in sel_months if m in all_months] or all_months
    df = df[df["month_label"].isin(sel_months)]
    pot = pot[pot["month_label"].isin(sel_months)] if not pot.empty else pot
    kpi = kpi[kpi["month_label"].isin(sel_months)] if not kpi.empty else kpi
    specs = sorted(df["spec_group"].dropna().unique())
    sel_specs = st.multiselect("Специализации", specs, default=specs,
                               key=f"f_specs_{sel_clinic}")
    sel_specs = [s for s in sel_specs if s in specs] or specs
    df = df[df["spec_group"].isin(sel_specs)]
    pot = pot[pot["spec_group"].isin(sel_specs)] if not pot.empty else pot
    if df.empty:
        st.error("Нет данных после фильтрации.")
        st.stop()

clin = clinic_monthly(df)
spec = spec_monthly(df)
months = month_order(df)

PLOTLY = dict(template="plotly_white", height=420)


def kpi_delta(cur, old):
    if old and pd.notna(old) and old != 0:
        return (cur - old) / old
    return None


def val_fill(values, low, high):
    """Градиентная заливка столбцов по значению.
    Нижний цвет low подбираем контрастным фону, чтобы маленькие
    значения не сливались с белым фоном."""
    return dict(color=list(values), colorscale=[[0.0, low], [1.0, high]], showscale=False)

# ================================ ВКЛАДКИ =================================== #
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📈 Динамика клиники",
    "⚡ Потенциал и мощности",
    "🩺 Специализации",
    "👨‍⚕️ Динамика врачей",
    "🏆 Сравнение врачей",
    "📋 Данные",
])

# ------------------------------ ТАБ 1 -------------------------------------- #

def _render_tab1():
        last, prev = clin.iloc[-1], clin.iloc[-2] if len(clin) > 1 else None

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Выручка", f"{last['revenue']:,.0f} ₽".replace(",", " "),
                  None if prev is None else f"{kpi_delta(last['revenue'], prev['revenue']):+.1%}")
        c2.metric("Загрузка клиники", f"{last['load']:.1%}",
                  None if prev is None else f"{(last['load'] - prev['load']):+.1%}")
        c3.metric("Посещения", f"{int(last['visits'])}",
                  None if prev is None else f"{kpi_delta(last['visits'], prev['visits']):+.1%}")
        c4.metric("Пациенты", f"{int(last['patients'])}",
                  None if prev is None else f"{kpi_delta(last['patients'], prev['patients']):+.1%}")
        c5.metric("Доля первичных", f"{last['first_share']:.1%}",
                  None if prev is None else f"{(last['first_share'] - prev['first_share']):+.1%}")

        # --- Выручка: градиентная заливка по значению + подписи с дельтой ---
        rev = clin["revenue"]
        deltas = rev.pct_change()
        texts = []
        for v, d in zip(rev, deltas):
            t = f"{v / 1e6:,.2f} млн ₽".replace(",", " ")
            if pd.notna(d):
                color = "#2E7D32" if d >= 0 else "#C62828"
                t += f"<br><span style='font-size:12px; color:{color};'>{d:+.1%} к пред. месяцу</span>"
            texts.append(t)
        fig = go.Figure()
        fig.add_bar(x=clin["month_label"], y=rev, name="Выручка",
                    marker=val_fill(rev, "#C6DBEF", "#08306B"),
                    text=texts, textposition="outside")
        fig.update_layout(title="Выручка клиники по месяцам (заливка темнеет с ростом суммы)",
                          xaxis={"categoryorder": "array", "categoryarray": months},
                          yaxis=dict(title="Выручка, ₽"), showlegend=False, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

        # --- Загрузка: отдельный график с подписями ---
        fig = go.Figure()
        fig.add_bar(x=clin["month_label"], y=clin["load"], name="Загрузка",
                    marker=val_fill(clin["load"], "#FD8D3C", "#7F2704"),
                    text=[f"{v:.1%}" for v in clin["load"]], textposition="outside")
        fig.update_layout(title="Загрузка клиники по месяцам (время с пациентом / часы по графику)",
                          xaxis={"categoryorder": "array", "categoryarray": months},
                          yaxis=dict(title="Загрузка", tickformat=".0%"), showlegend=False, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            fig = go.Figure()
            fig.add_bar(x=clin["month_label"], y=clin["hours_plan"], name="Часы по графику", marker_color="#9ECAE1")
            fig.add_bar(x=clin["month_label"], y=clin["hours_patient"],
                        name="Время с пациентом (факт. приемов)", marker_color="#08519C")
            fig.update_layout(title="Часы по графику vs время с пациентом", barmode="group",
                              xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = go.Figure()
            fig.add_bar(x=clin["month_label"], y=clin["visits"], name="Посещения", marker_color="#4C78A8")
            fig.add_bar(x=clin["month_label"], y=clin["patients"], name="Пациенты", marker_color="#72B7B2")
            fig.add_bar(x=clin["month_label"], y=clin["first_visits"], name="Первичные", marker_color="#F58518")
            fig.update_layout(title="Посещения / пациенты / первичные", barmode="group",
                              yaxis=dict(title="кол-во"),
                              xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)
            fig = px.line(clin, x="month_label", y="visits_per_patient", markers=True,
                          text="visits_per_patient", title="Посещений на 1 пациента (отдельно)",
                          labels={"visits_per_patient": "визитов на пациента", "month_label": ""})
            fig.update_traces(texttemplate="%{text:.2f}", textposition="top center")
            fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)
            st.caption(
                "Логика: **Посещения** — все визиты за месяц; **Пациенты** — уникальные люди "
                "(визитов всегда ≥ пациентов: один человек может прийти несколько раз); "
                "**Первичные** — первые обращения в клинику за месяц."
            )

        c1, c2 = st.columns(2)
        with c1:
            fig = px.line(clin, x="month_label", y="cost_visit", markers=True, text="cost_visit",
                          title="Средняя стоимость посещения", labels={"cost_visit": "₽ / посещение", "month_label": ""})
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="top center")
            fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            fig = px.line(clin, x="month_label", y="rev_per_patient", markers=True, text="rev_per_patient",
                          title="Выручка на пациента", labels={"rev_per_patient": "₽ / пациент", "month_label": ""})
            fig.update_traces(texttemplate="%{text:,.0f}", textposition="top center")
            fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)

    # ------------------------------ ТАБ 2: ПОТЕНЦИАЛ --------------------------- #

def _render_tab2():
        if pot.empty:
            st.info("В загруженных отчетах нет листа «Потенциал» — динамика потенциала недоступна.")
        else:
            pm = pot_monthly(pot).merge(
                clin[["period", "revenue", "hours_patient"]], on="period", how="left")
            pm["capacity_use_calc"] = np.where(pm["hours_total"] > 0,
                                               pm["hours_patient"] / pm["hours_total"], np.nan)
            pm["achieve_calc"] = np.where(pm["potential"] > 0, pm["revenue"] / pm["potential"], np.nan)

            if not kpi.empty:
                kpi_cols = [c for c in ["period", "capacity_use", "pot_achieved",
                                        "alt_potential", "diff_classic"] if c in kpi.columns]
                if len(kpi_cols) > 1:
                    pm = pm.merge(kpi[kpi_cols].drop_duplicates("period"), on="period", how="left")

            last = pm.iloc[-1]
            prev = pm.iloc[-2] if len(pm) > 1 else None

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Потенциал (план)", f"{last['potential']:,.0f} ₽".replace(",", " "),
                      None if prev is None else f"{kpi_delta(last['potential'], prev['potential']):+.1%}")
            c2.metric("Выручка факт", f"{last['revenue']:,.0f} ₽".replace(",", " "),
                      None if prev is None else f"{kpi_delta(last['revenue'], prev['revenue']):+.1%}")
            c3.metric("Достижение потенциала", f"{last['achieve_calc']:.1%}",
                      None if prev is None else f"{(last['achieve_calc'] - prev['achieve_calc']):+.1%}")
            c4.metric("Использование мощностей", f"{last['capacity_use_calc']:.1%}",
                      None if prev is None else f"{(last['capacity_use_calc'] - prev['capacity_use_calc']):+.1%}")

            # Потенциал vs факт
            fig = go.Figure()
            fig.add_bar(x=pm["month_label"], y=pm["potential"], name="Потенциал (план)",
                        marker=val_fill(pm["potential"], "#74C476", "#00441B"))
            fig.add_bar(x=pm["month_label"], y=pm["revenue"], name="Выручка факт",
                        marker=val_fill(pm["revenue"], "#C6DBEF", "#08306B"))
            fig.update_layout(title="Потенциал vs фактическая выручка по месяцам", barmode="group",
                              xaxis={"categoryorder": "array", "categoryarray": months},
                              yaxis=dict(title="₽"), **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1:
                fig = go.Figure()
                fig.add_bar(x=pm["month_label"], y=pm["achieve_calc"], name="Достижение потенциала",
                            marker=val_fill(pm["achieve_calc"], "#9E9AC8", "#3F007D"),
                            text=[f"{v:.1%}" for v in pm["achieve_calc"]], textposition="outside")
                if "pot_achieved" in pm.columns and pm["pot_achieved"].notna().any():
                    fig.add_scatter(x=pm["month_label"], y=pm["pot_achieved"], name="Из отчета",
                                    mode="lines+markers", line=dict(color="#F58518", dash="dash"))
                fig.update_layout(title="Достижение потенциала (выручка / потенциал)",
                                  xaxis={"categoryorder": "array", "categoryarray": months},
                                  yaxis=dict(tickformat=".0%"), **PLOTLY)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig = go.Figure()
                fig.add_bar(x=pm["month_label"], y=pm["capacity_use_calc"], name="Использование мощностей",
                            marker=val_fill(pm["capacity_use_calc"], "#FD8D3C", "#7F2704"),
                            text=[f"{v:.1%}" for v in pm["capacity_use_calc"]], textposition="outside")
                if "capacity_use" in pm.columns and pm["capacity_use"].notna().any():
                    fig.add_scatter(x=pm["month_label"], y=pm["capacity_use"], name="Из отчета",
                                    mode="lines+markers", line=dict(color="#08519C", dash="dash"))
                fig.update_layout(title="Использование мощностей (время с пациентом / рабочие часы)",
                                  xaxis={"categoryorder": "array", "categoryarray": months},
                                  yaxis=dict(tickformat=".0%"), **PLOTLY)
                st.plotly_chart(fig, use_container_width=True)

            if "alt_potential" in pm.columns and pm["alt_potential"].notna().any():
                fig = go.Figure()
                fig.add_bar(x=pm["month_label"], y=pm["potential"], name="Классический потенциал",
                            marker_color="#74C476")
                fig.add_bar(x=pm["month_label"], y=pm["alt_potential"], name="Альтернативный потенциал",
                            marker_color="#F58518")
                if "diff_classic" in pm.columns and pm["diff_classic"].notna().any():
                    fig.add_scatter(x=pm["month_label"], y=pm["diff_classic"], name="Разница с классическим",
                                    mode="lines+markers", yaxis="y2", line=dict(color="#C62828", width=3))
                fig.update_layout(title="Альтернативный потенциал",
                                  xaxis={"categoryorder": "array", "categoryarray": months},
                                  yaxis=dict(title="₽"), yaxis2=dict(title="Разница, ₽", overlaying="y", side="right"),
                                  **PLOTLY)
                st.plotly_chart(fig, use_container_width=True)

            conclusions = (kpi[kpi["conclusion"].notna()][["month_label", "conclusion"]]
                           if (not kpi.empty and "conclusion" in kpi.columns) else pd.DataFrame())
            if not conclusions.empty:
                st.subheader("Выводы из отчетов")
                for _, r in conclusions.iterrows():
                    st.markdown(f"**{r['month_label']}:** {r['conclusion']}")

            st.subheader("Сводка по месяцам")
            t = pm.copy()
            t["potential"] = t["potential"].map("{:,.0f}".format)
            t["revenue"] = t["revenue"].map("{:,.0f}".format)
            t["achieve_calc"] = t["achieve_calc"].map("{:.1%}".format)
            t["capacity_use_calc"] = t["capacity_use_calc"].map("{:.1%}".format)
            t = t[["month_full", "rates", "hours_total", "potential", "revenue",
                   "achieve_calc", "capacity_use_calc"]]
            t.columns = ["Месяц", "Ставок", "Рабочих часов (план)", "Потенциал, ₽",
                         "Выручка факт, ₽", "Достижение потенциала", "Использование мощностей"]
            st.dataframe(t, use_container_width=True, hide_index=True)

            st.subheader("Потенциал по специализациям (из листа «Потенциал»)")
            pot_spec = pot.groupby(["month_label", "spec_group"], as_index=False).agg(
                hours_total=("hours_total", "sum"), potential=("potential_use", "sum"))
            pot_spec = pot_spec.merge(
                spec[["month_label", "spec_group", "revenue"]], on=["month_label", "spec_group"], how="left")
            pot_spec["achieve"] = np.where(pot_spec["potential"] > 0,
                                           pot_spec["revenue"] / pot_spec["potential"], np.nan)
            pot_pivot = pot_spec.pivot_table(index="spec_group", columns="month_label",
                                             values="achieve", aggfunc="sum")
            pot_pivot = pot_pivot.reindex(pot_spec.groupby("spec_group")["potential"]
                                          .sum().sort_values(ascending=False).index)
            fig = px.imshow(pot_pivot, text_auto=".0%", aspect="auto", color_continuous_scale="RdYlGn",
                            title="Достижение потенциала по специализациям и месяцам")
            fig.update_layout(height=max(400, 20 * len(pot_pivot)))
            st.plotly_chart(fig, use_container_width=True)

    # ------------------------------ ТАБ 3: СПЕЦИАЛИЗАЦИИ ---------------------- #

def _render_tab3():
        top_specs = spec.groupby("spec_group")["revenue"].sum().sort_values(ascending=False).index.tolist()
        sel_top = st.multiselect("Специализации на графике", top_specs, default=top_specs[:8],
                                 key=f"tab3_specs_{sel_clinic}")
        sel_top = [s for s in sel_top if s in top_specs] or top_specs[:8]
        d2 = spec[spec["spec_group"].isin(sel_top)]
        if d2.empty:
            st.info("Нет данных по выбранным специализациям.")
            return

        fig = px.bar(d2, x="month_label", y="revenue", color="spec_group",
                     title="Выручка по специализациям (структура по месяцам)",
                     labels={"revenue": "Выручка, ₽", "month_label": "", "spec_group": "Специализация"})
        fig.update_layout(barmode="stack", xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

        fig = px.line(d2, x="month_label", y="load", color="spec_group", markers=True,
                      title="Загрузка по специализациям",
                      labels={"load": "Загрузка", "month_label": "", "spec_group": "Специализация"})
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

        if not pot.empty:
            pot_spec2 = pot.groupby(["month_label", "spec_group"], as_index=False).agg(
                potential=("potential_use", "sum"))
            pot_spec2 = pot_spec2.merge(spec[["month_label", "spec_group", "revenue"]],
                                        on=["month_label", "spec_group"], how="inner")
            pot_spec2["achieve"] = np.where(pot_spec2["potential"] > 0,
                                            pot_spec2["revenue"] / pot_spec2["potential"], np.nan)
            pot_spec2 = pot_spec2[pot_spec2["spec_group"].isin(sel_top)]
            fig = px.line(pot_spec2, x="month_label", y="achieve", color="spec_group", markers=True,
                          title="Достижение потенциала по специализациям (выручка / потенциал)",
                          labels={"achieve": "Достижение потенциала", "month_label": "", "spec_group": "Специализация"})
            fig.update_yaxes(tickformat=".0%")
            fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
            st.plotly_chart(fig, use_container_width=True)

        spec_pivot = spec.pivot_table(index="spec_group", columns="month_label",
                                      values="revenue", aggfunc="sum").reindex(top_specs)
        st.subheader("Выручка по специализациям, ₽")
        st.dataframe(spec_pivot.style.format("{:,.0f}", na_rep="—"))

    # ------------------------------ ТАБ 4: ДИНАМИКА ВРАЧЕЙ --------------------- #

def _render_tab4():
        st.subheader("Динамика врачей: отдельный график по каждой метрике")
        c1, c2 = st.columns([1, 2])
        with c1:
            spec_list = sorted(df["spec_group"].dropna().unique())
            sel_spec = st.selectbox("Специализация", spec_list, key=f"tab4_spec_{sel_clinic}")
        docs_in_spec = sorted(df.loc[df["spec_group"] == sel_spec, "doctor"].unique())
        with c2:
            sel_docs = st.multiselect("Врачи этой специализации", docs_in_spec, default=docs_in_spec,
                                      key=f"tab4_docs_{sel_clinic}_{sel_spec}")
        # защита от устаревшего состояния виджетов (при смене файлов/клиники):
        # выбранные врачи, которых нет в списке, отбрасываем; если все отброшены — берем всех
        sel_docs = [d for d in sel_docs if d in docs_in_spec] or docs_in_spec
        d4 = df[(df["spec_group"] == sel_spec) & (df["doctor"].isin(sel_docs))].sort_values("period")
        if d4.empty:
            st.info("Нет данных по выбранным врачам — проверьте выбор специализации и врачей.")
            return

        metric_charts = [
            ("Рабочих часов по графику", "hours_plan", "ч", False),
            ("Время с пациентом (факт. приемов)", "hours_patient", "ч", False),
            ("Стоимость оказанных услуг", "revenue", "₽", False),
            ("Загрузка", "load", "", True),
            ("Стоимость фактического часа", "cost_hour", "₽", False),
            ("Посещения", "visits", "", False),
            ("Пациенты", "patients", "", False),
            ("Первичные", "first_visits", "", False),
            ("Доля первичных", "first_share", "", True),
            ("Стоимость посещения", "cost_visit", "₽", False),
        ]
        single_month = len(months) < 2
        if single_month:
            st.info("По выбранной клинике загружен только один месяц — графики показаны "
                    "столбцами. Загрузите отчеты за несколько месяцев, чтобы увидеть динамику.")
        cols = st.columns(2)
        for i, (title, colname, unit, is_pct) in enumerate(metric_charts):
            with cols[i % 2]:
                ylabel = f"{title}{', ' + unit if unit else ''}"
                if single_month:
                    # один месяц: столбцы по врачам вместо линий из одной точки
                    fig = px.bar(d4, x="doctor", y=colname, color="doctor",
                                 title=f"{title} — {months[0]}",
                                 labels={colname: ylabel, "doctor": "Врач"})
                else:
                    fig = px.line(d4, x="month_label", y=colname, color="doctor", markers=True,
                                  title=title,
                                  labels={colname: ylabel, "month_label": "", "doctor": "Врач"})
                if is_pct:
                    fig.update_yaxes(tickformat=".0%")
                if not single_month:
                    fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months})
                fig.update_layout(height=330, template="plotly_white")
                st.plotly_chart(fig, use_container_width=True)

        st.subheader("Таблица по выбранным врачам")
        show_cols = ["month_full", "doctor", "spec_raw", "hours_plan", "hours_patient",
                     "revenue", "load", "cost_hour", "cost_visit", "visits",
                     "patients", "first_visits", "first_share"]
        t4 = d4[show_cols].copy()
        t4.columns = ["Месяц", "Врач", "Специализация", "Часы по графику", "Время с пациентом",
                      "Выручка", "Загрузка", "₽/час", "₽/посещение", "Посещения",
                      "Пациенты", "Первичные", "Доля первичных"]
        st.dataframe(
            t4.style.format({"Выручка": "{:,.0f}", "Загрузка": "{:.1%}", "₽/час": "{:,.0f}",
                             "₽/посещение": "{:,.0f}", "Доля первичных": "{:.1%}",
                             "Часы по графику": "{:,.1f}", "Время с пациентом": "{:,.1f}"}),
            use_container_width=True, hide_index=True,
        )

    # ------------------------------ ТАБ 5: СРАВНЕНИЕ ВРАЧЕЙ ------------------- #

def _render_tab5():
        last_lbl = months[-1]
        d5 = df[df["month_label"] == last_lbl]

        c1, c2 = st.columns([1, 2])
        with c1:
            st.subheader(f"Топ-10 врачей по выручке — {last_lbl}")
            top10 = d5.nlargest(10, "revenue")[["doctor", "spec_group", "revenue"]]
            top10.columns = ["Врач", "Специализация", "Выручка"]
            st.dataframe(top10.style.format({"Выручка": "{:,.0f}"}), use_container_width=True, hide_index=True)

        with c2:
            fig = px.scatter(d5, x="load", y="revenue", color="spec_group", hover_name="doctor",
                             size="visits", title=f"Врачи {last_lbl}: загрузка × выручка (размер — посещения)",
                             labels={"load": "Загрузка", "revenue": "Выручка, ₽", "spec_group": "Специализация"})
            fig.update_xaxes(tickformat=".0%")
            fig.update_layout(**PLOTLY)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Матрица «врач × месяц» — загрузка")
        pivot_load = df.pivot_table(index="doctor", columns="month_label", values="load")
        rev_by_doc = df.groupby("doctor")["revenue"].sum().sort_values(ascending=False)
        pivot_load = pivot_load.reindex(rev_by_doc.index)
        fig = px.imshow(pivot_load, text_auto=".0%", aspect="auto", color_continuous_scale="Blues",
                        title="Загрузка врачей по месяцам")
        fig.update_layout(height=max(400, 18 * len(pivot_load)))
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Сводка по врачам за весь период")
        doc_sum = df.groupby(["doctor", "spec_group"], as_index=False).agg(
            revenue=("revenue", "sum"), hours_plan=("hours_plan", "sum"),
            hours_patient=("hours_patient", "sum"), visits=("visits", "sum"),
            patients=("patients", "sum"), first_visits=("first_visits", "sum"),
            months=("month_label", "nunique"),
        )
        doc_sum["load"] = np.where(doc_sum["hours_plan"] > 0, doc_sum["hours_patient"] / doc_sum["hours_plan"], np.nan)
        doc_sum = doc_sum.sort_values("revenue", ascending=False)
        doc_sum.columns = ["Врач", "Специализация", "Выручка", "Часы по графику", "Часы с пациентом",
                           "Посещения", "Пациенты", "Первичные", "Месяцев в отчетах", "Загрузка"]
        st.dataframe(
            doc_sum.style.format({"Выручка": "{:,.0f}", "Загрузка": "{:.1%}",
                                  "Часы по графику": "{:,.1f}", "Часы с пациентом": "{:,.1f}"}),
            use_container_width=True, hide_index=True,
        )

    # ------------------------------ ТАБ 6: ДАННЫЕ ------------------------------- #

def _render_tab6():
        st.subheader("Объединенная таблица (все отчеты, в хронологическом порядке)")
        out_cols = ["month_full", "source_file", "clinic", "doctor", "spec_raw", "spec_group",
                    "hours_plan", "hours_patient", "revenue", "visits", "patients", "first_visits",
                    "load", "cost_hour", "cost_visit", "revenue_per_patient", "first_share"]
        out = df[out_cols].copy()
        out.columns = ["Месяц", "Файл", "Клиника", "Врач", "Специализация (как в отчете)", "Группа",
                       "Часы по графику", "Время с пациентом", "Выручка", "Посещения",
                       "Пациенты", "Первичные", "Загрузка", "₽/час", "₽/посещение",
                       "₽/пациент", "Доля первичных"]

        c1, c2 = st.columns(2)
        csv = out.to_csv(index=False).encode("utf-8-sig")
        c1.download_button("⬇️ Скачать CSV (врачи)", csv, "doctors_analytics.csv", "text/csv")
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            out.to_excel(w, index=False, sheet_name="Данные")
            clin.rename(columns={"month_label": "Месяц"}).to_excel(w, index=False, sheet_name="Клиника по месяцам")
            if not pot.empty:
                pot_monthly(pot).rename(columns={"month_label": "Месяц"}).to_excel(
                    w, index=False, sheet_name="Потенциал по месяцам")
        c2.download_button("⬇️ Скачать Excel (все листы)", buf.getvalue(), "doctors_analytics.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.dataframe(out, use_container_width=True, hide_index=True)


# Изоляция ошибок: падение одной вкладки не ломает остальные,
# а ошибка показывается ПРЯМО ВНУТРИ проблемной вкладки.
for _t, _fn in zip((tab1, tab2, tab3, tab4, tab5, tab6),
                   (_render_tab1, _render_tab2, _render_tab3,
                    _render_tab4, _render_tab5, _render_tab6)):
    with _t:
        try:
            _fn()
        except Exception as _e:
            if type(_e).__name__ == "StopException":
                raise
            st.error(f"⚠️ Ошибка в этой вкладке: {type(_e).__name__}: {_e}")
            st.exception(_e)
