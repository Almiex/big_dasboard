# -*- coding: utf-8 -*-
"""
Аналитика работы врачей клиники (Streamlit app)

Пользователь загружает от 2 до 12 Excel-отчетов формата
«Отчет по докторам, ММ.ГГГГ» (ключевой лист — «Отчет по докторам»).
Приложение само:
  - находит лист и строку заголовка в каждом файле;
  - вытаскивает период отчета (ММ.ГГГГ) из шапки и СОРТИРУЕТ отчеты
    хронологически, независимо от порядка загрузки;
  - считает производные метрики (загрузка, стоимость часа и т.д.);
  - строит дашборды: динамика клиники, специализации, динамика врачей,
    сравнение врачей, выгрузка данных.
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

# Алиасы заголовков (нормализуются: lower + схлопывание пробелов)
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

# Схлопывание специализаций в аналитические группы
def group_spec(raw: str) -> str:
    s = norm(raw)
    if not s or s == "nan":
        return "Без специализации"
    if any(k in s for k in ("ультра", "уздг", "узи", "функциональн")):
        return "УЗИ / ФД"
    if "травмат" in s or "ортопед" in s:
        return "Травматология-ортопедия"
    if "терапевт" in s or "пульмонолог" in s:
        return "Терапия"
    if "дерматовенеролог" in s:
        return "Дерматовенерология"
    if "терапевт-дерматолог" in s:
        return "Дерматология"
    if "акушер" in s or "гинеколог" in s:
        return "Гинекология"
    if "хирург" in s:
        return "Хирургия"
    if "невролог" in s:
        return "Неврология"
    if "кардиолог" in s:
        return "Кардиология"
    if "уролог" in s:
        return "Урология"
    if "гастро" in s:
        return "Гастроэнтерология"
    if "эндокринолог" in s:
        return "Эндокринология"
    if "онколог" in s:
        return "Онкология"
    if "лор" in s or "отоларинголог" in s:
        return "ЛОР"
    if "проктолог" in s:
        return "Проктология"
    if "флеболог" in s:
        return "Флебология"
    if "психолог" in s or "психиатр" in s:
        return "Психология"
    return str(raw).strip()


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


# --------------------------------------------------------------------------- #
# Парсинг одного отчета
# --------------------------------------------------------------------------- #
def find_header_row(raw: pd.DataFrame):
    """Возвращает (row_idx, {canonical_name: col_idx}) для строки заголовка."""
    for i in range(min(len(raw), 15)):
        row = raw.iloc[i]
        normed = [norm(v) for v in row.tolist()]
        if any(v in ("№", "n", "no") for v in normed) and any("фио" in v for v in normed):
            mapping = {}
            for col, v in enumerate(normed):
                for canon, aliases in HEADER_ALIASES.items():
                    if canon in mapping:
                        continue
                    if v in aliases or any(v.startswith(a) for a in aliases):
                        mapping[canon] = col
                        break
            if "doctor" in mapping and "revenue" in mapping:
                return i, mapping
    return None, None


def find_period(raw: pd.DataFrame, fallback_name: str):
    """Ищет период отчета: сначала в первых строках листа, потом в имени файла."""
    for i in range(min(len(raw), 5)):
        for v in raw.iloc[i].tolist():
            p = parse_period_from_text(v)
            if p:
                return p
    return parse_period_from_text(fallback_name) or (None, None)


def parse_report(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    """
    Парсит один Excel-файл. Ищет лист, содержащий «Отчет по докторам»,
    в нем — строку заголовка, затем читает строки врачей до «Итого по клинике».
    """
    xls = pd.ExcelFile(BytesIO(file_bytes))
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

    # 2) период
    year, month = find_period(raw, file_name)
    if year is None:
        raise ValueError(f"Не удалось определить период (ММ.ГГГГ) в файле: {file_name}")

    # 3) строка заголовка
    h_row, mapping = find_header_row(raw)
    if h_row is None:
        raise ValueError(f"Не найдена строка заголовка (№ / ФИО врача) в файле: {file_name}")

    # 4) строки данных
    records = []
    for i in range(h_row + 1, len(raw)):
        row = raw.iloc[i]
        first = norm(row.iloc[0])
        if "итог" in " ".join(norm(v) for v in row.tolist()):
            break  # «Итого по клинике» и ниже — не данные
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
    df["spec_raw"] = df["spec"].astype(str).str.strip()
    df["spec_group"] = df["spec_raw"].map(group_spec)

    # 6) производные метрики
    df["load"] = np.where(df["hours_plan"] > 0, df["hours_patient"] / df["hours_plan"], np.nan)
    df["cost_hour"] = np.where(df["hours_patient"] > 0, df["revenue"] / df["hours_patient"], np.nan)
    df["cost_visit"] = np.where(df["visits"] > 0, df["revenue"] / df["visits"], np.nan)
    df["revenue_per_patient"] = np.where(df["patients"] > 0, df["revenue"] / df["patients"], np.nan)
    df["visits_per_patient"] = np.where(df["patients"] > 0, df["visits"] / df["patients"], np.nan)
    df["first_share"] = np.where(df["patients"] > 0, df["first_visits"] / df["patients"], np.nan)

    # 7) служебные поля периода
    df["year"], df["month"] = year, month
    df["period"] = pd.Period(freq="M", year=year, month=month)
    df["month_label"] = f"{MONTHS_RU[month]} {year}"
    df["month_full"] = f"{MONTHS_RU_FULL[month]} {year}"
    df["source_file"] = file_name
    return df


@st.cache_data(show_spinner=False)
def load_all_reports(files_bytes: list, files_names: list) -> pd.DataFrame:
    frames, errors = [], []
    for b, n in zip(files_bytes, files_names):
        try:
            frames.append(parse_report(b, n))
        except Exception as e:
            errors.append(str(e))
    if errors:
        for e in errors:
            st.warning(e)
    if not frames:
        st.error("Не удалось распознать ни одного отчета.")
        st.stop()
    df = pd.concat(frames, ignore_index=True)
    # ХРОНОЛОГИЯ: сортируем по реальному периоду, а не по порядку загрузки
    df = df.sort_values(["period", "doctor"]).reset_index(drop=True)
    return df


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


def month_order(df: pd.DataFrame) -> list:
    return df.sort_values("period")["month_label"].drop_duplicates().tolist()


# --------------------------------------------------------------------------- #
# UI
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="Аналитика врачей клиники", page_icon="🏥", layout="wide")

st.title("🏥 Аналитика работы врачей клиники")
st.caption(
    "Загрузите от 2 до 12 отчетов «Отчет по докторам, ММ.ГГГГ» (.xlsx). "
    "Порядок загрузки не важен — период определяется из шапки каждого файла "
    "и отчеты выстраиваются в хронологию автоматически."
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
    df = load_all_reports([f.read() for f in uploaded], [f.name for f in uploaded])

periods = df.sort_values("period")["period"].unique()
if len(periods) < 2:
    st.warning("Распознан только один период — динамика будет ограничена.")

# ------------------------------ фильтры ------------------------------------ #
with st.sidebar:
    st.header("🔎 Фильтры")
    all_months = month_order(df)
    sel_months = st.multiselect("Периоды", all_months, default=all_months)
    df = df[df["month_label"].isin(sel_months)]
    specs = sorted(df["spec_group"].dropna().unique())
    sel_specs = st.multiselect("Специализации", specs, default=specs)
    df = df[df["spec_group"].isin(sel_specs)]
    if df.empty:
        st.error("Нет данных после фильтрации.")
        st.stop()
    st.caption(f"Период анализа: **{df['month_full'].min()} — {df['month_full'].max()}**")

clin = clinic_monthly(df)
spec = spec_monthly(df)
months = month_order(df)
x_order = {"month_label": months}

PLOTLY = dict(template="plotly_white", height=420)

def bar_line(g, bar_col, line_col, title, y1_title, y2_title):
    fig = go.Figure()
    fig.add_bar(x=g["month_label"], y=g[bar_col], name=y1_title, marker_color="#4C78A8")
    fig.add_scatter(x=g["month_label"], y=g[line_col], name=y2_title, yaxis="y2",
                    mode="lines+markers", line=dict(color="#F58518", width=3))
    fig.update_layout(
        title=title, xaxis={"categoryorder": "array", "categoryarray": months},
        yaxis=dict(title=y1_title), yaxis2=dict(title=y2_title, overlaying="y", side="right"),
        **PLOTLY,
    )
    return fig

# ================================ ВКЛАДКИ =================================== #
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📈 Динамика клиники",
    "🩺 Специализации",
    "👨‍⚕️ Динамика врачей",
    "🏆 Сравнение врачей",
    "📋 Данные",
])

# ------------------------------ ТАБ 1 -------------------------------------- #
with tab1:
    last, prev = clin.iloc[-1], clin.iloc[-2] if len(clin) > 1 else None

    def kpi_delta(cur, old):
        if old and pd.notna(old) and old != 0:
            return (cur - old) / old
        return None

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

    st.plotly_chart(bar_line(clin, "revenue", "load", "Выручка и загрузка клиники по месяцам",
                             "Выручка, ₽", "Загрузка"), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure()
        fig.add_bar(x=clin["month_label"], y=clin["hours_plan"], name="Часы по графику", marker_color="#9ECAE1")
        fig.add_bar(x=clin["month_label"], y=clin["hours_patient"], name="Время с пациентом", marker_color="#08519C")
        fig.update_layout(title="Часы по графику vs время с пациентом", barmode="group",
                          xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = go.Figure()
        fig.add_bar(x=clin["month_label"], y=clin["visits"], name="Посещения", marker_color="#4C78A8")
        fig.add_bar(x=clin["month_label"], y=clin["patients"], name="Пациенты", marker_color="#72B7B2")
        fig.add_bar(x=clin["month_label"], y=clin["first_visits"], name="Первичные", marker_color="#F58518")
        fig.update_layout(title="Посещения / пациенты / первичные", barmode="group",
                          xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

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

# ------------------------------ ТАБ 2 -------------------------------------- #
with tab2:
    top_specs = spec.groupby("spec_group")["revenue"].sum().sort_values(ascending=False).index.tolist()
    sel_top = st.multiselect("Специализации на графике", top_specs, default=top_specs[:8])
    d2 = spec[spec["spec_group"].isin(sel_top)]

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

    spec_pivot = spec.pivot_table(index="spec_group", columns="month_label",
                                  values="revenue", aggfunc="sum").reindex(top_specs)
    st.subheader("Выручка по специализациям, ₽")
    st.dataframe(spec_pivot.style.format("{:,.0f}", na_rep="—"))

# ------------------------------ ТАБ 3 -------------------------------------- #
with tab3:
    doctors = sorted(df["doctor"].unique())
    sel_doc = st.multiselect("Врачи", doctors, default=doctors[: min(5, len(doctors))])
    d3 = df[df["doctor"].isin(sel_doc)].sort_values("period")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(d3, x="month_label", y="revenue", color="doctor", markers=True,
                      title="Выручка врачей по месяцам",
                      labels={"revenue": "₽", "month_label": "", "doctor": "Врач"})
        fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.line(d3, x="month_label", y="load", color="doctor", markers=True,
                      title="Загрузка врачей по месяцам",
                      labels={"load": "Загрузка", "month_label": "", "doctor": "Врач"})
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(d3, x="month_label", y="cost_hour", color="doctor", markers=True,
                      title="Стоимость фактического часа",
                      labels={"cost_hour": "₽ / час", "month_label": "", "doctor": "Врач"})
        fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.line(d3, x="month_label", y="first_share", color="doctor", markers=True,
                      title="Доля первичных пациентов",
                      labels={"first_share": "Доля первичных", "month_label": "", "doctor": "Врач"})
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(xaxis={"categoryorder": "array", "categoryarray": months}, **PLOTLY)
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Таблица по выбранным врачам")
    show_cols = ["month_full", "doctor", "spec_raw", "revenue", "load", "cost_hour",
                 "cost_visit", "visits", "patients", "first_share"]
    t3 = d3[show_cols].copy()
    t3.columns = ["Месяц", "Врач", "Специализация", "Выручка", "Загрузка",
                  "₽/час", "₽/посещение", "Посещения", "Пациенты", "Доля первичных"]
    st.dataframe(
        t3.style.format({"Выручка": "{:,.0f}", "Загрузка": "{:.1%}", "₽/час": "{:,.0f}",
                         "₽/посещение": "{:,.0f}", "Доля первичных": "{:.1%}"}),
        use_container_width=True, hide_index=True,
    )

# ------------------------------ ТАБ 4 -------------------------------------- #
with tab4:
    last_lbl = months[-1]
    d4 = df[df["month_label"] == last_lbl]

    c1, c2 = st.columns([1, 2])
    with c1:
        st.subheader(f"Топ-10 врачей по выручке — {last_lbl}")
        top10 = d4.nlargest(10, "revenue")[["doctor", "spec_group", "revenue"]]
        top10.columns = ["Врач", "Специализация", "Выручка"]
        st.dataframe(top10.style.format({"Выручка": "{:,.0f}"}), use_container_width=True, hide_index=True)

    with c2:
        fig = px.scatter(d4, x="load", y="revenue", color="spec_group", hover_name="doctor",
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
                              "Часы по графику": "{:,.0f}", "Часы с пациентом": "{:,.0f}"}),
        use_container_width=True, hide_index=True,
    )

# ------------------------------ ТАБ 5 -------------------------------------- #
with tab5:
    st.subheader("Объединенная таблица (все отчеты, в хронологическом порядке)")
    out_cols = ["month_full", "source_file", "doctor", "spec_raw", "spec_group", "hours_plan",
                "hours_patient", "revenue", "visits", "patients", "first_visits", "load",
                "cost_hour", "cost_visit", "revenue_per_patient", "first_share"]
    out = df[out_cols].copy()
    out.columns = ["Месяц", "Файл", "Врач", "Специализация (как в отчете)", "Группа",
                   "Часы по графику", "Время с пациентом", "Выручка", "Посещения",
                   "Пациенты", "Первичные", "Загрузка", "₽/час", "₽/посещение",
                   "₽/пациент", "Доля первичных"]

    c1, c2 = st.columns(2)
    csv = out.to_csv(index=False).encode("utf-8-sig")
    c1.download_button("⬇️ Скачать CSV", csv, "doctors_analytics.csv", "text/csv")
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        out.to_excel(w, index=False, sheet_name="Данные")
        clin.rename(columns={"month_label": "Месяц"}).to_excel(w, index=False, sheet_name="Клиника по месяцам")
    c2.download_button("⬇️ Скачать Excel", buf.getvalue(), "doctors_analytics.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(out, use_container_width=True, hide_index=True)
