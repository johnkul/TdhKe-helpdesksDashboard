from pathlib import Path
import html
import re

import altair as alt
import pandas as pd
import streamlit as st


# Page configuration
st.set_page_config(
    page_title="Tdh Kenya Helpdesk Data Dashboard",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Source path for the dashboard logo
LOGO_PATH = Path("./assets/tdh-logo.png")

# Source file and dashboard constants
DATA_FILE_PATH = Path("./data/HELPDESK_DashboardData_Tdh_Kenya_D2.xlsx")

PII_COLUMNS = [
    "staff_name",
    "information_seeker_name",
    "residence_neighborhood_compound_house",
    "information_seeker_phone",
    "alternative_phone",
    "information_seeker_individual_number",
    "information_seeker_ration_or_wristband_number",
    "gps_latitude",
    "gps_longitude",
]

AGE_GROUP_ORDER = [
    "0-5 Years",
    "6-11 Years",
    "12-17 Years",
    "18-35 Years",
    "36-49 Years",
    "50-64 Years",
    "65 Years and Above",
    "[Missing]",
]

CHILD_AGE_GROUPS = {
    "0-5 Years",
    "6-11 Years",
    "12-17 Years",
}

ADULT_AGE_GROUPS = {
    "18-35 Years",
    "36-49 Years",
    "50-64 Years",
    "65 Years and Above",
}

GENDER_ORDER = ["Girl", "Boy", "Woman", "Man", "Transgender", "[Missing]"]

GENDER_COLORS = {
    "Girl": "#7C3AED",
    "Boy": "#2563EB",
    "Woman": "#DB2777",
    "Man": "#059669",
    "Transgender": "#F59E0B",
    "[Missing]": "#9CA3AF",
}

FILTER_KEYS = [
    "camp_location_filter",
    "helpdesk_location_filter",
    "information_seeker_type_filter",
    "information_seeker_gender_filter",
    "age_group_filter",
    "request_category_filter",
]

CORE_RECORD_COLUMNS = [
    "record_id",
    "interview_date",
    "camp_location",
    "helpdesk_location",
    "information_seeker_type",
    "information_seeker_gender",
    "age_group",
    "derived_life_stage",
    "information_seeker_type_raw",
    "information_seeker_gender_raw",
    "type_age_correction_flag",
    "gender_age_correction_flag",
    "request_category",
    "referral_status",
    "follow_up_required_clean",
]


# Visual theme and custom Streamlit styling
def inject_theme_css():
    st.markdown(
        """
        <style>
        .stApp {
            background:
                linear-gradient(180deg, #EEF6F2 0%, #F7F4EC 280px, #F8FAFC 100%);
        }

        .block-container {
            padding-top: 1.4rem;
            padding-bottom: 2rem;
        }

        section[data-testid="stSidebar"] {
            background: #E7F0EB;
            border-right: 1px solid #CBD5D1;
        }

        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3 {
            color: #12312F;
        }

        div[data-testid="stMetric"] {
            background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
            border: 1px solid #D8E2DC;
            border-left: 5px solid #2F7D69;
            border-radius: 8px;
            padding: 15px 16px;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
        }

        div[data-testid="stMetric"] label {
            color: #475569;
            font-size: 0.86rem;
            font-weight: 650;
        }

        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: #0F172A;
            font-weight: 800;
        }

        div[data-testid="stMetricDelta"] {
            color: #2F7D69;
        }

        .dashboard-subtitle {
            color: #475569;
            font-size: 0.98rem;
            margin-top: -0.35rem;
            margin-bottom: 1rem;
        }

        .insight-card {
            background: linear-gradient(135deg, #FFFFFF 0%, #F7FBF8 100%);
            border: 1px solid #D8E2DC;
            border-top: 4px solid #D9A441;
            border-radius: 8px;
            padding: 15px 16px;
            min-height: 116px;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
        }

        .insight-label {
            color: #5B6B66;
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0;
            font-weight: 800;
            margin-bottom: 0.35rem;
        }

        .insight-value {
            color: #102A2A;
            font-size: 1.08rem;
            font-weight: 800;
            line-height: 1.25;
            margin-bottom: 0.45rem;
        }

        .insight-detail {
            color: #52635E;
            font-size: 0.88rem;
            line-height: 1.35;
        }

        .empty-state {
            background: #F7FBF8;
            border: 1px solid #D8E2DC;
            border-left: 5px solid #D9A441;
            border-radius: 8px;
            padding: 22px 24px;
            margin-top: 1rem;
        }

        .empty-title {
            color: #102A2A;
            font-size: 1.1rem;
            font-weight: 800;
            margin-bottom: 0.35rem;
        }

        .empty-body {
            color: #475569;
            font-size: 0.95rem;
            line-height: 1.45;
        }

        .section-note {
            color: #64748B;
            font-size: 0.9rem;
            margin-top: -0.4rem;
            margin-bottom: 0.6rem;
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid #D8E2DC;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);
        }

        button[kind="secondary"] {
            border-color: #2F7D69;
            color: #12312F;
        }

        button[kind="secondary"]:hover {
            border-color: #D9A441;
            color: #102A2A;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }

        .stTabs [data-baseweb="tab"] {
            background: #FFFFFF;
            border: 1px solid #D8E2DC;
            border-radius: 8px 8px 0 0;
            padding: 8px 14px;
        }

        .stTabs [aria-selected="true"] {
            background: #E7F0EB;
            color: #12312F;
            font-weight: 800;
        }

        .developer-footer {
            margin-top: 2rem;
            padding: 16px 0 6px 0;
            border-top: 1px solid #D8E2DC;
            color: #52635E;
            font-size: 0.86rem;
            text-align: center;
            line-height: 1.5;
        }

        .developer-footer strong {
            color: #12312F;
            font-weight: 800;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# Text cleaning and label helpers
def clean_text(value):
    if pd.isna(value):
        return pd.NA
    value = str(value).strip()
    return pd.NA if value == "" else " ".join(value.split())


def age_group_life_stage(age_group):
    age_group = clean_text(age_group)

    if pd.isna(age_group):
        return pd.NA

    if age_group in CHILD_AGE_GROUPS:
        return "Child"

    if age_group in ADULT_AGE_GROUPS:
        return "Adult"

    value = str(age_group).lower()
    numbers = [int(number) for number in re.findall(r"\d+", value)]

    if numbers:
        first_age = numbers[0]
        return "Child" if first_age < 18 else "Adult"

    return pd.NA


def normalize_gender_by_life_stage(gender, life_stage):
    gender = clean_text(gender)
    life_stage = clean_text(life_stage)

    if pd.isna(gender):
        return "[Missing]"

    if pd.isna(life_stage):
        return gender

    if life_stage == "Adult":
        return {
            "Girl": "Woman",
            "Boy": "Man",
        }.get(gender, gender)

    if life_stage == "Child":
        return {
            "Woman": "Girl",
            "Man": "Boy",
        }.get(gender, gender)

    return gender


def safe_label_from_code(value):
    value = str(value)
    value = value.replace("concern_", "")
    value = value.replace("info_", "")
    value = value.replace("ref_partner_", "")
    value = value.replace("_", " ")
    return value.title()


def short_axis_label(value, max_chars=28):
    value = str(value)
    if len(value) <= max_chars:
        return value
    return value[: max_chars - 3] + "..."


def build_label_map(mapping, prefix):
    if mapping is None:
        return {}

    label_map = {}
    selected = mapping[mapping["cleaned_column_name"].astype(str).str.startswith(prefix)]

    for _, row in selected.iterrows():
        cleaned_name = row["cleaned_column_name"]

        if str(cleaned_name).endswith("_specify"):
            continue

        original_name = clean_text(row["original_column_name"])

        if pd.isna(original_name):
            label = safe_label_from_code(cleaned_name)
        elif "/" in str(original_name):
            label = str(original_name).split("/", 1)[1].strip()
        else:
            label = str(original_name).strip()

        label = label.replace("Protection_concerns", "Protection concerns")
        label = label.replace("General_protection", "General protection")
        label = label.replace("Commitee", "Committee")
        label = label.replace("intergration", "integration")

        label_map[cleaned_name] = label

    return label_map


# Data loading, reshaping, and KPI preparation
@st.cache_data
def load_data():
    if not DATA_FILE_PATH.exists():
        st.error(f"File not found: {DATA_FILE_PATH}")
        st.stop()

    records = pd.read_excel(DATA_FILE_PATH, sheet_name="cleaned_data")

    try:
        mapping = pd.read_excel(DATA_FILE_PATH, sheet_name="Column Mapping")
    except Exception:
        mapping = None

    records["source_row_number"] = records.index + 2
    records["record_id"] = records["source_row_number"].map(lambda row: f"HD-{row:05d}")

    records["interview_date"] = pd.to_datetime(records["interview_date"], errors="coerce")
    records["referral_date"] = pd.to_datetime(records["referral_date"], errors="coerce")

    records["year"] = records["interview_date"].dt.year
    records["month_number"] = records["interview_date"].dt.month
    records["year_month"] = records["interview_date"].dt.to_period("M").astype(str)
    records["month_label"] = records["interview_date"].dt.strftime("%b %Y")

    records["age_group"] = records["information_seeker_age"].map(clean_text)
    records["derived_life_stage"] = records["age_group"].map(age_group_life_stage)

    records["request_category"] = records["request_type_protection_or_information"].map(clean_text)
    records["action_taken_clean"] = records["action_taken"].map(clean_text)
    records["follow_up_required_clean"] = records["follow_up_required"].map(clean_text)

    records["helpdesk_location"] = records["helpdesk_camp_location"].map(clean_text)
    records["helpdesk_location"] = records["helpdesk_location"].fillna(
        records["helpdesk_village"].map(clean_text)
    )
    records["helpdesk_location"] = records["helpdesk_location"].fillna("[Not recorded]")

    records["information_seeker_type_raw"] = records["information_seeker_type"].map(clean_text)
    records["information_seeker_gender_raw"] = records["information_seeker_gender"].map(clean_text)

    records["information_seeker_type"] = records["derived_life_stage"].fillna(
        records["information_seeker_type_raw"]
    )

    records["information_seeker_gender"] = records.apply(
        lambda row: normalize_gender_by_life_stage(
            row["information_seeker_gender_raw"],
            row["information_seeker_type"],
        ),
        axis=1,
    )
    records["information_seeker_gender"] = records["information_seeker_gender"].fillna("[Missing]")

    records["type_age_correction_flag"] = (
        records["information_seeker_type_raw"].fillna("[Missing]")
        != records["information_seeker_type"].fillna("[Missing]")
    )

    records["gender_age_correction_flag"] = (
        records["information_seeker_gender_raw"].fillna("[Missing]")
        != records["information_seeker_gender"].fillna("[Missing]")
    )

    records["referral_status"] = "No referral"
    records.loc[
        records["action_taken_clean"].eq("Case referrred to Tdh national staff"),
        "referral_status",
    ] = "Referred to Tdh national staff"
    records.loc[
        records["action_taken_clean"].eq("Case referred to partner agencies"),
        "referral_status",
    ] = "Referred to partner agency"
    records.loc[
        records["action_taken_clean"].eq(
            "Case not referred to any partner BUT information counselling provided"
        ),
        "referral_status",
    ] = "Information counselling only"
    records.loc[
        records["action_taken_clean"].eq("Action not taken at all"),
        "referral_status",
    ] = "No action taken"

    core_fields = [
        "interview_date",
        "information_seeker_type",
        "camp_location",
        "information_seeker_gender",
        "age_group",
        "request_category",
        "action_taken_clean",
    ]

    records = records[records[core_fields].notna().all(axis=1)].copy()

    id_cols = [
        "record_id",
        "interview_date",
        "year_month",
        "camp_location",
        "helpdesk_location",
        "information_seeker_type",
        "information_seeker_gender",
        "age_group",
        "derived_life_stage",
        "information_seeker_type_raw",
        "information_seeker_gender_raw",
        "type_age_correction_flag",
        "gender_age_correction_flag",
        "request_category",
        "referral_status",
        "follow_up_required_clean",
    ]

    protection_cols = [
        col for col in records.columns
        if col.startswith("concern_") and not col.endswith("_specify")
    ]
    information_cols = [
        col for col in records.columns
        if col.startswith("info_") and not col.endswith("_specify")
    ]
    referral_cols = [
        col for col in records.columns
        if col.startswith("ref_partner_") and not col.endswith("_specify")
    ]

    protection_label_map = build_label_map(mapping, "concern_")
    information_label_map = build_label_map(mapping, "info_")
    referral_label_map = build_label_map(mapping, "ref_partner_")

    protection = records[id_cols + protection_cols].melt(
        id_vars=id_cols,
        value_vars=protection_cols,
        var_name="protection_concern_code",
        value_name="selected",
    )
    protection = protection[pd.to_numeric(protection["selected"], errors="coerce").eq(1)]
    protection = protection.drop(columns="selected")
    protection["protection_concern"] = protection["protection_concern_code"].map(
        protection_label_map
    )
    protection["protection_concern"] = protection["protection_concern"].fillna(
        protection["protection_concern_code"].map(safe_label_from_code)
    )

    information = records[id_cols + information_cols].melt(
        id_vars=id_cols,
        value_vars=information_cols,
        var_name="general_information_code",
        value_name="selected",
    )
    information = information[pd.to_numeric(information["selected"], errors="coerce").eq(1)]
    information = information.drop(columns="selected")
    information["general_information_need"] = information["general_information_code"].map(
        information_label_map
    )
    information["general_information_need"] = information["general_information_need"].fillna(
        information["general_information_code"].map(safe_label_from_code)
    )

    referrals = records[id_cols + referral_cols].melt(
        id_vars=id_cols,
        value_vars=referral_cols,
        var_name="referral_partner_code",
        value_name="selected",
    )
    referrals = referrals[pd.to_numeric(referrals["selected"], errors="coerce").eq(1)]
    referrals = referrals.drop(columns="selected")
    referrals["referral_partner"] = referrals["referral_partner_code"].map(referral_label_map)
    referrals["referral_partner"] = referrals["referral_partner"].fillna(
        referrals["referral_partner_code"].map(safe_label_from_code)
    )

    dashboard_records = records.drop(
        columns=[col for col in PII_COLUMNS if col in records.columns],
        errors="ignore",
    )

    kpis = pd.DataFrame(
        {
            "metric": [
                "valid_dashboard_records",
                "protection_concern_records",
                "general_information_records",
                "partner_referral_records",
                "follow_up_required_records",
                "type_age_corrections",
                "gender_age_corrections",
            ],
            "value": [
                len(dashboard_records),
                dashboard_records["request_category"].eq("Reporting a protection concern").sum(),
                dashboard_records["request_category"].eq(
                    "Seeking general protection information"
                ).sum(),
                dashboard_records["referral_status"].eq("Referred to partner agency").sum(),
                dashboard_records["follow_up_required_clean"].eq("Yes").sum(),
                dashboard_records["type_age_correction_flag"].sum(),
                dashboard_records["gender_age_correction_flag"].sum(),
            ],
        }
    )

    return dashboard_records, protection, information, referrals, kpis


# Formatting and filter helpers
def format_number(value):
    return f"{int(value):,}"


def format_rate(numerator, denominator):
    if denominator == 0:
        return "0.0%"
    return f"{numerator / denominator:.1%}"


def filter_options(series, ordered_values=None):
    values = series.dropna().astype(str).unique().tolist()

    if ordered_values:
        ordered = [value for value in ordered_values if value in values]
        remaining = sorted([value for value in values if value not in ordered])
        return ordered + remaining

    return sorted(values)


def sanitize_multiselect_state(key, options):
    current = st.session_state.get(key, [])
    allowed = set(options)
    cleaned = [value for value in current if value in allowed]

    if cleaned != current:
        st.session_state[key] = cleaned


def sanitize_date_state(key, min_date, max_date, default_date):
    current = st.session_state.get(key, default_date)

    if current < min_date or current > max_date:
        st.session_state[key] = default_date


def reset_filters(default_from_date, max_date):
    st.session_state["from_date_filter"] = default_from_date
    st.session_state["to_date_filter"] = max_date

    for key in FILTER_KEYS:
        st.session_state[key] = []

    st.session_state["records_search"] = ""


def apply_filters(frame, filters):
    filtered = frame.copy()

    if "interview_date" in filtered.columns:
        start_date = filters["start_date"]
        end_exclusive = filters["end_date"] + pd.Timedelta(days=1)

        filtered = filtered[
            filtered["interview_date"].ge(start_date)
            & filtered["interview_date"].lt(end_exclusive)
        ]

    for column, selected in [
        ("camp_location", filters["camp_location"]),
        ("helpdesk_location", filters["helpdesk_location"]),
        ("information_seeker_type", filters["information_seeker_type"]),
        ("information_seeker_gender", filters["information_seeker_gender"]),
        ("age_group", filters["age_group"]),
        ("request_category", filters["request_category"]),
    ]:
        if selected and column in filtered.columns:
            filtered = filtered[filtered[column].astype(str).isin(selected)]

    return filtered


# Chart helpers
def gender_color(field):
    return alt.Color(
        field,
        title="Gender",
        scale=alt.Scale(
            domain=GENDER_ORDER,
            range=[GENDER_COLORS[gender] for gender in GENDER_ORDER],
        ),
        sort=GENDER_ORDER,
    )


def polish_chart(chart):
    return (
        chart.configure_axis(
            labelColor="#475569",
            titleColor="#334155",
            gridColor="#D8E2DC",
            domainColor="#CBD5D1",
            tickColor="#CBD5D1",
            labelFontSize=11,
            titleFontSize=12,
        )
        .configure_legend(
            labelColor="#334155",
            titleColor="#334155",
            labelFontSize=11,
            titleFontSize=12,
            orient="bottom",
        )
        .configure_view(strokeWidth=0)
        .configure(background="#FFFFFF")
    )


def gender_pivot_table(frame, category_column, category_label, top_n=None):
    if frame.empty or category_column not in frame.columns:
        return pd.DataFrame()

    table = (
        frame.groupby([category_column, "information_seeker_gender"], dropna=False)
        .size()
        .reset_index(name="records")
    )

    if top_n:
        top_values = (
            table.groupby(category_column)["records"]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index
        )
        table = table[table[category_column].isin(top_values)]

    pivot = table.pivot_table(
        index=category_column,
        columns="information_seeker_gender",
        values="records",
        aggfunc="sum",
        fill_value=0,
    )

    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.reset_index()
    pivot = pivot.rename(columns={category_column: category_label})

    ordered_columns = [category_label]

    for gender in GENDER_ORDER:
        if gender in pivot.columns:
            ordered_columns.append(gender)

    other_gender_columns = [
        col for col in pivot.columns
        if col not in ordered_columns and col != "Total"
    ]

    ordered_columns.extend(other_gender_columns)
    ordered_columns.append("Total")
    pivot = pivot[ordered_columns]

    if category_column == "age_group":
        age_order_map = {age: index for index, age in enumerate(AGE_GROUP_ORDER)}
        pivot["_sort_order"] = pivot[category_label].map(age_order_map).fillna(999)
        pivot = pivot.sort_values("_sort_order").drop(columns="_sort_order")
    else:
        pivot = pivot.sort_values("Total", ascending=False)

    numeric_columns = [col for col in pivot.columns if col != category_label]

    total_row = {category_label: "Total"}
    for col in numeric_columns:
        total_row[col] = pivot[col].sum()

    return pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)


# Table styling helpers
def style_total_table(table, label_column):
    numeric_columns = [col for col in table.columns if col != label_column]

    def highlight_total_row(row):
        if row[label_column] == "Total":
            return [
                "background-color: #DDEDE5; color: #102A2A; font-weight: 800;"
                for _ in row
            ]

        background = "#FFFFFF" if row.name % 2 == 0 else "#F7FAF8"
        return [f"background-color: {background};" for _ in row]

    def highlight_total_column(column):
        if column.name == "Total":
            return [
                "background-color: #FFF4D8; color: #102A2A; font-weight: 800;"
                for _ in column
            ]
        return ["" for _ in column]

    return (
        table.style
        .format({col: "{:,.0f}" for col in numeric_columns})
        .apply(highlight_total_row, axis=1)
        .apply(highlight_total_column, axis=0)
        .set_properties(
            subset=[label_column],
            **{
                "text-align": "left",
                "font-weight": "650",
                "white-space": "normal",
            },
        )
        .set_properties(
            subset=numeric_columns,
            **{
                "text-align": "center",
                "font-variant-numeric": "tabular-nums",
            },
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#12312F"),
                        ("color", "#FFFFFF"),
                        ("font-weight", "800"),
                        ("text-align", "center"),
                        ("border", "1px solid #D8E2DC"),
                    ],
                },
                {
                    "selector": "td",
                    "props": [
                        ("border", "1px solid #E5E7EB"),
                        ("padding", "7px 9px"),
                    ],
                },
            ]
        )
    )


def style_records_table(table):
    display_table = table.copy()

    date_columns = [
        col for col in display_table.columns
        if "date" in col.lower()
    ]
    numeric_columns = display_table.select_dtypes(include="number").columns.tolist()

    formatters = {
        col: (
            lambda value: ""
            if pd.isna(value)
            else pd.to_datetime(value).strftime("%d %b %Y")
        )
        for col in date_columns
    }
    formatters.update({col: "{:,.0f}" for col in numeric_columns})

    def zebra_rows(row):
        background = "#FFFFFF" if row.name % 2 == 0 else "#F7FAF8"
        return [f"background-color: {background};" for _ in row]

    return (
        display_table.style
        .format(formatters)
        .apply(zebra_rows, axis=1)
        .set_properties(
            **{
                "border": "1px solid #E5E7EB",
                "padding": "7px 9px",
                "white-space": "normal",
            }
        )
        .set_table_styles(
            [
                {
                    "selector": "th",
                    "props": [
                        ("background-color", "#12312F"),
                        ("color", "#FFFFFF"),
                        ("font-weight", "800"),
                        ("text-align", "left"),
                        ("border", "1px solid #D8E2DC"),
                    ],
                },
            ]
        )
    )


# Chart and table rendering helpers
def show_gender_table(frame, category_column, category_label, top_n=None):
    table = gender_pivot_table(frame, category_column, category_label, top_n=top_n)

    if table.empty:
        st.info("No records match the selected filters.")
        return

    st.dataframe(
        style_total_table(table, category_label),
        use_container_width=True,
        hide_index=True,
    )


def gender_wide_chart_data(frame, category_column, top_n=None):
    table = gender_pivot_table(frame, category_column, category_column, top_n=top_n)

    if table.empty:
        return pd.DataFrame()

    table = table[table[category_column] != "Total"]
    chart_data = table.set_index(category_column)

    if "Total" in chart_data.columns:
        chart_data = chart_data.drop(columns="Total")

    return chart_data


def draw_gender_bar(frame, category_column, top_n=None, height=430):
    chart_data = gender_wide_chart_data(frame, category_column, top_n=top_n)

    if chart_data.empty:
        st.info("No records match the selected filters.")
        return

    chart_data = chart_data.reset_index()
    gender_columns = [col for col in chart_data.columns if col != category_column]

    if category_column == "age_group":
        category_order = chart_data[category_column].astype(str).tolist()
    else:
        category_order = (
            chart_data.assign(Total=chart_data[gender_columns].sum(axis=1))
            .sort_values("Total", ascending=True)[category_column]
            .astype(str)
            .tolist()
        )

    long_chart = chart_data.melt(
        id_vars=[category_column],
        value_vars=gender_columns,
        var_name="Gender",
        value_name="Records",
    )

    chart_height = max(height, min(900, 36 * len(category_order) + 80))

    chart = (
        alt.Chart(long_chart)
        .mark_bar(cornerRadiusEnd=2, stroke="#FFFFFF", strokeWidth=0.5)
        .encode(
            y=alt.Y(
                f"{category_column}:N",
                sort=category_order,
                title=None,
                axis=alt.Axis(labelLimit=700, labelFontSize=11, labelPadding=6),
            ),
            x=alt.X("Records:Q", title="Records", stack="zero"),
            color=gender_color("Gender:N"),
            tooltip=[
                alt.Tooltip(f"{category_column}:N", title="Category"),
                alt.Tooltip("Gender:N", title="Gender"),
                alt.Tooltip("Records:Q", title="Records", format=","),
            ],
        )
        .properties(height=chart_height)
    )

    st.altair_chart(polish_chart(chart), use_container_width=True)


def draw_gender_column_bar(frame, category_column, top_n=None, height=360):
    chart_data = gender_wide_chart_data(frame, category_column, top_n=top_n)

    if chart_data.empty:
        st.info("No records match the selected filters.")
        return

    chart_data = chart_data.reset_index()
    gender_columns = [col for col in chart_data.columns if col != category_column]

    if category_column == "age_group":
        category_order = chart_data[category_column].astype(str).tolist()
        max_chars = 18
    elif category_column == "helpdesk_location":
        category_order = (
            chart_data.assign(Total=chart_data[gender_columns].sum(axis=1))
            .sort_values("Total", ascending=False)[category_column]
            .astype(str)
            .tolist()
        )
        max_chars = 16
    else:
        category_order = (
            chart_data.assign(Total=chart_data[gender_columns].sum(axis=1))
            .sort_values("Total", ascending=False)[category_column]
            .astype(str)
            .tolist()
        )
        max_chars = 24

    chart_data["axis_label"] = chart_data[category_column].map(
        lambda value: short_axis_label(value, max_chars=max_chars)
    )

    axis_order = [
        chart_data.loc[chart_data[category_column].astype(str).eq(value), "axis_label"].iloc[0]
        for value in category_order
        if not chart_data.loc[chart_data[category_column].astype(str).eq(value)].empty
    ]

    long_chart = chart_data.melt(
        id_vars=[category_column, "axis_label"],
        value_vars=gender_columns,
        var_name="Gender",
        value_name="Records",
    )

    chart = (
        alt.Chart(long_chart)
        .mark_bar(cornerRadiusEnd=2, stroke="#FFFFFF", strokeWidth=0.5)
        .encode(
            x=alt.X(
                "axis_label:N",
                sort=axis_order,
                title=None,
                axis=alt.Axis(labelAngle=-30, labelLimit=150, labelFontSize=10),
            ),
            y=alt.Y("Records:Q", title="Records", stack="zero"),
            color=gender_color("Gender:N"),
            tooltip=[
                alt.Tooltip(f"{category_column}:N", title="Category"),
                alt.Tooltip("Gender:N", title="Gender"),
                alt.Tooltip("Records:Q", title="Records", format=","),
            ],
        )
        .properties(height=height)
    )

    st.altair_chart(polish_chart(chart), use_container_width=True)


def draw_monthly_gender_column_bar(frame, height=340):
    if frame.empty:
        st.info("No records match the selected filters.")
        return

    monthly = (
        frame.groupby(["year_month", "information_seeker_gender"], dropna=False)
        .size()
        .reset_index(name="Records")
    )

    if monthly.empty:
        st.info("No monthly trend data for the selected filters.")
        return

    month_order = sorted(monthly["year_month"].dropna().astype(str).unique().tolist())

    line = (
        alt.Chart(monthly)
        .mark_line(point=True, strokeWidth=3)
        .encode(
            x=alt.X(
                "year_month:N",
                sort=month_order,
                title=None,
                axis=alt.Axis(labelAngle=-30, labelFontSize=11),
            ),
            y=alt.Y("Records:Q", title="Records"),
            color=gender_color("information_seeker_gender:N"),
            tooltip=[
                alt.Tooltip("year_month:N", title="Month"),
                alt.Tooltip("information_seeker_gender:N", title="Gender"),
                alt.Tooltip("Records:Q", title="Records", format=","),
            ],
        )
        .properties(height=height)
    )

    st.altair_chart(polish_chart(line), use_container_width=True)


# Insight cards, empty states, footer, and record search
def top_value(frame, column):
    if frame.empty or column not in frame.columns:
        return "None", 0

    counts = frame[column].dropna().astype(str).value_counts()

    if counts.empty:
        return "None", 0

    return counts.index[0], int(counts.iloc[0])


def escape_text(value):
    return html.escape(str(value))


def show_insight_card(column, label, value, detail):
    with column:
        st.markdown(
            f"""
            <div class="insight-card">
                <div class="insight-label">{escape_text(label)}</div>
                <div class="insight-value">{escape_text(value)}</div>
                <div class="insight-detail">{escape_text(detail)}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def show_empty_state(from_date, to_date):
    st.markdown(
        f"""
        <div class="empty-state">
            <div class="empty-title">No records match the selected filters</div>
            <div class="empty-body">
                The current selection covers {escape_text(from_date.strftime("%d %b %Y"))}
                to {escape_text(to_date.strftime("%d %b %Y"))}. Clear one or more filters,
                expand the date range, or use Reset filters from the sidebar.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_footer():
    st.markdown(
        """
        <div class="developer-footer">
            <div><strong>ImpactLens Africa</strong></div>
            <div>Turning Data Into Human Impact</div>
            <div>Developed by John Kul, MEAL Officer, Tdh.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def filter_label(values, max_items=3):
    if not values:
        return "All"

    values = list(values)
    shown = values[:max_items]
    suffix = "" if len(values) <= max_items else f" +{len(values) - max_items} more"
    return ", ".join(shown) + suffix


def search_records(frame, query):
    if not query:
        return frame

    searchable = frame.copy()
    mask = pd.Series(False, index=searchable.index)

    for column in searchable.columns:
        mask = mask | searchable[column].astype(str).str.contains(
            query,
            case=False,
            regex=False,
            na=False,
        )

    return searchable[mask]


# Load source data and prepare date boundaries
inject_theme_css()

records, protection, information, referrals, kpis = load_data()

if records.empty:
    st.error("No valid dashboard records were found in the source file.")
    st.stop()

min_date = records["interview_date"].min().date()
max_date = records["interview_date"].max().date()
default_from_date = min_date

calendar_min_date = pd.Timestamp(year=min_date.year, month=1, day=1).date()
calendar_max_date = pd.Timestamp(year=max_date.year, month=12, day=31).date()

if "from_date_filter" not in st.session_state:
    st.session_state["from_date_filter"] = default_from_date

if "to_date_filter" not in st.session_state:
    st.session_state["to_date_filter"] = max_date

sanitize_date_state("from_date_filter", calendar_min_date, calendar_max_date, default_from_date)
sanitize_date_state("to_date_filter", calendar_min_date, calendar_max_date, max_date)


# Sidebar filters
# Filter order is intentionally cascading: Camp -> Helpdesk -> Information seeker -> Gender -> Age -> Request
with st.sidebar:
    st.header("Filters")

    st.button(
        "Reset filters",
        use_container_width=True,
        on_click=reset_filters,
        args=(default_from_date, max_date),
    )

    st.markdown("### Date range")

    selected_from_date = st.date_input(
        "From:",
        min_value=calendar_min_date,
        max_value=calendar_max_date,
        key="from_date_filter",
    )

    selected_to_date = st.date_input(
        "To:",
        min_value=calendar_min_date,
        max_value=calendar_max_date,
        key="to_date_filter",
    )

    if selected_from_date > selected_to_date:
        st.error("The From date cannot be later than the To date.")
        st.stop()

    from_date = selected_from_date
    to_date = selected_to_date

    if from_date < min_date:
        st.warning(f"The earliest available record is {min_date.strftime('%d %b %Y')}.")
        from_date = min_date

    if to_date > max_date:
        st.warning(f"The latest available record is {max_date.strftime('%d %b %Y')}.")
        to_date = max_date

    start_date = pd.to_datetime(from_date)
    end_date = pd.to_datetime(to_date)

    date_filtered_records = records[
        records["interview_date"].ge(start_date)
        & records["interview_date"].lt(end_date + pd.Timedelta(days=1))
    ].copy()

    camp_options = filter_options(date_filtered_records["camp_location"])
    sanitize_multiselect_state("camp_location_filter", camp_options)

    selected_camp_locations = st.multiselect(
        "Camp location",
        camp_options,
        key="camp_location_filter",
    )

    selected_helpdesk_locations = []
    selected_information_seeker_types = []
    selected_genders = []
    selected_age_groups = []
    selected_request_categories = []

    if not selected_camp_locations:
        for key in [
            "helpdesk_location_filter",
            "information_seeker_type_filter",
            "information_seeker_gender_filter",
            "age_group_filter",
            "request_category_filter",
        ]:
            st.session_state[key] = []

        st.info("Select a camp location first to unlock the remaining filters.")

    else:
        camp_filtered_records = date_filtered_records[
            date_filtered_records["camp_location"].astype(str).isin(selected_camp_locations)
        ].copy()

        helpdesk_options = filter_options(camp_filtered_records["helpdesk_location"])
        sanitize_multiselect_state("helpdesk_location_filter", helpdesk_options)

        selected_helpdesk_locations = st.multiselect(
            "Helpdesk location",
            helpdesk_options,
            key="helpdesk_location_filter",
        )

        if selected_helpdesk_locations:
            helpdesk_filtered_records = camp_filtered_records[
                camp_filtered_records["helpdesk_location"]
                .astype(str)
                .isin(selected_helpdesk_locations)
            ].copy()
        else:
            helpdesk_filtered_records = camp_filtered_records.copy()

        seeker_type_options = filter_options(helpdesk_filtered_records["information_seeker_type"])
        sanitize_multiselect_state("information_seeker_type_filter", seeker_type_options)

        selected_information_seeker_types = st.multiselect(
            "Information seeker type",
            seeker_type_options,
            key="information_seeker_type_filter",
        )

        if selected_information_seeker_types:
            seeker_filtered_records = helpdesk_filtered_records[
                helpdesk_filtered_records["information_seeker_type"]
                .astype(str)
                .isin(selected_information_seeker_types)
            ].copy()
        else:
            seeker_filtered_records = helpdesk_filtered_records.copy()

        gender_options = filter_options(
            seeker_filtered_records["information_seeker_gender"],
            ordered_values=GENDER_ORDER,
        )
        sanitize_multiselect_state("information_seeker_gender_filter", gender_options)

        selected_genders = st.multiselect(
            "Gender",
            gender_options,
            key="information_seeker_gender_filter",
        )

        if selected_genders:
            gender_filtered_records = seeker_filtered_records[
                seeker_filtered_records["information_seeker_gender"]
                .astype(str)
                .isin(selected_genders)
            ].copy()
        else:
            gender_filtered_records = seeker_filtered_records.copy()

        age_options = filter_options(
            gender_filtered_records["age_group"],
            ordered_values=AGE_GROUP_ORDER,
        )
        sanitize_multiselect_state("age_group_filter", age_options)

        selected_age_groups = st.multiselect(
            "Age group",
            age_options,
            key="age_group_filter",
        )

        if selected_age_groups:
            age_filtered_records = gender_filtered_records[
                gender_filtered_records["age_group"]
                .astype(str)
                .isin(selected_age_groups)
            ].copy()
        else:
            age_filtered_records = gender_filtered_records.copy()

        request_options = filter_options(age_filtered_records["request_category"])
        sanitize_multiselect_state("request_category_filter", request_options)

        selected_request_categories = st.multiselect(
            "Request category",
            request_options,
            key="request_category_filter",
        )

    st.divider()
    st.caption(f"Camp: {filter_label(selected_camp_locations)}")
    st.caption(f"Helpdesk: {filter_label(selected_helpdesk_locations)}")
    st.caption(f"Gender: {filter_label(selected_genders)}")
    st.caption(f"Age: {filter_label(selected_age_groups)}")


# Apply filters to each dashboard dataset
filters = {
    "start_date": start_date,
    "end_date": end_date,
    "camp_location": selected_camp_locations,
    "helpdesk_location": selected_helpdesk_locations,
    "information_seeker_type": selected_information_seeker_types,
    "information_seeker_gender": selected_genders,
    "age_group": selected_age_groups,
    "request_category": selected_request_categories,
}

filtered_records = apply_filters(records, filters)
filtered_protection = apply_filters(protection, filters)
filtered_information = apply_filters(information, filters)
filtered_referrals = apply_filters(referrals, filters)


# KPI calculations
total_records = len(filtered_records)
all_records = len(records)

protection_records = filtered_records["request_category"].eq(
    "Reporting a protection concern"
).sum()
information_records = filtered_records["request_category"].eq(
    "Seeking general protection information"
).sum()
partner_referrals = filtered_records["referral_status"].eq(
    "Referred to partner agency"
).sum()
follow_up = filtered_records["follow_up_required_clean"].eq("Yes").sum()


# Dashboard header and KPI row
header_left, header_right = st.columns([0.12, 0.88], vertical_alignment="center")

with header_left:
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=86)

with header_right:
    st.title("Tdh Kenya Helpdesk Data Dashboard")

st.markdown(
    f"""
    <div class="dashboard-subtitle">
        Showing {format_number(total_records)} of {format_number(all_records)} records |
        {from_date.strftime("%d %b %Y")} to {to_date.strftime("%d %b %Y")}
    </div>
    """,
    unsafe_allow_html=True,
)

kpi_cols = st.columns(5)
kpi_cols[0].metric(
    "Valid records",
    format_number(total_records),
    f"{format_rate(total_records, all_records)} of source",
)
kpi_cols[1].metric(
    "Protection concerns",
    format_number(protection_records),
    f"{format_rate(protection_records, total_records)} of records",
)
kpi_cols[2].metric(
    "Information requests",
    format_number(information_records),
    f"{format_rate(information_records, total_records)} of records",
)
kpi_cols[3].metric(
    "Partner referrals",
    format_number(partner_referrals),
    f"{format_rate(partner_referrals, total_records)} referral rate",
)
kpi_cols[4].metric(
    "Follow-up required",
    format_number(follow_up),
    f"{format_rate(follow_up, total_records)} follow-up rate",
)

if filtered_records.empty:
    show_empty_state(from_date, to_date)
    show_footer()
    st.stop()


# Quick insight cards
top_location, top_location_count = top_value(filtered_records, "helpdesk_location")
top_request, top_request_count = top_value(filtered_records, "request_category")
top_concern, top_concern_count = top_value(filtered_protection, "protection_concern")
top_follow_up, top_follow_up_count = top_value(filtered_records, "follow_up_required_clean")

st.subheader("Quick Insights")
insight_cols = st.columns(4)
show_insight_card(
    insight_cols[0],
    "Busiest helpdesk",
    top_location,
    f"{format_number(top_location_count)} records",
)
show_insight_card(
    insight_cols[1],
    "Top request type",
    top_request,
    f"{format_number(top_request_count)} records",
)
show_insight_card(
    insight_cols[2],
    "Top protection concern",
    top_concern,
    f"{format_number(top_concern_count)} mentions",
)
show_insight_card(
    insight_cols[3],
    "Follow-up status",
    top_follow_up,
    f"{format_number(top_follow_up_count)} records",
)

st.divider()


# Dashboard tabs
tab_overview, tab_concerns, tab_information, tab_referrals, tab_data = st.tabs(
    ["Overview", "Concerns", "Information", "Referrals", "Records"]
)


# Overview tab
with tab_overview:
    left, right = st.columns([1.2, 1])

    with left:
        st.subheader("Monthly Requests by Gender")
        st.markdown(
            '<div class="section-note">Records grouped by interview month and gender.</div>',
            unsafe_allow_html=True,
        )
        draw_monthly_gender_column_bar(filtered_records, height=340)

    with right:
        st.subheader("Request Type by Gender")
        st.markdown(
            '<div class="section-note">Protection concerns and information requests by gender.</div>',
            unsafe_allow_html=True,
        )
        draw_gender_column_bar(filtered_records, "request_category", height=340)

    st.subheader("Request Type Table")
    show_gender_table(filtered_records, "request_category", "Request type")

    st.divider()

    st.subheader("Demographics by Gender")
    demo_cols = st.columns(2)

    with demo_cols[0]:
        st.caption("Information seeker type")
        draw_gender_column_bar(filtered_records, "information_seeker_type", height=300)
        show_gender_table(filtered_records, "information_seeker_type", "Information seeker type")

    with demo_cols[1]:
        st.caption("Age group")
        draw_gender_column_bar(filtered_records, "age_group", height=360)
        show_gender_table(filtered_records, "age_group", "Age group")

    st.divider()

    st.subheader("Location by Gender")
    location_cols = st.columns(2)

    with location_cols[0]:
        st.caption("Camp location")
        draw_gender_column_bar(filtered_records, "camp_location", height=320)
        show_gender_table(filtered_records, "camp_location", "Camp location")

    with location_cols[1]:
        st.caption("Helpdesk location")
        draw_gender_column_bar(filtered_records, "helpdesk_location", height=380)
        show_gender_table(filtered_records, "helpdesk_location", "Helpdesk location")


# Protection concerns tab
with tab_concerns:
    st.subheader("Top Protection Concerns by Gender")

    concern_top_n = st.radio(
        "Show top concerns",
        [10, 15, 25],
        horizontal=True,
        index=1,
        key="concern_top_n",
    )

    draw_gender_bar(
        filtered_protection,
        "protection_concern",
        top_n=concern_top_n,
        height=640,
    )
    show_gender_table(
        filtered_protection,
        "protection_concern",
        "Protection concern",
        top_n=concern_top_n,
    )


# Information needs tab
with tab_information:
    st.subheader("Top General Information Needs by Gender")

    information_top_n = st.radio(
        "Show top information needs",
        [10, 15, 25],
        horizontal=True,
        index=1,
        key="information_top_n",
    )

    draw_gender_bar(
        filtered_information,
        "general_information_need",
        top_n=information_top_n,
        height=640,
    )
    show_gender_table(
        filtered_information,
        "general_information_need",
        "General information need",
        top_n=information_top_n,
    )


# Referrals tab
with tab_referrals:
    st.subheader("Action and Follow-up by Gender")

    action_cols = st.columns(2)

    with action_cols[0]:
        st.caption("Referral status")
        draw_gender_column_bar(filtered_records, "referral_status", height=360)
        show_gender_table(filtered_records, "referral_status", "Referral status")

    with action_cols[1]:
        st.caption("Follow-up required")
        draw_gender_column_bar(filtered_records, "follow_up_required_clean", height=320)
        show_gender_table(filtered_records, "follow_up_required_clean", "Follow-up required")

    st.divider()

    st.subheader("Referral Partners by Gender")

    referral_top_n = st.radio(
        "Show top referral partners",
        [10, 15, 25],
        horizontal=True,
        index=1,
        key="referral_top_n",
    )

    draw_gender_bar(
        filtered_referrals,
        "referral_partner",
        top_n=referral_top_n,
        height=560,
    )
    show_gender_table(
        filtered_referrals,
        "referral_partner",
        "Referral partner",
        top_n=referral_top_n,
    )


# Records tab
with tab_data:
    st.subheader("Filtered Records")

    ordered_columns = [
        col for col in CORE_RECORD_COLUMNS if col in filtered_records.columns
    ] + [
        col for col in filtered_records.columns if col not in CORE_RECORD_COLUMNS
    ]

    default_columns = [col for col in CORE_RECORD_COLUMNS if col in ordered_columns]

    if "record_columns" not in st.session_state:
        st.session_state["record_columns"] = default_columns

    st.session_state["record_columns"] = [
        col for col in st.session_state["record_columns"] if col in ordered_columns
    ]

    selected_columns = st.multiselect("Columns", ordered_columns, key="record_columns")

    if not selected_columns:
        selected_columns = default_columns

    query = st.text_input(
        "Search filtered records",
        placeholder="Search by record ID, location, category, status...",
        key="records_search",
    )

    searched_records = search_records(filtered_records, query)

    st.caption(
        f"Showing {format_number(len(searched_records))} matching records "
        f"from {format_number(len(filtered_records))} filtered records."
    )

    st.dataframe(
        style_records_table(searched_records[selected_columns]),
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "Download filtered records",
        data=searched_records.to_csv(index=False).encode("utf-8"),
        file_name="filtered_helpdesk_records.csv",
        mime="text/csv",
        use_container_width=True,
    )

    with st.expander("Source KPI summary"):
        st.dataframe(
            style_records_table(kpis),
            use_container_width=True,
            hide_index=True,
        )


# Developer footer
show_footer()