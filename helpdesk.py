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


# Source paths
BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "assets" / "tdh-logo.png"
DATA_FILE_PATH = BASE_DIR / "data" / "HELPDESK_DashboardData_Tdh_Kenya_D2.xlsx"


# Source file and dashboard constants
PII_COLUMNS = [
    "information_seeker_name",
    "residence_neighborhood_compound_house",
    "information_seeker_phone",
    "alternative_phone",
    "information_seeker_individual_number",
    "information_seeker_ration_or_wristband_number",
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

CHILD_AGE_GROUPS = {"0-5 Years", "6-11 Years", "12-17 Years"}
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

WGQ_DISABILITY_DOMAINS = {
    "difficulty_seeing": "Visual Impairment",
    "difficulty_hearing": "Hearing Impairment",
    "difficulty_walking_or_climbing": "Physical/Mobility Impairment",
    "difficulty_walking_or_climbing_steps": "Physical/Mobility Impairment",
    "difficulty_remembering_or_concentrating": "Cognitive Impairment",
    "difficulty_self_care": "Self-Care Impairment",
    "difficulty_communicating": "Speech Impairment",
}

ADULT_DISABILITY_CATEGORY_COLUMNS = [
    "information_seeker_disability_type_other",
]

DISABILITY_TYPE_STANDARD_MAP = {
    "visual impairment": "Visual Impairment",
    "visual disability": "Visual Impairment",
    "seeing impairment": "Visual Impairment",
    "seeing disability": "Visual Impairment",
    "hearing impairment": "Hearing Impairment",
    "hearing disability": "Hearing Impairment",
    "physical disability": "Physical/Mobility Impairment",
    "physical impairment": "Physical/Mobility Impairment",
    "physical/mobility disability": "Physical/Mobility Impairment",
    "physical/mobility impairment": "Physical/Mobility Impairment",
    "mobility disability": "Physical/Mobility Impairment",
    "mobility impairment": "Physical/Mobility Impairment",
    "walking disability": "Physical/Mobility Impairment",
    "walking impairment": "Physical/Mobility Impairment",
    "cognitive impairment": "Cognitive Impairment",
    "cognitive disability": "Cognitive Impairment",
    "remembering or concentrating difficulty": "Cognitive Impairment",
    "remembering/concentrating difficulty": "Cognitive Impairment",
    "self-care disability": "Self-Care Impairment",
    "self-care impairment": "Self-Care Impairment",
    "self care disability": "Self-Care Impairment",
    "self care impairment": "Self-Care Impairment",
    "communication disability": "Speech Impairment",
    "communication impairment": "Speech Impairment",
    "speech impairment": "Speech Impairment",
    "speech disability": "Speech Impairment",
    "speech difficulty": "Speech Impairment",
    "autism": "Cognitive Impairment",
    "adhd": "Cognitive Impairment",
    "neurological impairment": "Cognitive Impairment",
    "neurological impairments": "Cognitive Impairment",
    "multiple disabilities": "Multiple Impairments",
    "multiple disability": "Multiple Impairments",
    "multiple impairments": "Multiple Impairments",
    "multiple impairment": "Multiple Impairments",
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
    "household_type",
    "staff_name",
    "gps_latitude",
    "gps_longitude",
    "information_seeker_type",
    "information_seeker_gender",
    "age_group",
    "derived_life_stage",
    "information_seeker_type_raw",
    "information_seeker_gender_raw",
    "type_age_correction_flag",
    "gender_age_correction_flag",
    "disability_status",
    "disability_type",
    "adult_wgq_disability_status",
    "adult_wgq_disability_type",
    "adult_wgq_disability_domains",
    "adult_wgq_domain_count",
    "adult_wgq_impairment_count",
    "adult_duplicate_impairment_mentions",
    "adult_wgq_domain_count_category",
    "adult_wgq_multiplicity",
    "adult_wgq_max_score",
    "adult_wgq_severity",
    "adult_disability_exclusion_risk",
    "adult_additional_disability_category",
    "child_disability_status",
    "child_disability_type",
    "child_disability_type_other",
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


def normalize_response(value):
    value = clean_text(value)

    if pd.isna(value):
        return None

    value = str(value).strip().lower()
    value = value.replace("_", " ")
    value = value.replace("-", " ")
    value = re.sub(r"\s+", " ", value)

    return value


def standardize_disability_type(value):
    value = clean_text(value)

    if pd.isna(value):
        return "None"

    normalized = normalize_response(value)

    if normalized in [
        "none",
        "none of the above",
        "no",
        "not applicable",
        "n/a",
        "na",
        "nil",
    ]:
        return "None"

    if normalized in DISABILITY_TYPE_STANDARD_MAP:
        return DISABILITY_TYPE_STANDARD_MAP[normalized]

    if "multiple" in normalized:
        return "Multiple Impairments"

    if "visual" in normalized or "seeing" in normalized:
        return "Visual Impairment"

    if "hearing" in normalized:
        return "Hearing Impairment"

    if (
        "physical" in normalized
        or "mobility" in normalized
        or "walking" in normalized
        or "climbing" in normalized
    ):
        return "Physical/Mobility Impairment"

    if (
        "cognitive" in normalized
        or "remember" in normalized
        or "concentrat" in normalized
        or "autism" in normalized
        or "adhd" in normalized
        or "neurological" in normalized
    ):
        return "Cognitive Impairment"

    if "self care" in normalized or "self-care" in normalized:
        return "Self-Care Impairment"

    if "speech" in normalized or "communication" in normalized or "communicat" in normalized:
        return "Speech Impairment"

    return str(value)


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


# Dataset freshness helper for Streamlit cache
def data_file_signature(path):
    if not path.exists():
        return str(path), None, None, None

    stat = path.stat()
    return (
        str(path.resolve()),
        stat.st_mtime_ns,
        stat.st_size,
        pd.to_datetime(stat.st_mtime, unit="s").strftime("%d %b %Y %H:%M:%S"),
    )


# Linked form and derived demographic helpers
def age_group_life_stage(age_group):
    age_group = clean_text(age_group)

    if pd.isna(age_group):
        return pd.NA

    if age_group in CHILD_AGE_GROUPS:
        return "Child"

    if age_group in ADULT_AGE_GROUPS:
        return "Adult"

    numbers = [int(number) for number in re.findall(r"\d+", str(age_group))]

    if numbers:
        return "Child" if numbers[0] < 18 else "Adult"

    return pd.NA


def normalize_gender_by_life_stage(gender, life_stage):
    gender = clean_text(gender)
    life_stage = clean_text(life_stage)

    if pd.isna(gender):
        return "[Missing]"

    if pd.isna(life_stage):
        return gender

    if life_stage == "Adult":
        return {"Girl": "Woman", "Boy": "Man"}.get(gender, gender)

    if life_stage == "Child":
        return {"Woman": "Girl", "Man": "Boy"}.get(gender, gender)

    return gender


def is_host_community(value):
    value = normalize_response(value)

    if value is None:
        return False

    return "host" in value and "community" in value


def derive_linked_helpdesk_location(row):
    household_type = row.get("household_type")
    camp_location = clean_text(row.get("camp_location"))
    helpdesk_camp = clean_text(row.get("helpdesk_camp_location"))
    helpdesk_village = clean_text(row.get("helpdesk_village"))

    if is_host_community(household_type):
        if not pd.isna(camp_location):
            return f"Host community - {camp_location}"
        return "Host community"

    if not pd.isna(helpdesk_camp):
        return helpdesk_camp

    if not pd.isna(helpdesk_village):
        return helpdesk_village

    return "[Not recorded]"


# Disability derivation helpers
def is_adult(row):
    age_group = clean_text(row.get("age_group"))

    if not pd.isna(age_group):
        return age_group in ADULT_AGE_GROUPS

    return normalize_response(row.get("information_seeker_type")) == "adult"


def is_child(row):
    age_group = clean_text(row.get("age_group"))

    if not pd.isna(age_group):
        return age_group in CHILD_AGE_GROUPS

    return normalize_response(row.get("information_seeker_type")) == "child"


def wgq_score(value):
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)) and not pd.isna(value):
        score = int(value)
        if score in [1, 2, 3, 4]:
            return score

    response = normalize_response(value)

    if response is None:
        return None

    response = response.replace("can not", "cannot")

    if response.startswith("1") or response == "no difficulty":
        return 1

    if response.startswith("2") or response == "some difficulty":
        return 2

    if response.startswith("3") or response == "a lot of difficulty":
        return 3

    if response.startswith("4") or response == "cannot do at all":
        return 4

    return None


def adult_wgq_domain_scores(row):
    scores = {}

    if not is_adult(row):
        return scores

    for column, impairment_type in WGQ_DISABILITY_DOMAINS.items():
        if column not in row.index:
            continue

        score = wgq_score(row[column])

        if score is None:
            continue

        standardized_type = standardize_disability_type(impairment_type)

        if standardized_type not in scores:
            scores[standardized_type] = score
        else:
            scores[standardized_type] = max(scores[standardized_type], score)

    return scores


def derive_adult_wgq_disability_domains(row):
    scores = adult_wgq_domain_scores(row)

    impairment_types = [
        impairment_type
        for impairment_type, score in scores.items()
        if score in [3, 4]
    ]

    impairment_types = sorted(set(impairment_types))

    if impairment_types:
        return "; ".join(impairment_types)

    return "None"


def derive_adult_additional_disability_category(row):
    if not is_adult(row):
        return "None"

    for column in ADULT_DISABILITY_CATEGORY_COLUMNS:
        if column not in row.index:
            continue

        standardized_value = standardize_disability_type(row[column])

        if standardized_value != "None":
            return standardized_value

    return "None"


def derive_adult_disability_domains(row):
    impairment_types = []

    wgq_domains = derive_adult_wgq_disability_domains(row)
    additional_category = derive_adult_additional_disability_category(row)

    if wgq_domains != "None":
        impairment_types.extend(
            [standardize_disability_type(value) for value in wgq_domains.split("; ")]
        )

    if additional_category != "None":
        impairment_types.append(standardize_disability_type(additional_category))

    impairment_types = sorted(set(value for value in impairment_types if value != "None"))

    if impairment_types:
        return "; ".join(impairment_types)

    return "None"


def split_impairment_types(value):
    if pd.isna(value):
        return []

    value = str(value).strip()

    if value in ["", "None", "No Disability", "[Missing]"]:
        return []

    impairment_types = [item.strip() for item in value.split(";") if item.strip()]
    standardized_types = [standardize_disability_type(item) for item in impairment_types]

    return list(dict.fromkeys(item for item in standardized_types if item != "None"))


def adult_row_impairment_types(row):
    if not is_adult(row):
        return []

    return split_impairment_types(derive_adult_disability_domains(row))


def derive_adult_wgq_disability_status(row):
    if adult_row_impairment_types(row):
        return "Has Disability"

    return "No Disability"


def derive_adult_wgq_disability_type(row):
    impairment_types = adult_row_impairment_types(row)

    if not impairment_types:
        return "No Disability"

    if len(impairment_types) > 1:
        return "Multiple Impairments"

    return impairment_types[0]


def derive_adult_wgq_domain_count(row):
    return len(adult_row_impairment_types(row))


def derive_adult_wgq_impairment_count(row):
    return derive_adult_wgq_domain_count(row)


def derive_adult_duplicate_impairment_mentions(row):
    return max(derive_adult_wgq_impairment_count(row) - 1, 0)


def derive_adult_wgq_multiplicity(row):
    count = derive_adult_wgq_domain_count(row)

    if count == 0:
        return "No Disability"

    if count == 1:
        return "One Impairment"

    return "Multiple Impairments"


def derive_adult_wgq_domain_count_category(row):
    count = derive_adult_wgq_domain_count(row)

    if count == 0:
        return "No Disability"

    if count == 1:
        return "One Impairment"

    if count == 2:
        return "Two Impairments"

    return "Three or More Impairments"


def derive_adult_wgq_max_score(row):
    scores = adult_wgq_domain_scores(row)

    if not scores:
        return 1

    return max(scores.values())


def derive_adult_wgq_severity(row):
    max_score = derive_adult_wgq_max_score(row)

    if max_score in [1, 2]:
        return "No Disability"

    if max_score == 3:
        return "Disability"

    if max_score == 4:
        return "Severe Disability"

    return "No Disability"


def derive_adult_disability_exclusion_risk(row):
    scores = adult_wgq_domain_scores(row)

    if any(score in [2, 3, 4] for score in scores.values()):
        return "At risk of disability-related exclusion"

    if derive_adult_additional_disability_category(row) != "None":
        return "At risk of disability-related exclusion"

    return "Not at risk"


def derive_child_disability_status(row):
    if not is_child(row):
        return "No Disability"

    response = normalize_response(row.get("has_disability"))

    if response in ["yes", "y", "true", "1"]:
        return "Has Disability"

    return "No Disability"


def derive_child_disability_type(row):
    if not is_child(row):
        return "No Disability"

    if derive_child_disability_status(row) != "Has Disability":
        return "No Disability"

    disability_type = clean_text(row.get("child_disability_type"))
    disability_type_other = clean_text(row.get("child_disability_type_other"))

    if not pd.isna(disability_type):
        normalized_type = normalize_response(disability_type)

        if normalized_type not in [
            "other",
            "others",
            "other specify",
            "other specified",
            "none",
            "none of the above",
            "not applicable",
            "n/a",
            "na",
        ]:
            return standardize_disability_type(disability_type)

    if not pd.isna(disability_type_other):
        normalized_other = normalize_response(disability_type_other)

        if normalized_other not in [
            "none",
            "none of the above",
            "not applicable",
            "n/a",
            "na",
            "nil",
        ]:
            return standardize_disability_type(disability_type_other)

    return "Not specified"


def derive_combined_disability_status(row):
    if is_adult(row):
        return row.get("adult_wgq_disability_status", "No Disability")

    if is_child(row):
        return row.get("child_disability_status", "No Disability")

    return "No Disability"


def derive_combined_disability_type(row):
    if is_adult(row):
        return row.get("adult_wgq_disability_type", "No Disability")

    if is_child(row):
        return row.get("child_disability_type", "No Disability")

    return "No Disability"


def adult_wgq_disability_long(frame):
    rows = []

    if frame.empty:
        return pd.DataFrame()

    adult_frame = frame[frame["derived_life_stage"].astype(str).eq("Adult")].copy()

    for _, row in adult_frame.iterrows():
        seen_types = set()

        for column, impairment_type in WGQ_DISABILITY_DOMAINS.items():
            if column not in row.index:
                continue

            score = wgq_score(row[column])

            if score not in [3, 4]:
                continue

            standardized_type = standardize_disability_type(impairment_type)

            if standardized_type in seen_types:
                continue

            seen_types.add(standardized_type)

            rows.append(
                {
                    "record_id": row.get("record_id"),
                    "interview_date": row.get("interview_date"),
                    "camp_location": row.get("camp_location"),
                    "helpdesk_location": row.get("helpdesk_location"),
                    "information_seeker_gender": row.get("information_seeker_gender"),
                    "age_group": row.get("age_group"),
                    "wgq_domain_column": column,
                    "wgq_disability_type": standardized_type,
                    "wgq_score": score,
                }
            )

    return pd.DataFrame(rows)


def adult_specific_disability_long(frame):
    rows = []

    if frame.empty:
        return pd.DataFrame()

    adult_frame = frame[frame["derived_life_stage"].astype(str).eq("Adult")].copy()

    for _, row in adult_frame.iterrows():
        for impairment_type in adult_row_impairment_types(row):
            rows.append(
                {
                    "record_id": row.get("record_id"),
                    "information_seeker_gender": row.get("information_seeker_gender"),
                    "specific_impairment_type": impairment_type,
                    "specific_disability_type": impairment_type,
                }
            )

    return pd.DataFrame(rows)


def adult_person_impairment_frame(frame):
    rows = []

    if frame.empty:
        return pd.DataFrame()

    adult_frame = frame[frame["derived_life_stage"].astype(str).eq("Adult")].copy()

    for _, row in adult_frame.iterrows():
        impairment_types = adult_row_impairment_types(row)
        impairment_count = len(impairment_types)

        if impairment_count == 0:
            disability_status = "No Disability"
            person_impairment_type = "No Disability"
        elif impairment_count == 1:
            disability_status = "Has Disability"
            person_impairment_type = impairment_types[0]
        else:
            disability_status = "Has Disability"
            person_impairment_type = "Multiple Impairments"

        if impairment_count == 0:
            impairment_count_category = "No Disability"
        elif impairment_count == 1:
            impairment_count_category = "One Impairment"
        elif impairment_count == 2:
            impairment_count_category = "Two Impairments"
        else:
            impairment_count_category = "Three or More Impairments"

        rows.append(
            {
                "record_id": row.get("record_id"),
                "information_seeker_gender": row.get("information_seeker_gender"),
                "adult_disability_status": disability_status,
                "adult_person_impairment_type": person_impairment_type,
                "adult_impairment_count": impairment_count,
                "adult_impairment_count_category": impairment_count_category,
                "adult_impairment_multiplicity": (
                    "No Disability"
                    if impairment_count == 0
                    else "Single Impairment"
                    if impairment_count == 1
                    else "Multiple Impairments"
                ),
                "duplicate_impairment_mentions": max(impairment_count - 1, 0),
            }
        )

    return pd.DataFrame(rows)


def adult_impairment_reconciliation(frame):
    adult_person_frame = adult_person_impairment_frame(frame)
    adult_specific_frame = adult_specific_disability_long(frame)

    if adult_person_frame.empty:
        return pd.DataFrame(columns=["Metric", "Value", "Interpretation"])

    specific_mentions = len(adult_specific_frame)
    duplicate_mentions = int(adult_person_frame["duplicate_impairment_mentions"].sum())
    unique_adults_with_impairment = int(
        adult_person_frame["adult_disability_status"].eq("Has Disability").sum()
    )
    adults_with_multiple_impairments = int(
        adult_person_frame["adult_impairment_count"].gt(1).sum()
    )
    reconciled_unique_count = specific_mentions - duplicate_mentions

    return pd.DataFrame(
        {
            "Metric": [
                "Specific impairment mentions",
                "Duplicate impairment mentions",
                "Reconciled unique adults with impairment",
                "Unique adults with impairment",
                "Adults with multiple impairments",
            ],
            "Value": [
                specific_mentions,
                duplicate_mentions,
                reconciled_unique_count,
                unique_adults_with_impairment,
                adults_with_multiple_impairments,
            ],
            "Interpretation": [
                "Every selected impairment type is counted once, so one person can contribute more than one mention.",
                "Extra mentions caused by adults who have more than one impairment type.",
                "Specific impairment mentions minus duplicate impairment mentions.",
                "Person-level adult disability prevalence used in the prevalence tables.",
                "Adults represented under Multiple Impairments in person-level type tables.",
            ],
        }
    )


def adult_specific_type_summary(frame):
    adult_frame = frame[frame["derived_life_stage"].astype(str).eq("Adult")].copy()
    total_adults = len(adult_frame)
    long_frame = adult_specific_disability_long(frame)

    if long_frame.empty:
        return pd.DataFrame(columns=["Impairment type", "Count", "Percentage"])

    summary = (
        long_frame.groupby("specific_impairment_type", dropna=False)
        .size()
        .reset_index(name="Count")
        .rename(columns={"specific_impairment_type": "Impairment type"})
        .sort_values("Count", ascending=False)
    )

    summary["Percentage"] = summary["Count"].map(
        lambda value: round((value / total_adults) * 100, 1) if total_adults else 0
    )

    return summary


def adult_additional_disability_category_summary(frame):
    adult_frame = frame[frame["derived_life_stage"].astype(str).eq("Adult")].copy()

    if adult_frame.empty or "adult_additional_disability_category" not in adult_frame.columns:
        return pd.DataFrame(columns=["Additional impairment category", "Count", "Percentage"])

    valid = adult_frame[
        adult_frame["adult_additional_disability_category"].ne("None")
    ].copy()

    if valid.empty:
        return pd.DataFrame(columns=["Additional impairment category", "Count", "Percentage"])

    summary = (
        valid.groupby("adult_additional_disability_category", dropna=False)
        .size()
        .reset_index(name="Count")
        .rename(columns={"adult_additional_disability_category": "Additional impairment category"})
        .sort_values("Count", ascending=False)
    )

    denominator = len(valid)
    summary["Percentage"] = summary["Count"].map(
        lambda value: round((value / denominator) * 100, 1) if denominator else 0
    )

    return summary


# Data loading, reshaping, and KPI preparation
@st.cache_data(show_spinner="Loading latest helpdesk dataset...")
def load_data(file_signature):
    if not DATA_FILE_PATH.exists():
        st.error(f"File not found: {DATA_FILE_PATH}")
        st.stop()

    records = pd.read_excel(DATA_FILE_PATH, sheet_name="cleaned_data")

    try:
        mapping = pd.read_excel(DATA_FILE_PATH, sheet_name="Column Mapping")
    except Exception:
        mapping = None

    required_columns = [
        "staff_name",
        "gps_latitude",
        "gps_longitude",
        "household_type",
        "helpdesk_camp_location",
        "helpdesk_village",
        "has_disability",
        "child_disability_type",
        "child_disability_type_other",
    ]

    required_columns.extend(WGQ_DISABILITY_DOMAINS.keys())
    required_columns.extend(ADULT_DISABILITY_CATEGORY_COLUMNS)

    for column in required_columns:
        if column not in records.columns:
            records[column] = pd.NA

    records["source_row_number"] = records.index + 2
    records["record_id"] = records["source_row_number"].map(lambda row: f"HD-{row:05d}")

    records["interview_date"] = pd.to_datetime(records["interview_date"], errors="coerce")

    if "referral_date" in records.columns:
        records["referral_date"] = pd.to_datetime(records["referral_date"], errors="coerce")

    records["year"] = records["interview_date"].dt.year
    records["month_number"] = records["interview_date"].dt.month
    records["year_month"] = records["interview_date"].dt.to_period("M").astype(str)
    records["month_label"] = records["interview_date"].dt.strftime("%b %Y")

    records["gps_latitude"] = pd.to_numeric(records["gps_latitude"], errors="coerce")
    records["gps_longitude"] = pd.to_numeric(records["gps_longitude"], errors="coerce")
    records["staff_name"] = records["staff_name"].map(clean_text)
    records["staff_name"] = records["staff_name"].fillna("[Not recorded]")

    records["household_type"] = records["household_type"].map(clean_text)
    records["age_group"] = records["information_seeker_age"].map(clean_text)
    records["derived_life_stage"] = records["age_group"].map(age_group_life_stage)

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

    records["request_category"] = records["request_type_protection_or_information"].map(clean_text)
    records["action_taken_clean"] = records["action_taken"].map(clean_text)
    records["follow_up_required_clean"] = records["follow_up_required"].map(clean_text)

    records["helpdesk_location"] = records.apply(derive_linked_helpdesk_location, axis=1)

    records["adult_wgq_disability_domains"] = records.apply(
        derive_adult_wgq_disability_domains,
        axis=1,
    )
    records["adult_additional_disability_category"] = records.apply(
        derive_adult_additional_disability_category,
        axis=1,
    )
    records["adult_wgq_disability_status"] = records.apply(
        derive_adult_wgq_disability_status,
        axis=1,
    )
    records["adult_wgq_disability_type"] = records.apply(
        derive_adult_wgq_disability_type,
        axis=1,
    )
    records["adult_wgq_domain_count"] = records.apply(
        derive_adult_wgq_domain_count,
        axis=1,
    )
    records["adult_wgq_impairment_count"] = records.apply(
        derive_adult_wgq_impairment_count,
        axis=1,
    )
    records["adult_duplicate_impairment_mentions"] = records.apply(
        derive_adult_duplicate_impairment_mentions,
        axis=1,
    )
    records["adult_wgq_domain_count_category"] = records.apply(
        derive_adult_wgq_domain_count_category,
        axis=1,
    )
    records["adult_wgq_multiplicity"] = records.apply(
        derive_adult_wgq_multiplicity,
        axis=1,
    )
    records["adult_wgq_max_score"] = records.apply(
        derive_adult_wgq_max_score,
        axis=1,
    )
    records["adult_wgq_severity"] = records.apply(
        derive_adult_wgq_severity,
        axis=1,
    )
    records["adult_disability_exclusion_risk"] = records.apply(
        derive_adult_disability_exclusion_risk,
        axis=1,
    )

    records["child_disability_status"] = records.apply(
        derive_child_disability_status,
        axis=1,
    )
    records["child_disability_type"] = records.apply(
        derive_child_disability_type,
        axis=1,
    )

    records["disability_status"] = records.apply(
        derive_combined_disability_status,
        axis=1,
    )
    records["disability_type"] = records.apply(
        derive_combined_disability_type,
        axis=1,
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
        "household_type",
        "staff_name",
        "gps_latitude",
        "gps_longitude",
        "information_seeker_type",
        "information_seeker_gender",
        "age_group",
        "derived_life_stage",
        "information_seeker_type_raw",
        "information_seeker_gender_raw",
        "type_age_correction_flag",
        "gender_age_correction_flag",
        "disability_status",
        "disability_type",
        "adult_wgq_disability_status",
        "adult_wgq_disability_type",
        "adult_wgq_disability_domains",
        "adult_wgq_domain_count",
        "adult_wgq_domain_count_category",
        "adult_wgq_multiplicity",
        "adult_wgq_max_score",
        "adult_wgq_severity",
        "adult_disability_exclusion_risk",
        "adult_additional_disability_category",
        "child_disability_status",
        "child_disability_type",
        "child_disability_type_other",
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
                "mapped_gps_records",
                "staff_recorded_records",
                "disability_records",
                "adult_disability_records",
                "child_disability_records",
                "adult_multiple_impairment_records",
                "gender_age_corrected_records",
                "type_age_corrected_records",
            ],
            "value": [
                len(dashboard_records),
                dashboard_records["request_category"].eq("Reporting a protection concern").sum(),
                dashboard_records["request_category"].eq(
                    "Seeking general protection information"
                ).sum(),
                dashboard_records["referral_status"].eq("Referred to partner agency").sum(),
                dashboard_records["follow_up_required_clean"].eq("Yes").sum(),
                dashboard_records[["gps_latitude", "gps_longitude"]].notna().all(axis=1).sum(),
                dashboard_records["staff_name"].ne("[Not recorded]").sum(),
                dashboard_records["disability_status"].eq("Has Disability").sum(),
                dashboard_records["adult_wgq_disability_status"].eq("Has Disability").sum(),
                dashboard_records["child_disability_status"].eq("Has Disability").sum(),
                dashboard_records["adult_wgq_multiplicity"].eq("Multiple Impairments").sum(),
                dashboard_records["gender_age_correction_flag"].sum(),
                dashboard_records["type_age_correction_flag"].sum(),
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


def basic_count_table(frame, category_column, category_label):
    if frame.empty or category_column not in frame.columns:
        return pd.DataFrame()

    table = (
        frame.groupby(category_column, dropna=False)
        .size()
        .reset_index(name="Records")
        .rename(columns={category_column: category_label})
        .sort_values("Records", ascending=False)
    )

    total = pd.DataFrame([{category_label: "Total", "Records": table["Records"].sum()}])
    return pd.concat([table, total], ignore_index=True)


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
    gps_columns = [
        col for col in display_table.columns
        if col in ["gps_latitude", "gps_longitude", "lat", "lon"]
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

    for col in numeric_columns:
        if col in gps_columns:
            formatters[col] = "{:,.6f}"
        elif "percentage" in str(col).lower():
            formatters[col] = "{:,.1f}"
        else:
            formatters[col] = "{:,.0f}"

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


def disability_gender_pivot_table(frame, category_column, category_label, top_n=None):
    required_columns = [category_column, "disability_status", "information_seeker_gender"]

    if frame.empty or any(column not in frame.columns for column in required_columns):
        return pd.DataFrame()

    working = frame[required_columns].copy()
    working[category_column] = working[category_column].map(clean_text)
    working["disability_status"] = working["disability_status"].map(clean_text)
    working["information_seeker_gender"] = working["information_seeker_gender"].map(clean_text)
    working = working.dropna(subset=[category_column, "disability_status", "information_seeker_gender"])

    if working.empty:
        return pd.DataFrame()

    grouped = (
        working.groupby(
            [category_column, "disability_status", "information_seeker_gender"],
            dropna=False,
        )
        .size()
        .reset_index(name="records")
    )

    if top_n:
        top_values = (
            grouped.groupby(category_column)["records"]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index
        )
        grouped = grouped[grouped[category_column].isin(top_values)]

    pivot = grouped.pivot_table(
        index=category_column,
        columns=["disability_status", "information_seeker_gender"],
        values="records",
        aggfunc="sum",
        fill_value=0,
    )

    status_order = ["Has Disability", "No Disability"]
    available_statuses = pivot.columns.get_level_values(0).unique().tolist()
    ordered_statuses = [status for status in status_order if status in available_statuses]
    ordered_statuses.extend(
        sorted(status for status in available_statuses if status not in ordered_statuses)
    )

    ordered_columns = []

    for status in ordered_statuses:
        available_genders = [
            gender
            for disability_status, gender in pivot.columns
            if disability_status == status
        ]
        ordered_genders = [gender for gender in GENDER_ORDER if gender in available_genders]
        ordered_genders.extend(
            sorted(gender for gender in available_genders if gender not in ordered_genders)
        )

        for gender in ordered_genders:
            ordered_columns.append((status, gender))

        status_columns = [(status, gender) for gender in ordered_genders]
        pivot[(status, "Total")] = pivot[status_columns].sum(axis=1)
        ordered_columns.append((status, "Total"))

    pivot[("Total", "Total")] = pivot[
        [(status, "Total") for status in ordered_statuses]
    ].sum(axis=1)
    ordered_columns.append(("Total", "Total"))

    pivot = pivot[ordered_columns]
    pivot = pivot.sort_values(("Total", "Total"), ascending=False)
    pivot.loc["Total"] = pivot.sum(axis=0)
    pivot.index.name = category_label
    pivot.columns = pd.MultiIndex.from_tuples(pivot.columns)

    return pivot.astype(int)


def style_disability_gender_table(table):
    def highlight_total_row(row):
        if row.name == "Total":
            return [
                "background-color: #DDEDE5; color: #102A2A; font-weight: 800;"
                for _ in row
            ]

        background = "#FFFFFF" if table.index.get_loc(row.name) % 2 == 0 else "#F7FAF8"
        return [f"background-color: {background};" for _ in row]

    def highlight_total_columns(column):
        if column.name[1] == "Total":
            return [
                "background-color: #FFF4D8; color: #102A2A; font-weight: 800;"
                for _ in column
            ]
        return ["" for _ in column]

    return (
        table.style
        .format("{:,.0f}")
        .apply(highlight_total_row, axis=1)
        .apply(highlight_total_columns, axis=0)
        .set_properties(
            **{
                "text-align": "center",
                "font-variant-numeric": "tabular-nums",
                "border": "1px solid #E5E7EB",
                "padding": "7px 9px",
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
                        ("text-align", "center"),
                        ("border", "1px solid #D8E2DC"),
                    ],
                },
            ]
        )
    )


def show_disability_gender_table(frame, category_column, category_label, top_n=None):
    table = disability_gender_pivot_table(
        frame,
        category_column,
        category_label,
        top_n=top_n,
    )

    if table.empty:
        st.info("No disability breakdown data match the selected filters.")
        return

    st.dataframe(
        style_disability_gender_table(table),
        use_container_width=True,
    )


def draw_disability_gender_bar(frame, category_column, top_n=None, height=430, disability_status_filter=None):
    required_columns = [category_column, "disability_status", "information_seeker_gender"]

    if frame.empty or any(column not in frame.columns for column in required_columns):
        st.info("No disability breakdown data match the selected filters.")
        return

    chart_data = frame[required_columns].copy()
    chart_data[category_column] = chart_data[category_column].map(clean_text)
    chart_data["disability_status"] = chart_data["disability_status"].map(clean_text)
    chart_data["information_seeker_gender"] = chart_data["information_seeker_gender"].map(clean_text)
    chart_data = chart_data.dropna(subset=required_columns)

    if disability_status_filter:
        chart_data = chart_data[
            chart_data["disability_status"].astype(str).eq(disability_status_filter)
        ].copy()

    if chart_data.empty:
        st.info("No disability breakdown data match the selected filters.")
        return

    grouped = (
        chart_data.groupby(
            [category_column, "disability_status", "information_seeker_gender"],
            dropna=False,
        )
        .size()
        .reset_index(name="Records")
    )

    if top_n:
        top_values = (
            grouped.groupby(category_column)["Records"]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index
        )
        grouped = grouped[grouped[category_column].isin(top_values)]

    if grouped.empty:
        st.info("No disability breakdown data match the selected filters.")
        return

    category_order = (
        grouped.groupby(category_column)["Records"]
        .sum()
        .sort_values(ascending=True)
        .index.astype(str)
        .tolist()
    )

    status_values = grouped["disability_status"].astype(str).unique().tolist()
    status_order = [
        status
        for status in ["Has Disability", "No Disability"]
        if status in status_values
    ]
    status_order.extend(sorted(status for status in status_values if status not in status_order))

    chart_height = max(240, min(520, 24 * len(category_order) + 70))

    base_chart = (
        alt.Chart(grouped)
        .mark_bar(cornerRadiusEnd=2, stroke="#FFFFFF", strokeWidth=0.5)
        .encode(
            y=alt.Y(
                f"{category_column}:N",
                sort=category_order,
                title=None,
                axis=alt.Axis(labelLimit=720, labelFontSize=11, labelPadding=6),
            ),
            x=alt.X("Records:Q", title="Records", stack="zero"),
            color=gender_color("information_seeker_gender:N"),
            tooltip=[
                alt.Tooltip(f"{category_column}:N", title="Category"),
                alt.Tooltip("disability_status:N", title="Disability status"),
                alt.Tooltip("information_seeker_gender:N", title="Gender"),
                alt.Tooltip("Records:Q", title="Records", format=","),
            ],
        )
        .properties(height=chart_height)
    )

    if len(status_order) == 1:
        chart = base_chart
    else:
        chart = base_chart.encode(
            row=alt.Row(
                "disability_status:N",
                sort=status_order,
                title="Disability status",
                header=alt.Header(labelColor="#12312F", labelFontSize=12, titleFontSize=12),
            )
        ).resolve_scale(x="shared", y="shared")

    st.altair_chart(polish_chart(chart), use_container_width=True)


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




def draw_total_donut(frame, category_column, category_label, height=320, min_label_share=0.04):
    if frame.empty or category_column not in frame.columns:
        st.info("No records match the selected filters.")
        return

    summary = (
        frame.groupby(category_column, dropna=False)
        .size()
        .reset_index(name="Records")
        .sort_values("Records", ascending=False)
    )

    if summary.empty or summary["Records"].sum() == 0:
        st.info("No summary data for the selected filters.")
        return

    summary[category_column] = summary[category_column].fillna("[Missing]").astype(str)
    summary["Share"] = summary["Records"] / summary["Records"].sum()
    summary["Share label"] = summary["Share"].map(
        lambda value: f"{value:.1%}" if value >= min_label_share else ""
    )

    donut = (
        alt.Chart(summary)
        .mark_arc(innerRadius=72, outerRadius=120, stroke="#FFFFFF", strokeWidth=2)
        .encode(
            theta=alt.Theta("Records:Q", stack=True),
            color=alt.Color(
                f"{category_column}:N",
                title=category_label,
                scale=alt.Scale(
                    range=["#2F7D69", "#D9A441", "#2563EB", "#DB2777", "#7C3AED", "#64748B"]
                ),
            ),
            tooltip=[
                alt.Tooltip(f"{category_column}:N", title=category_label),
                alt.Tooltip("Records:Q", title="Records", format=","),
                alt.Tooltip("Share:Q", title="Share", format=".1%"),
            ],
        )
    )

    labels = (
        alt.Chart(summary[summary["Share label"].ne("")])
        .mark_text(radius=145, fontSize=12, fontWeight=700, color="#334155")
        .encode(
            theta=alt.Theta("Records:Q", stack=True),
            text=alt.Text("Share label:N"),
        )
    )

    st.altair_chart(
        polish_chart((donut + labels).properties(height=height)),
        use_container_width=True,
    )


def draw_status_donut_pair(frame, status_column, height=300):
    if frame.empty or status_column not in frame.columns:
        st.info("No records match the selected filters.")
        return

    status_cols = st.columns(2)

    for column, status in zip(status_cols, ["No Disability", "Has Disability"]):
        with column:
            st.caption(status)
            status_frame = frame[frame[status_column].astype(str).eq(status)]
            draw_total_donut(
                status_frame,
                "information_seeker_gender",
                "Gender",
                height=height,
                min_label_share=0.06,
            )


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


def draw_count_bar(frame, category_column, category_label, height=360):
    if frame.empty or category_column not in frame.columns:
        st.info("No records match the selected filters.")
        return

    chart_data = (
        frame.groupby(category_column, dropna=False)
        .size()
        .reset_index(name="Records")
        .rename(columns={category_column: category_label})
        .sort_values("Records", ascending=False)
    )

    chart_data["axis_label"] = chart_data[category_label].map(
        lambda value: short_axis_label(value, max_chars=24)
    )

    chart = (
        alt.Chart(chart_data)
        .mark_bar(cornerRadiusEnd=2, color="#2F7D69")
        .encode(
            x=alt.X(
                "axis_label:N",
                sort=chart_data["axis_label"].tolist(),
                title=None,
                axis=alt.Axis(labelAngle=-30, labelLimit=160, labelFontSize=10),
            ),
            y=alt.Y("Records:Q", title="Records"),
            tooltip=[
                alt.Tooltip(f"{category_label}:N", title=category_label),
                alt.Tooltip("Records:Q", title="Records", format=","),
            ],
        )
        .properties(height=height)
    )

    st.altair_chart(polish_chart(chart), use_container_width=True)


# Map, insight cards, empty states, footer, and record search
def map_data(frame):
    if frame.empty:
        return pd.DataFrame()

    required_columns = ["gps_latitude", "gps_longitude"]

    if not all(col in frame.columns for col in required_columns):
        return pd.DataFrame()

    mapped = frame.dropna(subset=required_columns).copy()
    mapped = mapped[
        mapped["gps_latitude"].between(-90, 90)
        & mapped["gps_longitude"].between(-180, 180)
    ]

    if mapped.empty:
        return pd.DataFrame()

    mapped = mapped.rename(columns={"gps_latitude": "lat", "gps_longitude": "lon"})
    return mapped


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
            <div style="font-family: monospace; font-style: italic;">
                Developed by John Kul, MEAL Officer-Tdh.
            </div>
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

file_signature = data_file_signature(DATA_FILE_PATH)
records, protection, information, referrals, kpis = load_data(file_signature)

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
with st.sidebar:
    st.header("Filters")
    st.markdown(
        '<div class="section-note">Use the filters from top to bottom. Start with the date range, then select a camp to unlock linked location and demographic filters.</div>',
        unsafe_allow_html=True,
    )

    action_cols = st.columns(2)

    with action_cols[0]:
        if st.button(
            "Refresh data",
            use_container_width=True,
            help="Reload the latest Excel dataset and clear Streamlit cached data.",
        ):
            st.cache_data.clear()
            st.rerun()

    with action_cols[1]:
        st.button(
            "Clear filters",
            use_container_width=True,
            help="Reset the date range and all filter selections. This does not reload the dataset.",
            on_click=reset_filters,
            args=(default_from_date, max_date),
        )

    st.caption("Refresh data: Updates the source dataset.")
    st.caption("Clear filters: Resets your current selections.")

    with st.expander("Date range", expanded=True):
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

    selected_helpdesk_locations = []
    selected_information_seeker_types = []
    selected_genders = []
    selected_age_groups = []
    selected_request_categories = []

    with st.expander("Location filters", expanded=True):
        camp_options = filter_options(date_filtered_records["camp_location"])
        sanitize_multiselect_state("camp_location_filter", camp_options)

        selected_camp_locations = st.multiselect(
            "Camp location",
            camp_options,
            key="camp_location_filter",
        )

        if not selected_camp_locations:
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

    if not selected_camp_locations:
        for key in [
            "helpdesk_location_filter",
            "information_seeker_type_filter",
            "information_seeker_gender_filter",
            "age_group_filter",
            "request_category_filter",
        ]:
            st.session_state[key] = []

    else:
        if selected_helpdesk_locations:
            helpdesk_filtered_records = camp_filtered_records[
                camp_filtered_records["helpdesk_location"]
                .astype(str)
                .isin(selected_helpdesk_locations)
            ].copy()
        else:
            helpdesk_filtered_records = camp_filtered_records.copy()

        with st.expander("Demographic filters", expanded=True):
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

        with st.expander("Request filters", expanded=True):
            request_options = filter_options(age_filtered_records["request_category"])
            sanitize_multiselect_state("request_category_filter", request_options)

            selected_request_categories = st.multiselect(
                "Request category",
                request_options,
                key="request_category_filter",
            )

    with st.expander("Dataset status", expanded=False):
        st.caption(f"File modified: {file_signature[3] if file_signature[3] else 'Not found'}")
        st.caption(
            f"File size: {format_number(file_signature[2]) if file_signature[2] else 'Not found'} bytes"
        )
        st.caption(f"Loaded records: {format_number(len(records))}")
        st.caption(
            f"Data range: {min_date.strftime('%d %b %Y')} to {max_date.strftime('%d %b %Y')}"
        )

    with st.expander("Current selection", expanded=False):
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
disability_records = filtered_records["disability_status"].eq("Has Disability").sum()


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

kpi_cols = st.columns(6)
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
kpi_cols[5].metric(
    "Disability records",
    format_number(disability_records),
    f"{format_rate(disability_records, total_records)} of records",
)

if filtered_records.empty:
    show_empty_state(from_date, to_date)
    show_footer()
    st.stop()


# Quick insight cards
disability_type_records = filtered_records[
    filtered_records["disability_status"].eq("Has Disability")
]

top_location, top_location_count = top_value(filtered_records, "helpdesk_location")
top_request, top_request_count = top_value(filtered_records, "request_category")
top_concern, top_concern_count = top_value(filtered_protection, "protection_concern")
top_disability, top_disability_count = top_value(disability_type_records, "disability_type")

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
    "Top impairment type",
    top_disability,
    f"{format_number(top_disability_count)} records",
)

st.divider()


# Dashboard tabs
tab_overview, tab_disability, tab_concerns, tab_information, tab_referrals, tab_map, tab_data = st.tabs(
    ["Overview", "Disability", "Concerns", "Information", "Referrals", "Map", "Records"]
)


# Overview tab
with tab_overview:
    left, right = st.columns([1.2, 1])

    with left:
        st.subheader("Monthly Requests by gender")
        st.markdown(
            '<div class="section-note">Records grouped by interview month and gender.</div>',
            unsafe_allow_html=True,
        )
        draw_monthly_gender_column_bar(filtered_records, height=340)

    with right:
        st.subheader("Request type by information")
        st.markdown(
            '<div class="section-note">Total records by request type.</div>',
            unsafe_allow_html=True,
        )
        draw_total_donut(filtered_records, "request_category", "Request type", height=340)

    st.subheader("Request Type Table")
    show_gender_table(filtered_records, "request_category", "Request type")

    st.divider()

    st.subheader("Demographics by gender")
    st.caption("Information seeker type")
    draw_gender_column_bar(filtered_records, "information_seeker_type", height=300)
    show_gender_table(filtered_records, "information_seeker_type", "Information seeker type")

    st.markdown("#### Age group")
    st.caption("Age-group distribution by gender")
    draw_gender_column_bar(filtered_records, "age_group", height=420)
    show_gender_table(filtered_records, "age_group", "Age group")

    st.divider()

    st.subheader("Location by gender")
    st.caption("Camp location")
    draw_gender_column_bar(filtered_records, "camp_location", height=320)
    show_gender_table(filtered_records, "camp_location", "Camp location")

    st.markdown("#### Helpdesk location")
    st.caption("Helpdesk-level distribution by gender")
    draw_gender_column_bar(filtered_records, "helpdesk_location", height=460)
    show_gender_table(filtered_records, "helpdesk_location", "Helpdesk location")


# Protection concerns tab
with tab_concerns:
    st.subheader("Top Protection Concerns by gender")

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

    st.subheader("Protection Concerns by Disability Status and Gender")
    st.markdown(
        '<div class="section-note">Rows are protection concerns. Columns first split by disability status, then by gender.</div>',
        unsafe_allow_html=True,
    )
    concern_disability_top_n = st.radio(
        "Show top concerns by disability status",
        [10, 15, 25],
        horizontal=True,
        index=1,
        key="concern_disability_top_n",
    )
    draw_disability_gender_bar(
        filtered_protection,
        "protection_concern",
        top_n=concern_disability_top_n,
        height=520,
        disability_status_filter="Has Disability",
    )
    show_disability_gender_table(
        filtered_protection,
        "protection_concern",
        "Protection concern",
        top_n=concern_disability_top_n,
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

    st.subheader("Information Needs by Disability Status and Gender")
    st.markdown(
        '<div class="section-note">Rows are information needs. Columns first split by disability status, then by gender.</div>',
        unsafe_allow_html=True,
    )
    information_disability_top_n = st.radio(
        "Show top information needs by disability status",
        [10, 15, 25],
        horizontal=True,
        index=1,
        key="information_disability_top_n",
    )
    draw_disability_gender_bar(
        filtered_information,
        "general_information_need",
        top_n=information_disability_top_n,
        height=520,
        disability_status_filter="Has Disability",
    )
    show_disability_gender_table(
        filtered_information,
        "general_information_need",
        "General information need",
        top_n=information_disability_top_n,
    )


# Referrals tab
with tab_referrals:
    st.subheader("Action and Follow-up by Gender")

    st.caption("Referral status")
    draw_gender_column_bar(filtered_records, "referral_status", height=360)
    show_gender_table(filtered_records, "referral_status", "Referral status")

    st.markdown("#### Follow-up required")
    st.caption("Follow-up requirement by gender")
    draw_gender_column_bar(filtered_records, "follow_up_required_clean", height=360)
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


# Disability tab
with tab_disability:
    st.subheader("Disability Analysis")
    st.markdown(
        '<div class="section-note">Status labels remain Has Disability / No Disability. Specific types are standardized as impairments and shared across adults and children.</div>',
        unsafe_allow_html=True,
    )

    st.caption("Overall disability status by gender")
    draw_status_donut_pair(filtered_records, "disability_status", height=300)
    show_gender_table(filtered_records, "disability_status", "Disability status")

    st.markdown("#### Overall impairment type by gender")
    st.caption("Only records classified as Has Disability are included in this impairment type view.")
    disability_type_records = filtered_records[
        filtered_records["disability_status"].eq("Has Disability")
    ]
    draw_gender_column_bar(disability_type_records, "disability_type", height=430)
    show_gender_table(disability_type_records, "disability_type", "Impairment type")

    st.divider()

    st.subheader("Adult Disability Prevalence")
    adult_records = filtered_records[filtered_records["derived_life_stage"].eq("Adult")]
    adult_person_disability = adult_person_impairment_frame(adult_records)

    st.caption("Adult disability status by gender")
    draw_status_donut_pair(adult_person_disability, "adult_disability_status", height=300)
    show_gender_table(
        adult_person_disability,
        "adult_disability_status",
        "Adult disability status",
    )

    st.markdown("#### Adult impairment type summary by gender")
    st.caption("Adults with multiple impairments are represented as Multiple Impairments in this person-level summary.")
    adult_person_impairment_records = adult_person_disability[
        adult_person_disability["adult_disability_status"].eq("Has Disability")
    ]
    draw_gender_column_bar(
        adult_person_impairment_records,
        "adult_person_impairment_type",
        height=430,
    )
    show_gender_table(
        adult_person_impairment_records,
        "adult_person_impairment_type",
        "Adult impairment type",
    )

    st.divider()

    st.subheader("Adult Specific Impairment Breakdown")
    st.markdown(
        '<div class="section-note">This breakdown counts each specific impairment type. Adults with multiple impairments may appear in more than one type.</div>',
        unsafe_allow_html=True,
    )

    adult_specific_impairment_records = adult_specific_disability_long(adult_records)

    draw_gender_bar(
        adult_specific_impairment_records,
        "specific_impairment_type",
        height=420,
    )
    show_gender_table(
        adult_specific_impairment_records,
        "specific_impairment_type",
        "Specific impairment type",
    )

    st.subheader("Adult Specific Impairment Type Prevalence")
    st.dataframe(
        style_records_table(adult_specific_type_summary(adult_records)),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    st.subheader("Adult Multiple Impairment Analysis")
    st.markdown(
        '<div class="section-note">The donut chart uses total adult counts. The table retains the gender breakdown for detail.</div>',
        unsafe_allow_html=True,
    )
    st.caption("Single vs multiple impairments")
    draw_total_donut(
        adult_person_disability,
        "adult_impairment_multiplicity",
        "Number of impairment types",
        height=320,
    )
    show_gender_table(
        adult_person_disability,
        "adult_impairment_multiplicity",
        "Number of impairment types",
    )

    st.subheader("Adult Impairment Reconciliation")
    st.markdown(
        '<div class="section-note">Specific impairment mentions minus duplicate impairment mentions equals the unique adult disability prevalence count.</div>',
        unsafe_allow_html=True,
    )
    st.dataframe(
        style_records_table(adult_impairment_reconciliation(adult_records)),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()

    st.subheader("Child Disability")
    child_records = filtered_records[filtered_records["derived_life_stage"].eq("Child")]

    st.caption("Child disability status")
    draw_status_donut_pair(child_records, "child_disability_status", height=300)
    show_gender_table(child_records, "child_disability_status", "Child disability status")

    st.markdown("#### Child impairment type")
    st.caption("Only children classified as Has Disability are included in this impairment type view.")
    child_disability_records = child_records[
        child_records["child_disability_status"].eq("Has Disability")
    ]
    draw_gender_column_bar(child_disability_records, "child_disability_type", height=420)
    show_gender_table(child_disability_records, "child_disability_type", "Child impairment type")


# Map tab
with tab_map:
    st.subheader("Helpdesk Locations Map")
    st.markdown(
        '<div class="section-note">GPS coordinates represent helpdesk locations, not individual information seekers.</div>',
        unsafe_allow_html=True,
    )

    mapped_records = map_data(filtered_records)

    if mapped_records.empty:
        st.info("No valid GPS coordinates are available for the selected filters.")
    else:
        st.map(mapped_records[["lat", "lon"]], use_container_width=True)

        map_summary = (
            mapped_records.groupby(
                ["camp_location", "helpdesk_location", "lat", "lon"],
                dropna=False,
            )
            .size()
            .reset_index(name="records")
            .sort_values("records", ascending=False)
        )

        staff_summary = (
            filtered_records.groupby("staff_name", dropna=False)
            .size()
            .reset_index(name="records")
            .sort_values("records", ascending=False)
        )

        map_cols = st.columns(2)

        with map_cols[0]:
            st.subheader("Mapped Helpdesk Points")
            st.dataframe(
                style_records_table(map_summary),
                use_container_width=True,
                hide_index=True,
            )

        with map_cols[1]:
            st.subheader("Records by Staff")
            st.dataframe(
                style_records_table(staff_summary),
                use_container_width=True,
                hide_index=True,
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

