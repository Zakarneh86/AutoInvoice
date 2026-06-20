from functools import lru_cache
from pathlib import Path

import pandas as pd
from supabase import create_client


TABLES_DIR = Path("Tables")
TABLE_FILES = {
    "po_master": "po_master.csv",
    "po_daily_rates": "po_daily_rates.csv",
    "po_hourly_rates": "po_hourly_rates.csv",
    "po_working_hours": "po_working_hours.csv",
}


def has_supabase_config(secrets):
    return (
        "Supabase" in secrets
        and "url" in secrets["Supabase"]
        and "key" in secrets["Supabase"]
    )


@lru_cache(maxsize=1)
def get_supabase_client(url, key):
    return create_client(url, key)


def get_database_client(secrets):
    supabase_config = secrets["Supabase"]
    return get_supabase_client(supabase_config["url"], supabase_config["key"])


def read_csv_table(table_name):
    return pd.read_csv(TABLES_DIR / TABLE_FILES[table_name])


def write_csv_table(table_name, df):
    df.to_csv(TABLES_DIR / TABLE_FILES[table_name], index=False)


def read_supabase_table(table_name, secrets):
    client = get_database_client(secrets)
    response = client.table(table_name).select("*").execute()

    if response.data:
        return pd.DataFrame(response.data)

    return pd.DataFrame(columns=read_csv_table(table_name).columns)


def clean_records(df):
    clean_df = df.astype(object).where(pd.notna(df), None)
    return clean_df.to_dict(orient="records")


def write_supabase_table(table_name, df, secrets):
    client = get_database_client(secrets)
    client.table(table_name).delete().neq("po_number", "__never_match__").execute()

    records = clean_records(df)
    if records:
        client.table(table_name).insert(records).execute()


def read_table(table_name, secrets=None):
    if secrets is not None and has_supabase_config(secrets):
        return read_supabase_table(table_name, secrets)
    return read_csv_table(table_name)


def write_table(table_name, df, secrets=None):
    if secrets is not None and has_supabase_config(secrets):
        write_supabase_table(table_name, df, secrets)
    else:
        write_csv_table(table_name, df)


def get_orders_data(secrets=None):
    return (
        read_table("po_master", secrets),
        read_table("po_working_hours", secrets),
        read_table("po_daily_rates", secrets),
        read_table("po_hourly_rates", secrets),
    )


def align_to_columns(df, columns):
    aligned = df.copy()
    for column in columns:
        if column not in aligned.columns:
            aligned[column] = None
    return aligned[columns]


def upsert_by_po_number(table_name, new_rows, secrets=None):
    existing_rows = read_table(table_name, secrets)
    new_rows = align_to_columns(new_rows, existing_rows.columns)

    if new_rows.empty:
        return

    po_numbers = set(new_rows["po_number"].astype(str))

    if secrets is not None and has_supabase_config(secrets):
        client = get_database_client(secrets)
        delete_values = new_rows["po_number"].dropna().unique().tolist()
        if delete_values:
            client.table(table_name).delete().in_("po_number", delete_values).execute()

        records = clean_records(new_rows)
        if records:
            client.table(table_name).insert(records).execute()
        return

    existing_rows = existing_rows[
        ~existing_rows["po_number"].astype(str).isin(po_numbers)
    ]

    updated_rows = pd.concat([existing_rows, new_rows], ignore_index=True)
    write_table(table_name, updated_rows, secrets)


def save_po_tables(
    po_master_df,
    daily_rates_df,
    hourly_rates_df,
    working_hours_df,
    secrets=None,
):
    upsert_by_po_number("po_master", po_master_df, secrets)
    upsert_by_po_number("po_daily_rates", daily_rates_df, secrets)
    upsert_by_po_number("po_hourly_rates", hourly_rates_df, secrets)
    upsert_by_po_number("po_working_hours", working_hours_df, secrets)
