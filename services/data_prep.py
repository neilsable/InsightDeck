import pandas as pd

REQUIRED_COLS = {"day", "service", "usage_units", "cost_gbp", "incidents", "sla_pct"}

def load_and_validate_csv(csv_path) -> pd.DataFrame:
    df = pd.read_csv(csv_path)

    missing = REQUIRED_COLS - set(df.columns)
    if missing:
        raise ValueError(f"CSV missing required columns: {sorted(list(missing))}")

    # Basic typing cleanup
    df["day"] = pd.to_datetime(df["day"])
    df["service"] = df["service"].astype(str)
    df["usage_units"] = pd.to_numeric(df["usage_units"], errors="coerce").fillna(0).astype(int)
    df["cost_gbp"] = pd.to_numeric(df["cost_gbp"], errors="coerce").fillna(0.0).astype(float)
    df["incidents"] = pd.to_numeric(df["incidents"], errors="coerce").fillna(0).astype(int)
    df["sla_pct"] = pd.to_numeric(df["sla_pct"], errors="coerce").fillna(99.9).astype(float)

    return df
