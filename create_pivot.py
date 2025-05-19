
def create_pivot(df, config):
    from datetime import datetime, timedelta
    import pandas as pd
    from month_selector import process_history_columns

    def _excel_serial_to_date(serial):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(serial))
        except:
            return pd.NaT

    if "date_format" in config:
        date_col = config["columns"]
        if pd.api.types.is_numeric_dtype(df[date_col]):
            df[date_col] = df[date_col].apply(_excel_serial_to_date)
        else:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        new_col = f"{date_col}_年月"
        df[new_col] = df[date_col].dt.strftime(config["date_format"])
        df[new_col] = df[new_col].fillna("未知日期")
        config = config.copy()
        config["columns"] = new_col

    pivoted = pd.pivot_table(
        df,
        index=config["index"],
        columns=config["columns"],
        values=config["values"],
        aggfunc=config["aggfunc"],
        fill_value=0
    )

    pivoted.columns = [f"{col[0]}_{col[1]}" if isinstance(col, tuple) else str(col) for col in pivoted.columns]

    if pd.Series(pivoted.columns).duplicated().any():
        from pandas.io.parsers import ParserBase
        original_cols = pivoted.columns
        deduped_cols = ParserBase({'names': original_cols})._maybe_dedup_names(original_cols)
        pivoted.columns = deduped_cols

    pivoted = pivoted.reset_index()

    # 历史数据合并
    from config import CONFIG
    if CONFIG.get("selected_month") and config.get("values") and "未交订单数量" in config.get("values"):
        pivoted = process_history_columns(pivoted, config, CONFIG["selected_month"])

    return pivoted
