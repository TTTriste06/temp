import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    """
    将 DataFrame 中的三个主键列替换为新旧料号映射表中的新值，并对重复记录聚合（数值列求和）。

    参数:
    - df: 原始 DataFrame
    - mapping_df: 包含 ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
    - field_map: 当前表格中列名与映射字段的对应关系，如 {"规格": "产品型号", ...}
    - verbose: 是否输出替换信息

    返回:
    - 替换并聚合后的 DataFrame
    """
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    # 强制转换主键列为字符串
    for col in [spec_col, name_col, wafer_col]:
        df[col] = df[col].astype(str).str.strip()
    for col in ["旧规格", "旧品名", "旧晶圆品名"]:
        mapping_df[col] = mapping_df[col].astype(str).str.strip()

    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    try:
        df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

        matched = df_merged["新规格"].notna()
        match_count = matched.sum()
        unmatched_count = (~matched).sum()

        if verbose:
            msg = f"🎯 成功替换 {match_count} 行；未匹配 {unmatched_count} 行"
            try:
                import streamlit as st
                st.info(msg)
            except:
                print(msg)

        if unmatched_count > 0 and verbose:
            try:
                print("⚠️ 未匹配示例（前 5 行）：")
                print(df_merged[~matched][left_on].head())
            except:
                pass

        mask_None = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )
        
        df_merged.loc[mask_None, spec_col] = df_merged.loc[mask_None, "新规格"]
        df_merged.loc[mask_None, name_col] = df_merged.loc[mask_None, "新品名"]
        df_merged.loc[mask_None, wafer_col] = df_merged.loc[mask_None, "新晶圆品名"]

        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        group_cols = [spec_col, name_col, wafer_col]
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        return df_grouped

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return df
