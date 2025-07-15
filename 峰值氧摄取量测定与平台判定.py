# -*- coding: utf-8 -*-
"""
峰值氧摄取量测定与平台判定 & %VO₂AUC 数据准备
最终定制版（胡泽鹏博士，首都体育学院）
"""

import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
from pathlib import Path
import plotly.express as px
import re
from PIL import Image

# ====== 常量设置 ======
RAW_TIME_COL = "t"
ROLL_WINDOW_ROWS = 30 // 10
VO2_PLATEAU_THRESHOLD = 0.15
COLS_TO_PROCESS = ["V'O2", "V'CO2", "V'E", "HR", "VT", "BF"]
CANONICAL_NAMES = {
    "VO2": "V'O2", "V'O₂": "V'O2",
    "VCO2": "V'CO2", "VCO₂": "V'CO2",
    "VE": "V'E", "V'E ": "V'E",
    "HeartRate": "HR", "HR (bpm)": "HR",
    "t": "t", "Time": "t",
    "体重": "BodyMass"
}

# ====== 文件读取与准备 ======
def read_and_prepare_data(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0, skiprows=[1])
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns=CANONICAL_NAMES, inplace=True)
    if RAW_TIME_COL not in df:
        df[RAW_TIME_COL] = np.arange(len(df)) * 10
    elif df[RAW_TIME_COL].dtype == object or isinstance(df[RAW_TIME_COL].iloc[0], str):
        def parse_time(s):
            parts = re.split('[:.]', s)
            try:
                if len(parts) >= 3:
                    h, m, sec = map(int, parts[:3])
                    return h * 3600 + m * 60 + sec
                elif len(parts) == 2:
                    m, sec = map(int, parts)
                    return m * 60 + sec
                else:
                    return float(s)
            except:
                return np.nan
        df[RAW_TIME_COL] = df[RAW_TIME_COL].apply(parse_time)
        df[RAW_TIME_COL].ffill(inplace=True)

    if "BodyMass" not in df:
        df["BodyMass"] = 70.0
    else:
        df["BodyMass"].ffill(inplace=True)
        df["BodyMass"].bfill(inplace=True)
    return df

# ====== 平滑处理 ======
def perform_30s_rolling_average(df):
    df_rolled = df.copy()
    for col in COLS_TO_PROCESS:
        if col in df.columns:
            df_rolled[f"{col}_30s"] = df[col].rolling(window=ROLL_WINDOW_ROWS, min_periods=ROLL_WINDOW_ROWS).mean()
    return df_rolled

# ====== 平台判定逻辑 ======
def get_key_metrics_summary(df_smoothed):
    vo2_col = "V'O2_30s"
    vo2_kg_col = "V'O2/kg_30s"
    ve_col = "V'E_30s"
    hr_col = "HR_30s"
    rer_col = "RER_30s"

    vo2_series = df_smoothed[vo2_col].dropna()
    vo2_diff = vo2_series.diff()
    is_plateau = (vo2_diff.iloc[-2:] < VO2_PLATEAU_THRESHOLD).all() if len(vo2_diff) >= 2 else False

    vo2_peak_value = vo2_series.max()
    peak_idx = vo2_series.idxmax()
    vo2_peak_kg_value = df_smoothed.loc[peak_idx, vo2_kg_col] if vo2_kg_col in df_smoothed.columns and pd.notna(peak_idx) else np.nan

    summary = {
        "摄氧量平台状态": f"✔ VO₂max ({'达到平台' if is_plateau else '未达到平台, 为VO₂peak'})",
        "VO₂max/peak (L·min⁻¹)": round(vo2_peak_value, 2) if pd.notna(vo2_peak_value) else 'N/A',
        "VO₂max/peak (mL·kg⁻¹·min⁻¹)": round(vo2_peak_kg_value, 1) if pd.notna(vo2_peak_kg_value) else 'N/A',
        "VEmax (L·min⁻¹)": round(df_smoothed["V'E_30s"].max(), 1) if "V'E_30s" in df_smoothed.columns else 'N/A',
        "HRmax (bpm)": int(df_smoothed["HR_30s"].max()) if "HR_30s" in df_smoothed.columns and pd.notna(df_smoothed["HR_30s"].max()) else 'N/A',
        "RER (末3点平均)": round(df_smoothed["RER_30s"].dropna().tail(3).mean(), 2) if "RER_30s" in df_smoothed.columns else 'N/A'
    }
    return summary

# ====== 百分比表 ======
def get_vo2_percentage_table(df_smoothed, vo2_peak_value):
    if vo2_peak_value is None or vo2_peak_value == 0 or pd.isna(vo2_peak_value):
        return pd.DataFrame()
    df = df_smoothed.copy()
    df["VO2_%_of_Peak"] = (df["V'O2_30s"] / vo2_peak_value) * 100
    result_rows = []
    for p in range(10, 101, 10):
        target_rows = df[df["VO2_%_of_Peak"] >= p]
        if not target_rows.empty:
            first_row = target_rows.iloc[[0]].copy()
            first_row["目标百分比"] = f"{p}%"
            result_rows.append(first_row)
    if not result_rows:
        return pd.DataFrame()
    final_table = pd.concat(result_rows)
    cols = list(final_table.columns)
    final_table = final_table[['目标百分比', 'VO2_%_of_Peak'] + [c for c in cols if c not in ['目标百分比', 'VO2_%_of_Peak']]]
    return final_table

# ====== Excel 导出 ======
def to_excel_bytes(dfs: dict):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for sheet_name, df_data in dfs.items():
            if isinstance(df_data, pd.DataFrame):
                df_data.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return bio.getvalue()

# ====== Streamlit 界面 ======
st.set_page_config(page_title="峰值氧摄取量分析", layout="wide", page_icon="📈")

# 左上角署名
st.markdown("### Developed by Dr. Huzepeng, Capital University of Physical Education and Sports")

# 加载并展示右上角校徽
logo_path = "C:/Users/huzepeng/PycharmProjects/首体院圆标校徽png图可直接使用.png"  # <-- 确认实际本地路径
logo_img = Image.open(logo_path)
st.image(logo_img, width=100)

uploaded_file = st.sidebar.file_uploader("上传 Excel 原始数据", type=["xlsx", "xls"])

if uploaded_file:
    df_raw = read_and_prepare_data(uploaded_file)
    smoothed_df = perform_30s_rolling_average(df_raw)

    smooth_cols = [f"{c}_30s" for c in COLS_TO_PROCESS if f"{c}_30s" in smoothed_df.columns]
    original_cols = [col for col in df_raw.columns if col not in COLS_TO_PROCESS]
    final_df = smoothed_df[smooth_cols + original_cols]

    final_df["V'O2/kg_30s"] = final_df["V'O2_30s"] / final_df["BodyMass"] * 1000
    final_df["RER_30s"] = final_df["V'CO2_30s"] / final_df["V'O2_30s"]

    st.subheader("✅ 平滑后数据（仅显示平滑列 + 原表其它列）")
    st.dataframe(final_df)

    summary = get_key_metrics_summary(final_df)
    st.subheader("📊 关键指标")
    cols = st.columns(len(summary))
    for i, (k, v) in enumerate(summary.items()):
        cols[i].metric(label=k, value=v)

    pct_df = get_vo2_percentage_table(final_df, summary.get("VO₂max/peak (L·min⁻¹)"))
    if not pct_df.empty:
        st.subheader("🟠 %VO₂max 表格")
        st.dataframe(pct_df)

    st.subheader("📈 趋势图（可选择多个指标）")
    plot_cols = st.multiselect("选择指标：", smooth_cols, default=["V'O2_30s", "V'CO2_30s"])
    if plot_cols:
        fig = px.line(final_df, x=RAW_TIME_COL, y=plot_cols, title="平滑后趋势图")
        st.plotly_chart(fig, use_container_width=True)

    excel_bytes = to_excel_bytes({
        "Smoothed_Data": final_df,
        "VO2_Percentage_Table": pct_df
    })

    st.download_button(
        label="💾 下载最终完整结果 (.xlsx)",
        data=excel_bytes,
        file_name=f"{Path(uploaded_file.name).stem}_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("请在左侧上传 Excel 文件以开始分析。")
