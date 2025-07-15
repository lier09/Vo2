# -*- coding: utf-8 -*-
"""
名称：最大 / 峰值氧摄取量测定与平台判定 & %VO₂AUC 数据准备
版本：最终部署版（Streamlit Cloud & GitHub 兼容）
"""

import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
import os
from PIL import Image
import plotly.express as px

# ----------------- 图标加载 -----------------
current_dir = os.path.dirname(__file__)
logo_path = os.path.join(current_dir, "首体院圆标校徽png图可直接使用.png")
logo_img = Image.open(logo_path)

# ----------------- 页面配置 -----------------
st.set_page_config(page_title="峰值氧摄取量测定与平台判定", layout="wide", page_icon="📈")
st.image(logo_img, width=100)
st.markdown("### Developed by Dr. Huzepeng, Capital University of Physical Education and Sports")

# ----------------- 核心参数 -----------------
RAW_TIME_COL = "t"
COLS_TO_PROCESS = ["V'O2", "V'CO2", "V'E", "HR", "VT", "BF"]
ROLL_WINDOW_ROWS = 3
VO2_PLATEAU_THRESHOLD = 0.15  # L/min

# ----------------- 文件上传 -----------------
uploaded_file = st.sidebar.file_uploader("上传 MetaLyzer Excel 文件", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0, skiprows=[1])
    df.columns = [str(c).strip() for c in df.columns]

    # 处理时间列
    if RAW_TIME_COL in df.columns:
        if df[RAW_TIME_COL].dtype == object or isinstance(df[RAW_TIME_COL].iloc[0], str):
            def parse_time(s):
                try:
                    parts = s.split(":")
                    return int(parts[0]) * 60 + int(parts[1])
                except:
                    return np.nan
            df[RAW_TIME_COL] = df[RAW_TIME_COL].apply(parse_time)
            df[RAW_TIME_COL].ffill(inplace=True)
    else:
        df[RAW_TIME_COL] = np.arange(len(df)) * 10

    st.success("✅ 数据读取完成，已自动处理时间列")

    # ----------------- 平滑处理 -----------------
    df_smoothed = df.copy()
    for col in COLS_TO_PROCESS:
        if col in df_smoothed.columns:
            df_smoothed[f"{col}_30s"] = df_smoothed[col].rolling(window=ROLL_WINDOW_ROWS, min_periods=ROLL_WINDOW_ROWS).mean()

    # 重新计算派生指标
    if "BodyMass" in df.columns and "V'O2" in df.columns:
        df_smoothed["V'O2/kg_30s"] = df_smoothed["V'O2_30s"] / df["BodyMass"] * 1000

    # VO₂ 平台判定
    vo2_series = df_smoothed["V'O2_30s"].dropna()
    vo2_diff = vo2_series.diff()
    is_plateau = (vo2_diff.iloc[-2:] < (VO2_PLATEAU_THRESHOLD / 6)).all() if len(vo2_diff) >= 2 else False

    # 计算关键指标
    vo2_peak = df_smoothed["V'O2_30s"].max()
    vo2_peak_kg = df_smoothed["V'O2/kg_30s"].max() if "V'O2/kg_30s" in df_smoothed.columns else np.nan
    ve_max = df_smoothed["V'E_30s"].max() if "V'E_30s" in df_smoothed.columns else np.nan
    hr_max = df_smoothed["HR_30s"].max() if "HR_30s" in df_smoothed.columns else np.nan
    rer_end = df_smoothed["V'CO2_30s"].tail(3).mean() / df_smoothed["V'O2_30s"].tail(3).mean() if "V'CO2_30s" in df_smoothed.columns else np.nan

    st.subheader("📊 关键指标")
    st.metric("VO₂ 平台状态", "达到" if is_plateau else "未达到")
    st.metric("VO₂max/peak (L/min)", f"{vo2_peak:.2f}")
    st.metric("VO₂max/peak (mL/kg/min)", f"{vo2_peak_kg:.1f}" if not np.isnan(vo2_peak_kg) else "N/A")
    st.metric("VEmax (L/min)", f"{ve_max:.1f}" if not np.isnan(ve_max) else "N/A")
    st.metric("HRmax (bpm)", int(hr_max) if not np.isnan(hr_max) else "N/A")
    st.metric("RER (末3点平均)", f"{rer_end:.2f}" if not np.isnan(rer_end) else "N/A")

    # ----------------- 趋势图 -----------------
    st.subheader("📈 30秒平滑后趋势图")
    smoothed_cols = [c for c in df_smoothed.columns if "_30s" in c]
    selected_cols = st.multiselect("选择要绘制的指标", smoothed_cols, default=["V'O2_30s", "V'CO2_30s", "HR_30s"])

    if selected_cols:
        fig = px.line(df_smoothed, x=RAW_TIME_COL, y=selected_cols, title="平滑数据趋势图")
        st.plotly_chart(fig, use_container_width=True)

    # ----------------- 下载按钮 -----------------
    excel_io = BytesIO()
    with pd.ExcelWriter(excel_io, engine="xlsxwriter") as writer:
        df_smoothed.to_excel(writer, index=False, sheet_name="Smoothed_Data")
    excel_io.seek(0)

    st.download_button(
        label="💾 下载平滑后数据 (.xlsx)",
        data=excel_io,
        file_name="smoothed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("请在左侧上传您的 Excel 文件以开始分析。")

