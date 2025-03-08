import streamlit as st
import openpyxl
from openpyxl import Workbook
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO
from tqdm import tqdm, trange

wb = openpyxl.load_workbook('template.xlsx')

@st.cache_data()
def exec(aco, year):
    global data, wb
    df = data[data["ACO_Name"] == aco].reset_index(inplace = False, drop=True)
    df = df.replace(r'\*', '0', regex=True)
    df = df.replace(r'\-', '0', regex = True)

    data_map = {}
    sheet = "Cover"
    wb[sheet]["A13"] = f"{df.iloc[0]["ACO_ID"]}, {df.iloc[0]["ACO_Name"]}"
    wb[sheet]["A14"] = f"Financial Reconciliation Report, Performance Year {year}"
    def convert_date(date_str):
        date_string = "07/01/2019"
        input_format = "%m/%d/%Y"
        date_object = datetime.strptime(date_string, input_format)
        output_format = "%B %d, %Y"
        output_string = date_object.strftime(output_format)

        return output_string
    wb[sheet]["A16"] = f"{convert_date(df.iloc[0]["Current_Start_Date"])} Agreement Start Date"

    sheet = "Table 1 - Historical Benchmark"
    data_map[(sheet,"B10", int)] = "Per_Capita_Exp_ALL_ESRD_BY1"
    data_map[(sheet,"C10", int)] = "Per_Capita_Exp_ALL_ESRD_BY2"
    data_map[(sheet,"D10", int)] = "Per_Capita_Exp_ALL_ESRD_BY3"

    data_map[(sheet,"B11", int)] = "Per_Capita_Exp_ALL_DIS_BY1"
    data_map[(sheet,"C11", int)] = "Per_Capita_Exp_ALL_DIS_BY2"
    data_map[(sheet,"D11", int)] = "Per_Capita_Exp_ALL_DIS_BY3"

    data_map[(sheet,"B12", int)] = "Per_Capita_Exp_ALL_AGDU_BY1"
    data_map[(sheet,"C12", int)] = "Per_Capita_Exp_ALL_AGDU_BY2"
    data_map[(sheet,"D12", int)] = "Per_Capita_Exp_ALL_AGDU_BY3"

    data_map[(sheet,"B13", int)] = "Per_Capita_Exp_ALL_AGND_BY1"
    data_map[(sheet,"C13", int)] = "Per_Capita_Exp_ALL_AGND_BY2"
    data_map[(sheet,"D13", int)] = "Per_Capita_Exp_ALL_AGND_BY3"

    data_map[(sheet,"B15", float)] = "CMS_HCC_RiskScore_ESRD_BY1"
    data_map[(sheet,"C15", float)] = "CMS_HCC_RiskScore_ESRD_BY2"
    data_map[(sheet,"D15", float)] = "CMS_HCC_RiskScore_ESRD_BY3"

    data_map[(sheet,"B16", float)] = "CMS_HCC_RiskScore_DIS_BY1"
    data_map[(sheet,"C16", float)] = "CMS_HCC_RiskScore_DIS_BY2"
    data_map[(sheet,"D16", float)] = "CMS_HCC_RiskScore_DIS_BY3"

    data_map[(sheet,"B17", float)] = "CMS_HCC_RiskScore_AGDU_BY1"
    data_map[(sheet,"C17", float)] = "CMS_HCC_RiskScore_AGDU_BY2"
    data_map[(sheet,"D17", float)] = "CMS_HCC_RiskScore_AGDU_BY3"

    data_map[(sheet,"B18", float)] = "CMS_HCC_RiskScore_AGND_BY1"
    data_map[(sheet,"C18", float)] = "CMS_HCC_RiskScore_AGND_BY2"
    data_map[(sheet,"D18", float)] = "CMS_HCC_RiskScore_AGND_BY3"

    sheet = "Table 2 - Updated Benchmark"
    data_map[(sheet, "D13", float)] = "CMS_HCC_RiskScore_ESRD_PY"
    data_map[(sheet, "D14", float)] = "CMS_HCC_RiskScore_DIS_PY"
    data_map[(sheet, "D15", float)] = "CMS_HCC_RiskScore_AGDU_PY"
    data_map[(sheet, "D16", float)] = "CMS_HCC_RiskScore_AGND_PY"

    sheet = "Table 3 - Shared Savings Losses"
    data_map[(sheet, "B6", int)] = "N_AB"
    data_map[(sheet, "B7", int)] = "N_AB_Year_PY"
    data_map[(sheet, "B9", int)] = "Per_Capita_Exp_ALL_ESRD_PY"
    data_map[(sheet, "B10", int)] = "Per_Capita_Exp_ALL_DIS_PY"
    data_map[(sheet, "B11", int)] = "Per_Capita_Exp_ALL_AGDU_PY"
    data_map[(sheet, "B12", int)] = "Per_Capita_Exp_ALL_AGND_PY"
    data_map[(sheet, "B25", str)] = "MinSavPerc"
    data_map[(sheet, "B32", str)] = "QualScore"
    data_map[(sheet, "B34", str)] = "FinalShareRate"

    sheet = "Table 1 - Historical Benchmark"
    s = 0
    for i in ["ESRD", "DIS", "AGED_Dual", "AGED_NonDual"]:
        s += float(df.iloc[0][f"N_AB_Year_{i}_BY3"])
    wb[sheet]["D40"] = float(df.iloc[0]["N_AB_Year_ESRD_BY3"]) / s
    wb[sheet]["D41"] = float(df.iloc[0]["N_AB_Year_DIS_BY3"]) / s
    wb[sheet]["D42"] = float(df.iloc[0]["N_AB_Year_AGED_Dual_BY3"]) / s
    wb[sheet]["D43"] = float(df.iloc[0]["N_AB_Year_AGED_NonDual_BY3"]) / s

    sheet = "Table 2 - Updated Benchmark"
    s = 0
    for i in ["ESRD", "DIS", "AGED_Dual", "AGED_NonDual"]:
        s += float(df.iloc[0][f"N_AB_Year_{i}_PY"])
    wb[sheet]["D38"] = float(df.iloc[0]["N_AB_Year_ESRD_PY"]) / s
    wb[sheet]["D39"] = float(df.iloc[0]["N_AB_Year_DIS_PY"]) / s
    wb[sheet]["D40"] = float(df.iloc[0]["N_AB_Year_AGED_Dual_PY"]) / s
    wb[sheet]["D41"] = float(df.iloc[0]["N_AB_Year_AGED_NonDual_PY"]) / s

    for key in data_map.keys():
        ws = wb[key[0]]
        ws[key[1]] = key[2](df.iloc[0][f"{data_map[(key)]}"])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read()
    # wb.save(f"template_res_{aco}_{year}.xlsx")
    # st.write("waooww")

def print_info(aco, year):
    data = pd.read_csv(f"PY_Financial_and_Quality_Results_{year}.csv")
    df = data[data["ACO_Name"] == aco].reset_index(inplace = False, drop=True)
    df = df.replace(r'\*', '0', regex=True)
    df = df.replace(r'\-', '0', regex = True)
    for col in df.columns:
        print(f"{col}                          {df.iloc[0][col]}")
    return

st.title("MSSP Data Gen")

year = st.selectbox('Choose a year:', ['2022', '2023', '2024', '2025', '2026'])
data = pd.read_csv(f"PY_Financial_and_Quality_Results_{year}.csv")
data["ID_Name"] = data.apply(lambda row: f"{row["ACO_ID"]} - {row["ACO_Name"]}", axis = 1)

aco =  st.selectbox("Choose an ACO: ", data["ID_Name"]).split(" - ")[-1]

if st.button('Analyse', key='exec_button', help='Click this button to proceed', type="primary"):
    st.write('You clicked the button!')
    excel_data = exec(aco, year)

    st.download_button(
        label='Download Modified Excel File',
        data=excel_data,
        file_name=f"template_res_{aco}_{year}.xlsx",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )