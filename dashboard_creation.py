import pandas as pd
import numpy as np
import seaborn as sb
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference, LineChart, PieChart
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment  
from openpyxl.styles import PatternFill, Font
from datetime import datetime

today = datetime.today()
new_file_name = "dummy_data"
df = pd.read_excel(f"{new_file_name}_raw.xlsx")
severity_thresholds = {"Critical": 60, "High": 60, "Medium": 90, "Low": 90}

# time columns conversion
df['patch_publication_date'] = pd.to_datetime(df['patch_publication_date'], format = "%d/%m/%Y %H:%M:%S")
df['first_observed_date'] = pd.to_datetime(df['first_observed_date'], format="%d/%m/%Y %H:%M:%S")

# create a new column for today's date
df['comparison_date'] = pd.to_datetime(today, format="%d/%m/%Y %H:%M:%S")

# create a new column to calculate the difference in days
for idx, row in df.iterrows():
    if pd.isnull(row['patch_publication_date']):
        df.at[idx, "difference"] = (row['comparison_date'] - row['first_observed_date']).days
    else:
        df.at[idx, "difference"] = (row['comparison_date'] - row['patch_publication_date']).days

# create a new column to apply the overdue flah (Y or N)
df["overdue"] = df.apply(
    lambda row: "Y" if row["difference"] > severity_thresholds[row["severity"]] else "N",
    axis=1
)

# create a column to calculate the number of days overdue
df["overdue (days)"] = df.apply(
    lambda row: "" if row["overdue"] == "N" else row["difference"] - severity_thresholds[row["severity"]],
    axis=1
)

overdue_df = df.loc[df['overdue'] == 'Y'].reset_index()

# place this new table with additional columns into a new working sheet
with pd.ExcelWriter(f"{new_file_name}.xlsx", engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
    sheet = "Working_Sheet"
    df.to_excel(writer, sheet_name=sheet,index = False)
    ws = writer.sheets[sheet]

    header_row = 1
    first_data_row = header_row+1
    last_data_row = ws.max_row
    start_col = 1
    last_col = df.shape[1]

    ws.auto_filter.ref = f"{get_column_letter(start_col)}{header_row}:{get_column_letter(last_col)}{last_data_row}"


with pd.ExcelWriter(f"{new_file_name}.xlsx", engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
    sheet = "Overdue Vulnerabilities"
    overdue_df.to_excel(writer, sheet_name=sheet,index = False)
    ws = writer.sheets[sheet]

    header_row = 1
    first_data_row = header_row+1
    last_data_row = ws.max_row
    start_col = 1
    last_col = overdue_df.shape[1]

    ws.auto_filter.ref = f"{get_column_letter(start_col)}{header_row}:{get_column_letter(last_col)}{last_data_row}"

# Function creations
def univariate_table(df, variable):
    sub_df = df[variable].value_counts().reset_index()
    sub_df.columns = [variable, 'count']
    sub_df_rows = sub_df.shape[0]
    sub_df_columns = sub_df.shape[1]
    return sub_df, sub_df_rows, sub_df_columns

def pivot_table(df, index_col, col, val):
    pivot_table_df = pd.pivot_table(df, index=index_col, columns=col, values=val, aggfunc="count", fill_value=0).stack().reset_index()
    if isinstance(index_col, list):
        new_cols = index_col + [col, "count"]
    else:
        new_cols = [index_col, col, "count"]
    pivot_table_df.columns = new_cols
    pivot_table_df_rows = pivot_table_df.shape[0]
    pivot_table_df_columns = pivot_table_df.shape[1]
    return pivot_table_df, pivot_table_df_rows, pivot_table_df_columns

def barchart_creation(sheet, shape, width, height, chartType, chartStyle, variable, x_title, y_title, min_col, max_col, min_row, max_row, showVal, showSerName, showCatName, showLeaderLines):
    chart = BarChart()
    chart.type = chartType
    chart.style = chartStyle
    chart.title = f"Vulnerabilities by {variable}"
    chart.title.overlay = False
    chart.x_axis.title = x_title
    chart.y_axis.title = y_title
    chart.x_axis.overlay = False
    chart.y_axis.overlay = False
    data = Reference(sheet, min_col=min_col, min_row=min_row, max_row=max_row, max_col=max_col)
    cats = Reference(sheet, min_col=min_col-1, min_row=min_row+1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.dataLabels = DataLabelList() 
    chart.dataLabels.showVal = showVal
    chart.dataLabels.showSerName = showSerName
    chart.dataLabels.showCatName = showCatName
    chart.dataLabels.showLeaderLines = showLeaderLines
    chart.set_categories(cats)
    chart.shape = shape
    chart.width = width
    chart.height = height
    return chart

def piechart_creation(sheet, shape, width, height, variable, min_col, min_row, max_row, max_col, showVal, showSerName, showCatName, showPercent):
    chart = PieChart()
    chart.title = f"Vulnerabilities by {variable}"
    data = Reference(sheet, min_col=min_col, min_row=min_row, max_row=max_row, max_col=max_col)
    cats = Reference(sheet, min_col=min_col-1, min_row=min_row+1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.dataLabels = DataLabelList() 
    chart.dataLabels.showVal = showVal
    chart.dataLabels.showSerName = showSerName
    chart.dataLabels.showCatName = showCatName
    chart.dataLabels.showPercent = showPercent
    chart.set_categories(cats)
    chart.shape = shape
    chart.width = width
    chart.height = height
    return chart

def main():
    family_vul, family_vul_rows, family_vul_columns = univariate_table(df, 'plugin_family')
    severity_vul, severity_vul_rows, severity_vul_columns = univariate_table(df, 'severity')
    asset_group_vul, asset_group_vul_rows, asset_group_vul_columns = univariate_table(df, 'asset_group')
    overdue_vul, overdue_vul_rows, overdue_vul_columns = univariate_table(df, 'overdue')

    with pd.ExcelWriter(f"{new_file_name}.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        sheet = "Summary"
        family_vul.to_excel(writer, sheet_name=sheet, startrow=0, startcol=0, index=False)
        new_row = family_vul_rows + 2
        severity_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)
        new_row = new_row + severity_vul_rows + 2
        asset_group_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)
        new_row = new_row + asset_group_vul_rows + 2
        overdue_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)

    wb = load_workbook(f"{new_file_name}.xlsx")
    sheet = wb['Summary']

    family_vul_chart = barchart_creation(sheet, 4, 22.5, 10, "col", 10, "Family", "Family", "Count", 2, family_vul_columns, 1, family_vul_rows+1, True, False, False, False)
    sheet.add_chart(family_vul_chart, "D1")

    wb.save(f"{new_file_name}.xlsx")

if __name__ == "__main__":
    main()