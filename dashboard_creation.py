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

# place this new table with additional columns into a new sheet
sheet_name = "Comparison & Overdue Cal"
with pd.ExcelWriter(f"{new_file_name}.xlsx", engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
    sheet = sheet_name
    df.to_excel(writer, sheet_name=sheet,index = False)
    ws = writer.sheets[sheet]

    header_row = 1
    first_data_row = header_row+1
    last_data_row = ws.max_row
    start_col = 1
    last_col = df.shape[1]

    ws.auto_filter.ref = f"{get_column_letter(start_col)}{header_row}:{get_column_letter(last_col)}{last_data_row}"

sheet_name = "Overdue Vulnerabilities"
with pd.ExcelWriter(f"{new_file_name}.xlsx", engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
    sheet = sheet_name
    overdue_df.to_excel(writer, sheet_name=sheet,index = False)
    ws = writer.sheets[sheet]

    header_row = 1
    first_data_row = header_row+1
    last_data_row = ws.max_row
    start_col = 1
    last_col = overdue_df.shape[1]

    ws.auto_filter.ref = f"{get_column_letter(start_col)}{header_row}:{get_column_letter(last_col)}{last_data_row}"

# Function creations
def merge_cells_title(sheet, cell1, cell2, row_no, col, val, horizontal_val, vertical_val, color):
    sheet.merge_cells(f'{cell1}:{cell2}')
    cell = sheet.cell(row=row_no, column=col)
    cell.value = val
    cell.alignment = Alignment(horizontal=horizontal_val, vertical=vertical_val)
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

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
    return pivot_table_df

def pivot_table_wide(df, index_col, col, val):
    pivot_df = pivot_table(df, index_col, col, val)
    pivot_df_wide = (
        pivot_df
        .pivot(index=index_col, columns=col, values="count")
        .fillna(0)
        .reset_index()
    )
    pivot_df_wide_rows = pivot_df_wide.shape[0]
    pivot_df_wide_columns = pivot_df_wide.shape[1]
    return pivot_df_wide, pivot_df_wide_rows, pivot_df_wide_columns

def barchart_creation(sheet, width, height, chartType, variable, x_title, y_title, min_col, max_col, min_row, max_row, showVal, showSerName, showCatName, showLeaderLines, cell, chartStyle=None, chartGrouping=None, chartOverlap=None, shape=None):
    chart = BarChart()
    chart.type = chartType
    if chartStyle is not None:
        chart.style = chartStyle
    if chartGrouping is not None:
        chart.grouping = chartGrouping
    if chartOverlap is not None:
        chart.overlap = chartOverlap
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
    if shape is not None:
        chart.shape = shape
    chart.width = width
    chart.height = height
    sheet.add_chart(chart, cell)
    return chart

def piechart_creation(sheet, shape, width, height, variable, min_col, min_row, max_row, max_col, showVal, showSerName, showCatName, showPercent, cell):
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
    sheet.add_chart(chart, cell)
    return chart

def main():
    family_vul, family_vul_rows, family_vul_columns = univariate_table(df, 'plugin_family')
    severity_vul, severity_vul_rows, severity_vul_columns = univariate_table(df, 'severity')
    asset_group_vul, asset_group_vul_rows, asset_group_vul_columns = univariate_table(df, 'asset_group')
    overdue_vul, overdue_vul_rows, overdue_vul_columns = univariate_table(df, 'overdue')
    pivot_table_1_wide, pivot_table_1_wide_rows, pivot_table_1_wide_columns = pivot_table_wide(df, "plugin_family", "severity", "plugin_id")
    pivot_table_2_wide, pivot_table_2_wide_rows, pivot_table_2_wide_columns = pivot_table_wide(df, "asset_group", "severity", "plugin_id")
    pivot_table_3_wide, pivot_table_3_wide_rows, pivot_table_3_wide_columns = pivot_table_wide(df, "overdue", "severity", "plugin_id")
    pivot_table_4_wide, pivot_table_4_wide_rows, pivot_table_4_wide_columns = pivot_table_wide(df, ["asset_group", "plugin_family"], "severity", "plugin_id")

    sheet_name = "Uni-Variate Summary Table"
    with pd.ExcelWriter(f"{new_file_name}.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        sheet = sheet_name
        new_row = 3
        family_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)
        new_row = new_row + family_vul_rows + 2
        severity_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)
        new_row = new_row + severity_vul_rows + 2
        asset_group_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)
        new_row = new_row + asset_group_vul_rows + 2
        overdue_vul.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)

    wb = load_workbook(f"{new_file_name}.xlsx")
    sheet = wb[sheet_name]

    merge_cells_title(sheet, "A1", "B2", 1, 1, "Summary Table", "center", "center", "00FFFF00")
    merge_cells_title(sheet, "D1", "AB2", 1, 4, "Summary Table Charts", "center", "center", '0000FF00')

    min_row_table = 4
    max_row_table = min_row_table + family_vul_rows
    family_vul_chart = barchart_creation(sheet, 22.5, 7, "col", "Family", "Family", "Count", 2, family_vul_columns, min_row_table, max_row_table, True, False, False, False, "D4", chartStyle=10, shape=4)
    family_vul_piechart = piechart_creation(sheet, 4, 22.5, 7, "Family", 2, min_row_table, max_row_table, family_vul_columns, False, False, False, True, "Q4")
    min_row_table = min_row_table + family_vul_rows + 2
    max_row_table = max_row_table + 2 + severity_vul_rows
    severity_vul_chart = barchart_creation(sheet, 22.5, 7, "col", "Severity", "Severity", "Count", 2, severity_vul_columns, min_row_table, max_row_table, True, False, False, False, "D21", chartStyle=10, shape=4)
    severity_vul_piechart = piechart_creation(sheet, 4, 22.5, 7, "Severity", 2, min_row_table, max_row_table, severity_vul_columns, False, False, False, True, "Q21")
    min_row_table = min_row_table + severity_vul_rows + 2
    max_row_table = max_row_table + 2 + asset_group_vul_rows
    asset_group_vul_chart = barchart_creation(sheet, 22.5, 7, "col", "Asset Group", "Asset Group", "Count", 2, asset_group_vul_columns, min_row_table, max_row_table, True, False, False, False, "D38", chartStyle=10, shape=4)
    asset_group_vul_piechart = piechart_creation(sheet, 4, 22.5, 7, "Asset Group", 2, min_row_table, max_row_table, asset_group_vul_columns, False, False, False, True, "Q38")
    min_row_table = min_row_table + asset_group_vul_rows + 2
    max_row_table = max_row_table + 2 + overdue_vul_rows
    overdue_vul_chart = barchart_creation(sheet, 22.5, 7, "col", "Overdue", "Overdue", "Count", 2, overdue_vul_columns, min_row_table, max_row_table, True, False, False, False, "D55", chartStyle=10, shape=4)
    overdue_vul_piechart = piechart_creation(sheet, 4, 22.5, 7, "Overdue", 2, min_row_table, max_row_table, overdue_vul_columns, False, False, False, True, "Q55")

    wb.save(f"{new_file_name}.xlsx")

    sheet_name = "Plugin Family & Severity"
    with pd.ExcelWriter(f"{new_file_name}.xlsx", engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
        sheet = sheet_name
        new_row = 3
        pivot_table_1_wide.to_excel(writer, sheet_name=sheet, startrow=new_row, startcol=0, index=False)
        ws = writer.sheets[sheet]

        header_row = new_row+1
        first_data_row = header_row+1
        last_data_row = ws.max_row
        start_col = 1
        last_col = pivot_table_1_wide_columns

        ws.auto_filter.ref = f"{get_column_letter(start_col)}{header_row}:{get_column_letter(last_col)}{last_data_row}"

    wb = load_workbook(f"{new_file_name}.xlsx")
    sheet = wb[sheet_name]
    merge_cells_title(sheet, "A1", f"{get_column_letter(pivot_table_1_wide_columns)}2", 1, 1, f"{sheet_name} Table", "center", "center", "00FFFF00")
    merge_cells_title(sheet, f"{get_column_letter(pivot_table_1_wide_columns+3)}1", "AH2", 1, pivot_table_1_wide_columns+3, "Summary Table Charts", "center", "center", '0000FF00')

    min_row_table = 4
    max_row_table = min_row_table + pivot_table_1_wide_rows
    pivot_table_1_wide_chart = barchart_creation(sheet, 50.5, 30, "col", "Severity and Plugin Family", "Family", "Count", 2, pivot_table_1_wide_columns, min_row_table, max_row_table, True, False, False, False, f"{get_column_letter(pivot_table_1_wide_columns+3)}4", chartGrouping="percentStacked", chartOverlap=100)

    wb.save(f"{new_file_name}.xlsx")

if __name__ == "__main__":
    main()