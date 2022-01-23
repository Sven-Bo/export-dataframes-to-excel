import pandas as pd  # pip install pandas openpyxl

# Create Dataframes
df1 = pd.read_csv("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/tips.csv")
df2 = df1.groupby(by="sex").sum()[["tip"]]
df3 = df1.groupby(by="day").sum()[["tip"]]

# -----------------------------------------------------
# OPTION 1: (OVER)WRITE TO EXCEL FILE
# -----------------------------------------------------
with pd.ExcelWriter("output.xlsx") as writer:
    df1.to_excel(writer, sheet_name="Sheet_1")
    df2.to_excel(writer, sheet_name="Sheet_2")
    df3.to_excel(writer, sheet_name="Sheet_3")


# -----------------------------------------------------
# OPTION 2: ADD DATAFRAMES TO EXISTING WORKBOOK
# -----------------------------------------------------
with pd.ExcelWriter("output.xlsx", mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
    df1.to_excel(writer, sheet_name="Sheet_4")
    df2.to_excel(writer, sheet_name="Sheet_5")
    df3.to_excel(writer, sheet_name="Sheet_6")


# -----------------------------------------------------
# OPTION 3: ADD DATAFRAMES TO EXISTING WORKBOOK [ALTERNATIVE]
# -----------------------------------------------------
import xlwings as xw  # pip install xlwings

# Use dict for your sheet/df mapping
sheet_df_mapping = {"Sheet_7": df1, "Sheet_8": df2, "Sheet_9": df3}

# Open Excel in background
with xw.App(visible=False) as app:
    wb = app.books.open("output.xlsx")

    # List of current worksheet names
    current_sheets = [sheet.name for sheet in wb.sheets]

    # Iterate over sheet/df mapping
    # If sheet already exist, overwrite current cotent. Else, add new sheet
    for sheet_name in sheet_df_mapping.keys():
        if sheet_name in current_sheets:
            wb.sheets(sheet_name).range("A1").value = sheet_df_mapping.get(sheet_name)
        else:
            new_sheet = wb.sheets.add(after=wb.sheets.count)
            new_sheet.range("A1").value = sheet_df_mapping.get(sheet_name)
            new_sheet.name = sheet_name
    wb.save()
