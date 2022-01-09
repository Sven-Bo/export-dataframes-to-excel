import pandas as pd

# Create three dataframes
tips = pd.read_csv(
    "https://raw.githubusercontent.com/mwaskom/seaborn-data/master/tips.csv"
)
tips_by_gender = tips.groupby(by="sex").sum()[["tip"]]
tips_by_day = tips.groupby(by="day").sum()[["tip"]]

# Use XlsxWriter as the engine
writer = pd.ExcelWriter("tips.xlsx", engine="xlsxwriter")

# Write each dataframe to a different worksheet.
tips.to_excel(writer, sheet_name="Sheet1")
tips_by_gender.to_excel(writer, sheet_name="Sheet2")
tips_by_day.to_excel(writer, sheet_name="Sheet3")

# Close Excel writer and output the Excel file.
writer.save()
