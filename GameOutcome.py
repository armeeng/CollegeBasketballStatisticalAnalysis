import openpyxl

workbook = openpyxl.load_workbook('outputexcel.xlsx')
sheet = workbook.worksheets[0]

for row in sheet.iter_rows():
    values = [cell.value for cell in row]
    Team1Score = values[7]
    Team2Score = values[8]
    if Team1Score > Team2Score:
        GameOutcomeNorm = 1
    elif Team1Score < Team2Score:
        GameOutcomeNorm = 2
    else:
        GameOutcomeNorm = 0
    print(GameOutcomeNorm)

    row[4].value = GameOutcomeNorm


workbook.save('outputexcel.xlsx')

