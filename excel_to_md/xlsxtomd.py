import openpyxl

# Open xlsx file
wb = openpyxl.load_workbook('2 Corinthian-L3.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

# Open txt file for writing
outfile = open("2 Corinthian-L3.md", "w")

# Access each element of a column row by row and write it to txt file
for row in range(2, sheet.max_row + 1):
    name = sheet['D' + str(row)].value
    print name
    outfile.write(unicode(name).encode('utf-8').strip())
    outfile.write("\n")

# Close the file
outfile.close()
print 'Done.'
