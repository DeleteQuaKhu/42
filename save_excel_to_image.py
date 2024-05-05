import win32com.client as win32
import os

def save_excel_charts_as_images(excel_file_path, output_folder):
    xlApp = win32.gencache.EnsureDispatch('Excel.Application')

    # Open the workbook with the correct path
    workbook = xlApp.Workbooks.Open(excel_file_path)
    xlApp.Sheets("Sheet1").Select()
    xlApp.Visible = True

    xlSheet1 = workbook.Sheets(1)

    # Ensure to save any work before running script
    xlApp.DisplayAlerts = False

    i = 0
    for chart in xlSheet1.ChartObjects():
        chart.CopyPicture()
        # Create new temporary sheet
        temp_sheet = xlApp.ActiveWorkbook.Sheets.Add(After=xlApp.ActiveWorkbook.Sheets(1))
        temp_sheet.Name = "temp_sheet" + str(i)

        # Add chart object to new sheet
        cht = temp_sheet.ChartObjects().Add(0, 0, 300, 200)
        # Paste copied chart into new object
        cht.Chart.Paste()
        # Save image to specified directory
        output_path = os.path.join(output_folder, f"chart{i}.png")
        cht.Chart.Export(output_path)
        i += 1

    workbook.Close()
    # Restore default behavior
    xlApp.DisplayAlerts = True

# Specify the paths
excel_file_path = r"C:\Users\TechnoStar\Documents\macro\save png\1234 - Copy.xlsx"
output_folder_path = r"C:\Users\TechnoStar\Documents\macro\save png"

# Call the function to save charts as images
save_excel_charts_as_images(excel_file_path, output_folder_path)
