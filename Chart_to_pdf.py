import win32com.client as win32
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os

def chartsToPdf(excel_path, sheet_name):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(excel_path)
    ws = wb.Sheets(sheet_name)

    pdf_path = os.path.join(os.path.dirname(excel_path), f"{sheet_name}.pdf")
    temp_image = os.path.join(os.path.dirname(excel_path), "temp_chart.png")

    c = canvas.Canvas(pdf_path, pagesize=letter)
    width, height = letter

    try:
        chart_count = ws.ChartObjects().Count
        if chart_count == 0:
            print(f"No charts found in sheet: {sheet_name}")
        else:
            for i in range(1, chart_count + 1):
                chart = ws.ChartObjects(i)
                chart.Chart.Export(temp_image, "PNG")  
                c.drawImage(temp_image, 50, 200, width - 100, height - 300, preserveAspectRatio=True) 
                c.showPage()

        c.save()
        print(f"Charts exported successfully to {pdf_path}")

    except Exception as e:
        print("Error:", e)

    finally:
        if os.path.exists(temp_image):
            os.remove(temp_image)
        wb.Close(False)
        excel.Quit()

chartsToPdf(r"C:\Users\DELL\Downloads\Graph.xlsx", "Graph")
