import win32com.client
import time
import gc
import os

caminho_arquivo = r"C:\Users\mateus.souza\Documents\Excel\Desenhos Codificados 11-02-2026.xlsm"
caminho_xlsx = caminho_arquivo.replace(".xlsm", ".xlsx")

if os.path.exists(caminho_xlsx):
    os.remove(caminho_xlsx)

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(caminho_arquivo)

    # Atualizar dados
    wb.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()

    time.sleep(10)
    ws = wb.ActiveSheet

    ws.Columns("C:D").NumberFormat = "dd/mm/aaaa hh:mm"

    ws.Columns("C:D").HorizontalAlignment = -4108  

    wb.Save()

    wb.SaveAs(caminho_xlsx, 51)

    wb.Close(False)

finally:
    excel.Quit()
    del wb
    del excel
    gc.collect()

print("Atualização concluída.")