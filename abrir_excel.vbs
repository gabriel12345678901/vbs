Set objExcel_DM37 = CreateObject("Excel.Application")
Set objExcel_antenas = CreateObject("Excel.Application")
Set objExcel_decoders = CreateObject("Excel.Application")
Set objExcel_lnbf = CreateObject("Excel.Application")
Set objExcel_DM40 = CreateObject("Excel.Application")

diaOntem = Day(date)-1
mes = Month(date)
ano = Year(date)

'DM37
objExcel_DM37.Workbooks.Open("C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\DM0037_Miscelaneas_Aplicados_Click_" & ano & "_0" & mes & "_" & diaOntem & ".xlsx")
objExcel_DM37.Application.Visible = False
objExcel_DM37.ActiveWorkbook.SaveAs "C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\atualizar\DM0037_Miscelaneas_Aplicados_Click_" & ano & "_0" & mes & "_" & diaOntem & ".xlsx"
objExcel_DM37.ActiveWorkbook.Close

'ANTENAS
objExcel_antenas.Workbooks.Open("C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\DM0039_Antenas_Aplicadas_Baixa_" & ano & "_0" & mes & "_" & diaOntem & ".csv")
objExcel_antenas.Application.Visible = False
objExcel_antenas.ActiveWorkbook.SaveAs "C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\atualizar\DM0039_Antenas_Aplicadas_Baixa_" & ano & "_0" & mes & "_" & diaOntem & ".xls"
objExcel_antenas.ActiveWorkbook.Close

'DECODERS
objExcel_decoders.Workbooks.Open("C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\DM0039_Decoders_Aplicados_Baixa_" & ano & "_0" & mes & "_" & diaOntem & ".csv")
objExcel_decoders.Application.Visible = False
objExcel_decoders.ActiveWorkbook.SaveAs "C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\atualizar\DM0039_Decoders_Aplicados_Baixa_" & ano & "_0" & mes & "_" & diaOntem & ".xls"
objExcel_decoders.ActiveWorkbook.Close

'LNBF
objExcel_lnbf.Workbooks.Open("C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\DM0039_LNBF_Aplicados_Baixa_" & ano & "_0" & mes & "_" & diaOntem & ".csv")
objExcel_lnbf.Application.Visible = False
objExcel_lnbf.ActiveWorkbook.SaveAs "C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\atualizar\DM0039_LNBF_Aplicados_Baixa_" & ano & "_0" & mes & "_" & diaOntem & ".xls"
objExcel_lnbf.ActiveWorkbook.Close

'DM40
objExcel_DM40.Workbooks.Open("C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\DM0040_Materiais_Aplicados_Click_DTH_" & ano & "_0" & mes & "_" & diaOntem & ".xlsx")
objExcel_DM40.Application.Visible = False
objExcel_DM40.ActiveWorkbook.SaveAs "C:\Users\tr642633\Documents\valdirene\distribuicao_dth\relatorios utilizados\ftp\atualizar\DM0040_Materiais_Aplicados_Click_DTH_" & ano & "_0" & mes & "_" & diaOntem & ".xlsx"
objExcel_DM40.ActiveWorkbook.Close

'aqui pode dar erro:
'objExcel.Application.Quit antes era assim.
objExcel_DM37.Application.Quit
objExcel_antenas.Application.Quit
objExcel_decoders.Application.Quit
objExcel_lnbf.Application.Quit
objExcel_DM40.Application.Quit

WScript.Echo "Finished."
WScript.Quit