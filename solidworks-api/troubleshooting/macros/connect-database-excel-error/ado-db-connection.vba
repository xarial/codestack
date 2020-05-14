Set conn = CreateObject("ADODB.Connection")
Set records = CreateObject("ADODB.Recordset")
    
Dim xlsFilePath As String
xlsFilePath = swApp.GetCurrentMacroPathFolder() & "\" & EXCEL_FILE_NAME
    
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & xlsFilePath & _
            ";Extended Properties=""Excel 8.0;HDR=Yes;"";"