Option Explicit

Sub DesempaquetarJSON()
    Dim ws As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long, i As Long
    Dim jsonString As String
    Dim dict As Object
    Dim key As Variant
    Dim col As Long
    
    ' Establecer la hoja de origen y crear una nueva hoja para los resultados
    Set ws = ThisWorkbook.Worksheets("Base")
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Worksheets("Resultados_Desempaquetados")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Worksheets.Add
        wsOutput.Name = "Resultados_Desempaquetados"
    Else
        wsOutput.Cells.Clear
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Inicializar el diccionario para almacenar todas las claves únicas
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Primera pasada: recopilar todas las claves únicas
    For i = 1 To lastRow
        jsonString = ws.Cells(i, 1).Value
        Call ParseJSON(jsonString, dict)
    Next i
    
    ' Escribir encabezados
    col = 1
    For Each key In dict.Keys
        wsOutput.Cells(1, col).Value = key
        col = col + 1
    Next key
    
    ' Segunda pasada: escribir valores
    For i = 1 To lastRow
        jsonString = ws.Cells(i, 1).Value
        Call WriteJSONValues(jsonString, dict, wsOutput, i + 1)
    Next i
    
    wsOutput.Columns.AutoFit
    MsgBox "Proceso completado. Revisa la hoja 'Resultados_Desempaquetados'."
End Sub

Sub ParseJSON(jsonString As String, dict As Object)
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = """(\w+)""\s*:\s*""?([^""}{\[\],]+)""?"
    
    Set matches = regex.Execute(jsonString)
    
    For Each match In matches
        If Not dict.Exists(match.SubMatches(0)) Then
            dict.Add match.SubMatches(0), ""
        End If
    Next match
End Sub

Sub WriteJSONValues(jsonString As String, dict As Object, ws As Worksheet, row As Long)
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim col As Long
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = """(\w+)""\s*:\s*""?([^""}{\[\],]+)""?"
    
    Set matches = regex.Execute(jsonString)
    
    For Each match In matches
        col = Application.Match(match.SubMatches(0), ws.Rows(1), 0)
        If Not IsError(col) Then
            ws.Cells(row, col).Value = match.SubMatches(1)
        End If
    Next match
End Sub
