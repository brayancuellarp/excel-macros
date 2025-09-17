# excel-macros
Colección de macros de Excel en VBA


Sub ActualizarDatos()

    Dim wsLocal As Worksheet, wsTemp As Worksheet
    Dim lastRowLocal As Long, lastRowTemp As Long
    Dim colSFC_Local As Long, colSFC_Temp As Long
    Dim i As Long, j As Long
    Dim SFC_Buscar As String
    Dim rngFind As Range
    Dim colTemp As Variant
    
    ' Asignar hojas
    Set wsLocal = ThisWorkbook.Sheets("HojaLocal")        ' Cambia a tu hoja principal
    Set wsTemp = ThisWorkbook.Sheets("HojaTemporal")      ' Cambia a tu hoja temporal

    ' Última fila en ambas hojas
    lastRowLocal = wsLocal.Cells(wsLocal.Rows.Count, 1).End(xlUp).Row
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row

    ' Buscar la columna de SFC/PU en ambas hojas
    colSFC_Local = Application.Match("SFC/PU", wsLocal.Rows(1), 0)
    colSFC_Temp = Application.Match("SFC/PU", wsTemp.Rows(1), 0)

    If IsError(colSFC_Local) Or IsError(colSFC_Temp) Then
        MsgBox "No se encontró la columna 'SFC/PU' en alguna de las hojas.", vbCritical
        Exit Sub
    End If

    ' Recorrer cada fila en la hoja local
    For i = 2 To lastRowLocal
        SFC_Buscar = wsLocal.Cells(i, colSFC_Local).Value

        ' Buscar el mismo SFC/PU en la hoja temporal
        Set rngFind = wsTemp.Range(wsTemp.Cells(2, colSFC_Temp), wsTemp.Cells(lastRowTemp, colSFC_Temp)) _
                        .Find(What:=SFC_Buscar, LookAt:=xlWhole)

        If Not rngFind Is Nothing Then
            ' Recorrer columnas de la hoja local
            For j = 1 To wsLocal.Cells(1, wsLocal.Columns.Count).End(xlToLeft).Column
                ' Buscar si la cabecera existe también en la hoja temporal
                On Error Resume Next
                colTemp = Application.Match(wsLocal.Cells(1, j).Value, wsTemp.Rows(1), 0)
                On Error GoTo 0

                ' Si existe, copiar el valor correspondiente
                If Not IsError(colTemp) And Not IsEmpty(colTemp) Then
                    wsLocal.Cells(i, j).Value = wsTemp.Cells(rngFind.Row, colTemp).Value
                End If
            Next j
        End If
    Next i

    MsgBox "Datos actualizados correctamente.", vbInformation

End Sub
