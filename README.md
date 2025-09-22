# excel-macros
Colección de macros de Excel en VBA


Private Sub btn_1_Click()
    Call ActualizarDatos
End Sub

' Agregando la macro con el nombre que busca el botón
Sub Prueba_1()
    Call ActualizarDatos
End Sub

Sub ActualizarDatos()
    ' Agregando variables para manejo de errores y optimización
    Dim wsLocal As Worksheet, wsTemp As Worksheet
    Dim lastRowLocal As Long, lastRowTemp As Long
    Dim colSFC_Local As Long, colSFC_Temp As Long
    Dim i As Long, j As Long
    Dim SFC_Buscar As String
    Dim rngFind As Range
    Dim colTemp As Variant
    ' Agregando contadores para actualizaciones e inserciones
    Dim registrosActualizados As Long, registrosInsertados As Long
    Dim nuevaFila As Long
    
    ' Desactivar actualización de pantalla para mejor rendimiento
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Verificar que las hojas existan
    On Error Resume Next
    Set wsLocal = ThisWorkbook.Sheets("Consolidado")
    Set wsTemp = ThisWorkbook.Sheets("DatosTemporales")
    On Error GoTo ErrorHandler
    
    If wsLocal Is Nothing Then
        MsgBox "No se encontró la hoja 'Consolidado'", vbCritical
        GoTo Cleanup
    End If
    
    If wsTemp Is Nothing Then
        MsgBox "No se encontró la hoja 'DatosTemporales'", vbCritical
        GoTo Cleanup
    End If

    ' Verificar que ambas hojas tengan datos
    lastRowLocal = wsLocal.Cells(wsLocal.Rows.Count, 1).End(xlUp).Row
    lastRowTemp = wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).Row
    
    If lastRowTemp < 2 Then
        MsgBox "La hoja 'DatosTemporales' no tiene datos para procesar", vbInformation
        GoTo Cleanup
    End If

    ' Buscar la columna de SFC/PU en ambas hojas
    On Error Resume Next
    colSFC_Local = Application.Match("SFC/PU", wsLocal.Rows(1), 0)
    colSFC_Temp = Application.Match("SFC/PU", wsTemp.Rows(1), 0)
    On Error GoTo ErrorHandler

    If IsError(colSFC_Local) Or IsError(colSFC_Temp) Then
        MsgBox "No se encontró la columna 'SFC/PU' en alguna de las hojas." & vbCrLf & _
               "Verifique que los encabezados estén correctos.", vbCritical
        GoTo Cleanup
    End If
    
    ' Inicializar contadores para actualizaciones e inserciones
    registrosActualizados = 0
    registrosInsertados = 0

    ' Ahora recorremos los datos temporales para actualizar o insertar
    For i = 2 To lastRowTemp
        ' Mostrar progreso en la barra de estado
        Application.StatusBar = "Procesando registro " & (i - 1) & " de " & (lastRowTemp - 1)
        
        SFC_Buscar = Trim(CStr(wsTemp.Cells(i, colSFC_Temp).Value))
        
        ' Solo procesar si SFC_Buscar no está vacío
        If SFC_Buscar <> "" Then
            ' Buscar el SFC/PU en la hoja consolidado
            Set rngFind = Nothing
            If lastRowLocal >= 2 Then
                Set rngFind = wsLocal.Range(wsLocal.Cells(2, colSFC_Local), wsLocal.Cells(lastRowLocal, colSFC_Local)) _
                                .Find(What:=SFC_Buscar, LookAt:=xlWhole, MatchCase:=False)
            End If

            If Not rngFind Is Nothing Then
                ' SFC encontrado - ACTUALIZAR registro existente
                registrosActualizados = registrosActualizados + 1
                
                ' Recorrer columnas de la hoja temporal
                For j = 1 To wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
                    ' Buscar si la cabecera existe también en la hoja consolidado
                    On Error Resume Next
                    colTemp = Application.Match(wsTemp.Cells(1, j).Value, wsLocal.Rows(1), 0)
                    On Error GoTo ErrorHandler

                    ' Si existe, copiar el valor correspondiente
                    If Not IsError(colTemp) And Not IsEmpty(colTemp) Then
                        ' Solo actualizar si el valor temporal no está vacío
                        If Not IsEmpty(wsTemp.Cells(i, j).Value) Then
                            wsLocal.Cells(rngFind.Row, colTemp).Value = wsTemp.Cells(i, j).Value
                        End If
                    End If
                Next j
            Else
                ' SFC NO encontrado - INSERTAR nuevo registro
                registrosInsertados = registrosInsertados + 1
                nuevaFila = lastRowLocal + 1
                
                ' Recorrer columnas de la hoja temporal
                For j = 1 To wsTemp.Cells(1, wsTemp.Columns.Count).End(xlToLeft).Column
                    ' Buscar si la cabecera existe también en la hoja consolidado
                    On Error Resume Next
                    colTemp = Application.Match(wsTemp.Cells(1, j).Value, wsLocal.Rows(1), 0)
                    On Error GoTo ErrorHandler

                    ' Si existe, copiar el valor correspondiente
                    If Not IsError(colTemp) And Not IsEmpty(colTemp) Then
                        wsLocal.Cells(nuevaFila, colTemp).Value = wsTemp.Cells(i, j).Value
                    End If
                Next j
                
                ' Actualizar el último row local para la siguiente inserción
                lastRowLocal = nuevaFila
            End If
        End If
    Next i

    ' Mostrar resultado con ambos contadores
    MsgBox "Proceso completado exitosamente." & vbCrLf & _
           "Registros actualizados: " & registrosActualizados & vbCrLf & _
           "Registros insertados: " & registrosInsertados & vbCrLf & _
           "Total procesados: " & (registrosActualizados + registrosInsertados), vbInformation

    ' Limpiar datos temporales después del proceso exitoso
    Call LimpiarDatosTemporales(wsTemp)

    GoTo Cleanup

ErrorHandler:
    MsgBox "Error en la macro: " & Err.Description & vbCrLf & _
           "Línea de error: " & Erl, vbCritical

Cleanup:
    ' Restaurar configuración de Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' Limpiar objetos
    Set wsLocal = Nothing
    Set wsTemp = Nothing
    Set rngFind = Nothing

End Sub

' Nueva subrutina para limpiar datos temporales
Sub LimpiarDatosTemporales(ws As Worksheet)
    Dim lastRow As Long
    
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Solo limpiar si hay datos (más de la fila de encabezados)
    If lastRow > 1 Then
        ' Limpiar desde la fila 2 hasta la última fila con datos
        ws.Range("2:" & lastRow).ClearContents
        Application.StatusBar = "Datos temporales limpiados"
    End If
    
    On Error GoTo 0
End Sub
