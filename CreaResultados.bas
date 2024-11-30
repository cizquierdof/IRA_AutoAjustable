Attribute VB_Name = "CreaResultados"

' Genera la hoja de resultados
Function HojaResultados()
    Dim ws As Worksheet
    Dim wsOrigen As Worksheet
    Dim numFilas As Integer
    Dim posicionFila As Integer
    Dim rango As Range
    Dim cellIndex As Variant
    ' número de filas de la muestra
    Set wsOrigen = ThisWorkbook.Sheets("Muestra")
    numFilas = wsOrigen.ListObjects("muestra").ListRows.Count
    
    ' Crea la hoja con borrado de versión anterior
    BorrarHoja "Resultados"
    Set ws = CrearNuevaHoja("Resultados")
    Set HojaResultados = ws ' valor de retorno
    
' ZONA DE ENCABEZADOS GENERALES
    ' Pone el título general
     With ws.Range("B2")
        .Value = "Informe de revisión de la accesibilidad"
        .Font.Size = 24
        .Font.Name = "Arial"
        .Font.Color = RGB(52, 101, 195)
    End With
    With ws.Range("B3")
        .Value = "Análisis de accesibilidad en profundidad de un sitio web"
        .Font.Size = 20
        .Font.Name = "Arial"
        .Font.Color = RGB(52, 101, 180)
    End With
    
' ZONA DE NIVEL A
    'encabezado nivel A, la fila de inicio es fija no necesita cálculo
    'cellIndex = Array(6, 3, 6, 3) 'celda C6
    EncabezadoColores ws, 6, "nivel A"
    
    ' encabezados columnas nivel A, la fila de inicio es fija no necesita cálculo
    Dim criterios As Variant
    criterios = Array("1.1.1", "1.2.1", "1.2.2", "1.2.3", "1.3.1", "1.3.2", _
            "1.3.3", "1.4.1", "1.4.2", "2.1.1", "2.1.2", "2.1.4", "2.2.1", "2.2.2", "2.3.1", _
            "2.4.1", "2.4.2", "2.4.3", "2.4.4", "2.5.1", "2.5.2", "2.5.3", "2.5.4", "3.1.1", _
            "3.2.1", "3.2.2", "3.3.1", "3.3.2", "4.1.1", "4.1.2")
    cellIndex = Array(7, 4, 7, 33) 'celdas D7:AG7
    EncabezadosCriterios ws, cellIndex, criterios
     
    ' Mostrar cuadrícula de resultados desde fila 8, col2 a fila numFilas+8, col 33
    cellIndex = Array(8, 2, numFilas + 8 - 1, 33)
    Set rango = HallarRango(ws, cellIndex)
    PonerBordes rango
    
    'convierte en tabla la zona de recepción de resultados
    'Dim nombreTabla As String
    'Dim tabla As ListObject
    cellIndex = Array(7, 4, numFilas + 8 - 1, 33)
    Set rango = HallarRango(ws, cellIndex)
    FormatoCondicional rango
    'Set tabla = ws.ListObjects.Add(xlSrcRange, rango, , xlYes)
    'tabla.ShowAutoFilter = False
    'tabla.TableStyle = ""
    'tabla.Name = "resultadosA"
    
    'copiar nombre y url de la muestra
    cellIndex = Array(8, 2, numFilas + 7, 3)
    Set rango = HallarRango(ws, cellIndex)
    Dim rangoOrigen As Range
    Set rangoOrigen = wsOrigen.Range(wsOrigen.Cells(2, 2), wsOrigen.Cells(numFilas + 1, 3))
    rangoOrigen.Copy Destination:=rango
       
    posicionFila = 8 + numFilas ' posición de siguiente fila a llenar
    
    ' Fila de pasa, falla, no aplica y Resultado A
    FilasResultados ws, posicionFila, 31
    
    ' Actualiza la posición de la fila
    posicionFila = 8 + numFilas + 6
        
 ' ZONA DE nivel  AA
    ' encabezado nivel AA empieza en posicionFila
    EncabezadoColores ws, posicionFila, "nivel AA"

    ' encabezados columnas nivel AA
    criterios = Array("1.2.4", "1.2.5", "1.3.4", "1.3.5", "1.4.3", "1.4.4", "1.4.5", _
            "1.4.10", "1.4.11", "1.4.12", "1.4.13", "2.4.5", "2.4.6", "2.4.7", "3.1.2", _
            "3.2.3", "3.2.4", "3.3.3", "3.3.4", "4.1.3")
    cellIndex = Array(posicionFila + 1, 4, posicionFila + 1, 23)
    EncabezadosCriterios ws, cellIndex, criterios
    
    ' Mostrar cuadrícula de resultados desde fila 8, col2 a fila numFilas+8, col 33
    cellIndex = Array(posicionFila + 2, 2, posicionFila + numFilas + 1, 23)
    Set rango = HallarRango(ws, cellIndex)
    PonerBordes rango
    
    'convierte en tabla la zona de recepción de resultados
    cellIndex = Array(posicionFila + 1, 4, posicionFila + numFilas + 1, 23)
    Set rango = HallarRango(ws, cellIndex)
    FormatoCondicional rango
    'Set tabla = ws.ListObjects.Add(xlSrcRange, rango, , xlYes)
    'tabla.ShowAutoFilter = False
    'tabla.TableStyle = ""
    'tabla.Name = "resultadosAA"
    
    'copiar nombre y url de la muestra
    cellIndex = Array(posicionFila + 2, 2, posicionFila + numFilas, 3)
    Set rango = HallarRango(ws, cellIndex)
    rangoOrigen.Copy Destination:=rango
    
    posicionFila = posicionFila + numFilas + 2 ' posición de siguiente fila a llenar

    ' Fila de pasa, falla, no aplica y Resultado A
    FilasResultados ws, posicionFila, 21

End Function
'Llena las filas de resultados pasa, falla, no aplica y Resultado
Sub FilasResultados(ws As Worksheet, fila As Integer, columnas)
    Dim rango As Range
    Dim celdas As Variant
    
    celdas = Array(fila + 3, 3, fila + 3, 3)
    Set rango = HallarRango(ws, celdas)
    With rango
        .Value = "Resultados"
        .RowHeight = 80
        With .Font
            .Bold = True
            .Name = "Verdana"
            .Size = 24
        End With
        .Interior.Color = RGB(208, 206, 206)
        .VerticalAlignment = xlCenter
    End With
    ' encabezados filas pasa,falla, no aplica
    celdas = Array(fila, 3, fila, 3)
    Set rango = HallarRango(ws, celdas)
    With rango
        .Value = "Pasa"
        With .Font
            .Bold = True
            .Name = "Verdana"
            .Size = 12
        End With
    End With
    celdas = Array(fila + 1, 3, fila + 1, 3)
    Set rango = HallarRango(ws, celdas)
    With rango
        .Value = "Falla"
        With .Font
            .Bold = True
            .Name = "Verdana"
            .Size = 12
        End With
    End With
    celdas = Array(fila + 2, 3, fila + 2, 3)
    Set rango = HallarRango(ws, celdas)
    With rango
        .Value = "No aplica"
        With .Font
            .Bold = True
            .Name = "Verdana"
            .Size = 12
            .Color = RGB(255, 255, 255)
        End With
    End With
    celdas = Array(fila + 2, 3, fila + 2, columnas + 2)
    Set rango = HallarRango(ws, celdas)
    rango.Interior.Color = RGB(117, 113, 113)
    celdas = Array(fila, 3, fila + 3, columnas + 2)
    Set rango = HallarRango(ws, celdas)
    PonerBordes rango
End Sub
'Llena los encabezados de colores
Sub EncabezadoColores(ws As Worksheet, fila As Integer, nivel As String)
    Dim celdasRango As Variant
    Dim rango As Range
         
    celdasRango = Array(fila, 3, fila, 3) 'rango C3
    Set rango = HallarRango(ws, celdasRango)
With rango
        .Font.Name = "Verdana"
        .Font.Size = 12
        .Font.Bold = True
        .Value = nivel
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(201, 218, 248)
        .RowHeight = 15
        .ColumnWidth = 45
    End With
    ' colorines de cabecera nivel A se ponen en la misma fila que celdas(0)
    Dim numCriterios As Variant
    'establece el ancho de las zonas de color
    
    If nivel = "nivel A" Then
        colZona = Array(12, 26, 31, 33)
    Else
        colZona = Array(14, 17, 22, 23)
    End If
    
    celdasRango = Array(fila, 4, fila, colZona(0))
    Set rango = HallarRango(ws, celdasRango)
    rango.Interior.Color = RGB(212, 234, 107) ' zona de perceptible
    celdasRango = Array(fila, colZona(0) + 1, fila, colZona(1))
    Set rango = HallarRango(ws, celdasRango)
    rango.Interior.Color = RGB(255, 191, 0) ' zona de operable
    celdasRango = Array(fila, colZona(1) + 1, fila, colZona(2))
    Set rango = HallarRango(ws, celdasRango)
    rango.Interior.Color = RGB(255, 109, 109) ' zona de comprensible
    celdasRango = Array(fila, colZona(2) + 1, fila, colZona(3))
    Set rango = HallarRango(ws, celdasRango)
    rango.Interior.Color = RGB(89, 131, 176) ' zona de robusto

End Sub

Sub EncabezadosCriterios(ws As Worksheet, celdas As Variant, criterios As Variant)
    Dim rango As Range
    Dim celdaMuestra As Variant
    'rango para el encabezado Muestra es la columna anterior a celdas
    celdaMuestra = Array(celdas(0), celdas(1) - 1, celdas(0), celdas(1) - 1)
    Set rango = HallarRango(ws, celdaMuestra)
    
    With rango
        .Orientation = 0
        .Value = "Muestra"
        .Font.Name = "Verdana"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        PonerBordes rango
    End With
    Set rango = HallarRango(ws, celdas)
    With rango
        .Value = criterios
        .Orientation = 90
        .Font.Name = "Verdana"
        .Font.Size = 11.5
        .Font.Bold = True
        .RowHeight = 50
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 6
        PonerBordes rango
    End With
End Sub
