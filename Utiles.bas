Attribute VB_Name = "Utiles"

' Llena las características generales de las hojas de Principios
Sub CreaPrincipio(nombre As String, sc As Scripting.Dictionary, cabeceras() As String)
    Dim ws As Worksheet
    Dim codColor As Long
    
    'Selecciona el color del principio
    Select Case nombre
         Case "Perceptible"
            codColor = RGB(212, 234, 107)
         Case "Operable"
            codColor = RGB(255, 191, 0)
         Case "Comprensible"
            codColor = RGB(255, 109, 109)
         Case "Robusto"
            codColor = RGB(89, 131, 176)
    End Select
    
    ' elimina versión previa si existe
    BorrarHoja nombre
        
    ' Crea la hoja
    Set ws = CrearNuevaHoja(nombre) 'llamada a función
    
      ' cabecera específica
     With ws.Range("B6:M6")
        .RowHeight = 12
        .Interior.Color = codColor
    End With
    With ws.Range("B8:B9")
        .Value = cabeceras(1)
        With .Font
            .Name = "Arial"
            .Size = 18
            .Bold = True
        End With
    End With
    ws.Range("B9").Value = cabeceras(2)
    
    TablaCantidadCriterios ws, sc  'tabla cantidad de criterios
    
     LlenaHoja sc, ws  ' rellenar la hoja
End Sub

' crea una nueva hoja con nombre
Function CrearNuevaHoja(nombre As String) As Worksheet
    Dim nuevaHoja As Worksheet
    
    Set nuevaHoja = ThisWorkbook.Sheets.Add
    On Error Resume Next
    nuevaHoja.Name = nombre
    On Error GoTo 0
    
    Set CrearNuevaHoja = nuevaHoja
End Function

' elimina la hoja con el nombre indicado
Sub BorrarHoja(nombre As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombre)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub

'añade los encabezados comunes de las hojas
Sub SetEncabezados(ws As Worksheet)
    ' Pone cabecera general
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
    ws.Range("B4").RowHeight = 26
    ws.Range("B5").Value = "Revisión de la muestra"
    With ws.Range("B5:M5")
        .Font.Size = 18
        .Font.Name = "Helvetica Neue"
        .Interior.Color = RGB(204, 204, 204)
    End With
End Sub

' Llenar la tabla de cantidad de criterios de la cabecera de las hojas
Sub TablaCantidadCriterios(wsDestino As Worksheet, sc As Scripting.Dictionary)
    Dim rango As Range
    Dim numCriteriosNivelA As Integer
    Dim numCriteriosNivelAA As Integer
    nivelA = 0
    nivelAA = 0
    ' calcula num criterios de cada nivel
    For Each res In sc
        If sc(res) = "A" Then numCriteriosNivelA = numCriteriosNivelA + 1
        If sc(res) = "AA" Then numCriteriosNivelAA = numCriteriosNivelAA + 1
    Next res
    'Formatear la tabla
    Set rango = wsDestino.Range("B11:C14")
    PonerBordes rango
    With wsDestino.Range("B11:C11")
        .Merge
        .Interior.Color = RGB(232, 242, 161)
        .Font.Size = 10
        .Font.Name = "Arial"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Value = "Criterios de Conformidad"
    End With
    
    ' Añadir el contenido
    wsDestino.Range("b12:c12").Value = Array("A", numCriteriosNivelA)
    wsDestino.Range("b13:c13").Value = Array("AA", numCriteriosNivelAA)
    wsDestino.Range("b14:c14").Value = Array("Total", numCriteriosNivelA + numCriteriosNivelAA)
End Sub
 
 '=SI(ESBLANCO(Perceptible!D57);"";SI(Perceptible!D57="pasa";"P";SI(Perceptible!D57="falla";"F";Perceptible!D57)))
 ' wsResultados.Cells(filaInicioAA + j, i).Formula = "=Perceptible!D" & (fila + j)
 
' Rellenar las fórmulas de la hoja resultados
Sub FormulasResultados(principio As String)
    Dim wsPrincipio As Worksheet
    Dim wsResultados As Worksheet
    Dim i As Integer
    Dim filaInicioA As Integer, filaInicioAA As Integer
    Dim numFilas As Integer
    Dim fila As Integer, numTablas As Integer
    Dim columnaA As Integer, columnaAA As Integer
    Dim tabla As ListObject
    Dim nombresTablas() As String
    Dim formula As String
    
    Set wsPrincipio = ThisWorkbook.Sheets(principio)
    Set wsResultados = ThisWorkbook.Sheets("Resultados")
    
    numFilas = FilasMuestra() ' número de filas de la muestra
    
    ' define la primera columna para cada principio
    Select Case principio
     Case "Perceptible"
        columnaA = 4
        columnaAA = 4
    Case "Operable"
        columnaA = 13
        columnaAA = 15
    Case "Comprensible"
        columnaA = 27
        columnaAA = 18
    Case "Robusto"
        columnaA = 32
        columnaAA = 23
    End Select
    
    'recorre las tablas del principio y llenar las celdas de resultados
    numTablas = wsPrincipio.ListObjects.Count ' numero de tablas

    'TABLAS A
    For Each tabla In wsPrincipio.ListObjects
        If tabla.HeaderRowRange.Cells(1, 2) = "A" Then
            filaInicioA = 8 ' resultados A empieza en la fila 8
            For fila = 1 To numFilas 'para cada fila de la muestra
                formula = "=INDEX(" & tabla.Name & "[Resultado], " & fila & ")"
                wsResultados.Cells(filaInicioA + fila - 1, columnaA).FormulaR1C1 = formula
            Next fila
            columnaA = columnaA + 1
        End If
        
        'TABLAS AA
        If tabla.HeaderRowRange.Cells(1, 2) = "AA" Then
            filaInicioAA = filaInicioA + numFilas + 8
            For fila = 1 To numFilas 'para cada fila de la muestra
                formula = "=INDEX(" & tabla.Name & "[Resultado], " & fila & ")"
                wsResultados.Cells(filaInicioAA + fila - 1, columnaAA).FormulaR1C1 = formula
            Next fila
            columnaAA = columnaAA + 1
        End If
        
    Next tabla
    
 
    ' Rellenar área inferior de Resultados (Tipo AA)
    'For i = 1 To rangoTipoAA.Cells.Count
    '    wsResultados.Cells(44 + i - 1, 1).Formula = "=Perceptible!D" & rangoTipoAA.Cells(i).Row
    'Next i
End Sub

