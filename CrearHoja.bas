Attribute VB_Name = "CrearHoja"
Sub LlenaHoja(sc As Scripting.Dictionary, wsDestino As Worksheet)

    Dim wsOrigen As Worksheet
    Dim wsDatos As Worksheet
    Dim tablaOrigen As ListObject
    Dim tablaEntradasValidas As ListObject
    Dim ultimaFila As Integer
    Dim i As Long
    Dim encabezados As Variant
    Dim posicionFila As Integer
    Dim rango As Range
    
    
    ' Establece las hojas de trabajo
    Set wsOrigen = ThisWorkbook.Sheets("Muestra") ' Hoja de origen
    Set wsDatos = ThisWorkbook.Sheets("Datos") ' Hoja de donde se obtiene la lista de valores

    ' Establece la tabla de origen y la tabla de entradas válidas
    Set tablaOrigen = wsOrigen.ListObjects("muestra") ' Nombre de la tabla de origen
    Set tablaEntradasValidas = wsDatos.ListObjects("EntradasValidas") ' Nombre de la tabla con los valores válidos
    
    'llena la cabecera
    SetEncabezados wsDestino ' cabecera común

    ' Establece la fila inicial
    posicionFila = 18
    
    wsDestino.Columns("C").ColumnWidth = 56
    wsDestino.Columns("B").ColumnWidth = 15.55

' Itera a través de los encabezados y copia los datos de la tabla para cada sección
    For Each clave In sc
    
      ' Copia los encabezados de la tabla original y da formato
      Set rango = wsDestino.Range("A" & posicionFila & ":D" & posicionFila)
        With rango
            .Value = Array("id", sc(clave), clave, "Resultado")
            With .Font
                .Bold = True
                .Size = 12
                .Name = "Calibri"
            End With
            .Interior.Color = RGB(201, 218, 248) ' violeta claro
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 32
            PonerBordes rango
        End With
        ' formato específico para el indicador de nivel WCAG
        With wsDestino.Range("B" & posicionFila & ":B" & posicionFila)
            .Font.Size = 24
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(89, 131, 176)
        End With
        ' cambio de fuente para el t´titulo del SC
        wsDestino.Range("C" & posicionFila & ":C" & posicionFila).Font.Name = "Arial"

' Copia los datos de la tabla original a la nueva hoja
        For i = 1 To tablaOrigen.ListRows.Count
            With tablaOrigen.ListRows(i).Range
                wsDestino.Cells(posicionFila + i, 1).Value = .Cells(1, 1).Value
                wsDestino.Cells(posicionFila + i, 2).Value = .Cells(1, 2).Value
                wsDestino.Cells(posicionFila + i, 3).Value = .Cells(1, 3).Value
                ' Añadir bordes a las celdas de destino
                Set rango = wsDestino.Range(wsDestino.Cells(posicionFila + i, 1), wsDestino.Cells(posicionFila + i, 4))
                PonerBordes rango
            End With
        Next i

        ' Centra el contenido de la primera columna
        Set rango = wsDestino.Columns("A")
        With rango
            .HorizontalAlignment = xlCenter
            .ColumnWidth = 6
        End With
       
        ' Define la validación de datos para la columna "Resultado"
        With wsDestino
            .Range("D" & (posicionFila + 1) & ":D" & (posicionFila + tablaOrigen.ListRows.Count)).Validation.Delete ' Elimina cualquier validación existente
            .Range("D" & (posicionFila + 1) & ":D" & (posicionFila + tablaOrigen.ListRows.Count)).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="=" & wsDatos.Name & "!" & tablaEntradasValidas.ListColumns(1).DataBodyRange.Address
        End With

        ' Establece el formato condicional para la columna D
        Set rango = wsDestino.Columns("D")
        rango.FormatConditions.Delete 'Borrar formatos condicionales anteriores
        ' añadir formatos condicionales
            ' Falla
        With rango.FormatConditions.Add(Type:=xlTextString, String:="falla", TextOperator:=xlContains)
            .Font.Color = RGB(255, 0, 0) ' Rojo
            .Font.Bold = True
            .Interior.Color = RGB(253, 234, 236)
        End With
            ' Pasa
        With rango.FormatConditions.Add(Type:=xlTextString, String:="pasa", TextOperator:=xlContains)
            .Font.Color = RGB(60, 125, 34) ' Verde
            .Font.Bold = True
            .Interior.Color = RGB(237, 249, 244)
        End With
             ' N/A
        With rango.FormatConditions.Add(Type:=xlTextString, String:="n/a", TextOperator:=xlContains)
            .Interior.Color = RGB(181, 230, 162) ' Verde claro
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0) ' Negro
        End With
        
        'convierte en tabla
        Dim nombreTabla As String
        Dim tabla As ListObject
        Dim filaFinal As Integer
        filaFinal = posicionFila + tablaOrigen.ListRows.Count
        Set rango = wsDestino.Range(wsDestino.Cells(posicionFila, 1), wsDestino.Cells(filaFinal, 4))
        Set tabla = wsDestino.ListObjects.Add(xlSrcRange, rango, , xlYes)
        Dim posEspacio As Integer
        posEspacio = InStr(1, clave, " ")
        nombreTabla = Left(clave, posEspacio - 1)
        tabla.ShowAutoFilter = False
        tabla.TableStyle = ""
        tabla.Name = "T" & Replace(nombreTabla, ".", "_")
        
        ' Deja dos filas entre bloques
        posicionFila = posicionFila + tablaOrigen.ListRows.Count + 3
        
    Next clave
End Sub


