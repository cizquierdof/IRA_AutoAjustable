Attribute VB_Name = "Miscelaneo"

' añade bordes a un rango de celdas
Sub PonerBordes(rango As Range)
            With rango.Borders
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
End Sub

' Hallar rango de celdas de
Function HallarRango(ws As Worksheet, cellIndex As Variant) As Range
    Set HallarRango = ws.Range(ws.Cells(cellIndex(0), cellIndex(1)), ws.Cells(cellIndex(2), cellIndex(3)))
End Function

'Devuelve el num filas de la muestra
Function FilasMuestra() As Integer
 Dim muestra As Worksheet
 Set muestra = ThisWorkbook.Sheets("Muestra")
 FilasMuestra = muestra.ListObjects("muestra").ListRows.Count
End Function

Sub BorrarVentanaInmediato()
    With Application.VBE.Windows("Inmediato")
        .SetFocus
        SendKeys "^a" ' Seleccionar todo (Ctrl + A)
        SendKeys "{DEL}" ' Eliminar todo el contenido
    End With
End Sub

'Añade el formato condicional a un rango
Sub FormatoCondicional(rango As Range)
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

End Sub





