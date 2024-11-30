Attribute VB_Name = "Inicio"
Sub inicio()
Attribute inicio.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim principio As String
    Dim cabeceras(1 To 2) As String
    Dim ws As Worksheet
    Dim sc As Object
    Set sc = New Scripting.Dictionary

    ' Asignar la tecla Ctrl + Shift + M para ejecutar la macro "TuMacro"
    Application.OnKey "^+B", "BorrarVentanaInmediato"


    Application.AutoCorrect.AutoExpandListRange = False
    'Montar Robusto
    principio = "Robusto"
    cabeceras(1) = "Principio 4: Robusto"
    cabeceras(2) = "El contenido debe ser lo suficientemente robusto para que pueda ser interpretado de manera confiable"
    'a�adir criterios
    sc.RemoveAll
    sc.Add "4.1.1 Procesamiento", "A"
    sc.Add "4.1.2 Nombre, funci�n, valor", "A"
    sc.Add "4.1.3 Mensajes de estado", "AA"
    CreaPrincipio principio, sc, cabeceras ' monta toda la hoja
    
    'Montar Comprensible
    principio = "Comprensible"
    cabeceras(1) = "Principio 3: Comprensible"
    cabeceras(2) = "La informaci�n y el manejo de la interfaz de usuario deben ser comprensibles."
    'a�adir criterios
    sc.RemoveAll
    sc.Add "3.1.1 Idioma de la p�gina", "A"
    sc.Add "3.1.2 Idioma de las partes", "AA"
    sc.Add "3.2.1 Al recibir el foco", "A"
    sc.Add "3.2.2 Al recibir entradas", "A"
    sc.Add "3.2.3 Navegaci�n coherente", "AA"
    sc.Add "3.2.4 Identificaci�n coherente", "AA"
    sc.Add "3.3.1 Identificaci�n de errores", "A"
    sc.Add "3.3.2 Etiquetas o instrucciones", "A"
    sc.Add "3.3.3 Sugerencias ante errores", "AA"
    sc.Add "3.3.4 Prevenci�n de errores (legales, financieros, datos)", "AA"
    CreaPrincipio principio, sc, cabeceras ' monta toda la hoja
    
    'Montar Operable
    principio = "Operable"
    cabeceras(1) = "Principio 2: Operable"
    cabeceras(2) = "Los componentes de la interfaz de usuario y la navegaci�n deben ser operables."
    'a�adir criterios
    sc.RemoveAll
    sc.Add "2.1.1 Teclado", "A"
    sc.Add "2.1.2 Sin trampas para el foco del teclado", "A"
    sc.Add "2.1.4  Atajos con teclas de caracteres", "A"
    sc.Add "2.2.1 Tiempo ajustable", "A"
    sc.Add "2.2.2 Poner en pausa, detener, ocultar", "A"
    sc.Add "2.3.1 Umbral de tres destellos o menos", "A"
    sc.Add "2.4.1 Evitar bloques", "A"
    sc.Add "2.4.2 Titulado de p�ginas", "A"
    sc.Add "2.4.3 Orden del foco", "A"
    sc.Add "2.4.4 Prop�sito de los enlaces (en contexto)", "A"
    sc.Add "2.4.5 M�ltiples v�as", "AA"
    sc.Add "2.4.6 Encabezados y etiquetas", "AA"
    sc.Add "2.4.7 Foco visible", "AA"
    sc.Add "2.5.1 Gestos del puntero", "A"
    sc.Add "2.5.2 Cancelaci�n del puntero", "A"
    sc.Add "2.5.3 Etiqueta en el nombre", "A"
    sc.Add "2.5.4. Activaci�n mediante movimiento", "A"
    CreaPrincipio principio, sc, cabeceras ' monta toda la hoja
    
    'Montar perceptible
    principio = "Perceptible"
    cabeceras(1) = "Principio 1: Perceptible"
    cabeceras(2) = "La informaci�n y los componentes de la interfaz de usuario deben ser presentados a los usuarios de modo que ellos puedan percibirlos."
    'a�adir criterios
        sc.RemoveAll
    sc.Add "1.1.1 Contenido no textual", "A"
    sc.Add "1.2.1 S�lo audio y s�lo v�deo (grabado)", "A"
    sc.Add "1.2.2 Subt�tulos (grabados)", "A"
    sc.Add "1.2.3 Audiodescripci�n o Medio Alternativo (grabado)", "A"
    sc.Add "1.2.4 Subt�tulos (en directo)", "AA"
    sc.Add "1.2.5 Audiodescripci�n (grabado)", "AA"
    sc.Add "1.3.1 Informaci�n y relaciones", "A"
    sc.Add "1.3.2 Secuencia significativa", "A"
    sc.Add "1.3.3 Caracter�sticas sensoriales", "A"
    sc.Add "1.3.4 Orientaci�n", "AA"
    sc.Add "1.3.5 Identificar el prop�sito de campo", "AA"
    sc.Add "1.4.1 Uso del Color", "A"
    sc.Add "1.4.2 Control del audio", "A"
    sc.Add "1.4.3 Contraste (m�nimo)", "AA"
    sc.Add "1.4.4 Redimensi�n del texto", "AA"
    sc.Add "1.4.5 Im�genes de texto", "AA"
    sc.Add "1.4.10 Reflow", "AA"
    sc.Add "1.4.11 Contraste en elementos no textuales", "AA"
    sc.Add "1.4.12 Espaciado en el texto", "AA"
    sc.Add "1.4.13 Contenido en over o focus", "AA"
    CreaPrincipio principio, sc, cabeceras ' monta toda la hoja
    
    ' Montar resultados
    Set ws = HojaResultados()
    ws.Move after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count) 'Pone resultados al final
    
    'llena resultados con las f�rmulas de importaci�n de resultados
    Dim principios() As Variant
    Dim i As Integer
    principios = Array("Perceptible", "Operable", "Comprensible", "Robusto")
    For i = LBound(principios) To UBound(principios)
        principio = principios(i)
        FormulasResultados principio
    Next i
    
    ThisWorkbook.Sheets("Muestra").Move Before:=ThisWorkbook.Sheets(1)
End Sub
