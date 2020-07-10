Attribute VB_Name = "Module1"
Sub macroboletin()
Attribute macroboletin.VB_ProcData.VB_Invoke_Func = " \n14"
'
' macroboletin Macro
'

'
    Dim counter As String
    Dim hoja As String
    Dim i As Integer
    Dim hoja_principal As String
    Dim hoja_trabajo_T As String
    Dim hoja_trabajo_F As String
    Dim outStr As String
    Dim inicial As String
    Dim cont As Integer
    Dim j As Integer
    
    
    
    hoja_principal = InputBox("Introduce el nombre de la hoja donde están los nombres") 'Ponerle un input para que se ponga el nombre de la hoja
    cont = InputBox("Que numero de hoja es donde empiezan las tablas?")
    hoja_trabajo_T = "T" 'Aqui tambien vamos a concatenar la i
    hoja_trabajo_F = "F" 'Aqui vamos a concatenar la i
    
    'Agregar un IF o select case, dependiendo como se me vaya ocurriendo
    'Agregar un pequeño menú
    
    Sheets(hoja_principal).Select
    counter = Cells(1, 1).Text
    i = 1

    outStr = Left(Mid(counter, 1), 1)
    
    While counter <> ""
        Sheets(hoja_principal).Select
        counter = Cells(i, 1).Text
        hoja_trabajo = Sheets(i + cont - 1).Name
        On Error GoTo labell
        'MsgBox (counter)
        Worksheets(hoja_principal).Cells(i, 1).Copy Worksheets(hoja_trabajo).Range("A1")
        'IF tabla o figura, entonces hoja_trabajo sera T o F
        'incial = Left(Mid(counter, 1), 1) '
       ' If inicial = "T" Then
       '     hoja_trabajo = hoja_trabajo_T + CStr(i)
       ' ElseIf inicial = "F" Then
       '     hoja_trabajo =
        'MsgBox (counter)
        'Cells(1, 1) = counter
        'Sheets(hoja_trabajo).Select
        'Sheets(hoja_principal).Select
        i = i + 1
    Wend

    
    'Dim fin As String
    'MsgBox (hoja_principal)
    'MsgBox (hoja_trabajo)
    Sheets(hoja_principal).Select
    fin = Cells(j, i)
    Sheets(hoja_trabajo).Select
    Cells(1, 1) = fin
    'Range("A1").Select
    'fin = Selection.End(xlDown).Select
    'MsgBox (fin)
    'Sheets(hoja_trabajo).Select
    'Worksheets(hoja_principal).Cell(1, 1).Select.Selection.End(xlDown).Copy Worksheets(hoja_trabajo).Cell(1, 1)
    'Cells(j, 1) = fin
    'hoja = "F1"
    'rango = "A1"
    'Sheets("Hoja26").Select
    'Range(rango).Select
    'Selection.Copy
    'Sheets(hoja).Select
    'ActiveSheet.Paste
    'Range("A1").Select
    'ActiveSheet.Paste

labell:
    j = i
End Sub

