Attribute VB_Name = "metacpsX_Mod"
Option Explicit
Public NomAr As String
Public gcell As Range
Public hoja As Worksheet
Public wb As Workbook
Public origXML As String
Public origDIR As String
Public nomsec As String
Public niv As String
'Sub Macro1()
'
' Macro1 Macro
'

'
 '   With ActiveSheet.QueryTables.Add(Connection:= _
  '      "TEXT;C:\Users\Documents\CV_form.txt", Destination:=Range("$A$1"))
   '     .Name = "CV_form"
    '    .FieldNames = True
     '   .RowNumbers = False
      '  .FillAdjacentFormulas = False
       ' .PreserveFormatting = True
       ' .RefreshOnFileOpen = False
        '.RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
 '       .SaveData = True
  '      .AdjustColumnWidth = True
   '     .RefreshPeriod = 0
    '    .TextFilePromptOnRefresh = False
     '   .TextFilePlatform = 1252
      '  .TextFileStartRow = 1
       ' .TextFileParseType = xlDelimited
        '.TextFileTextQualifier = xlTextQualifierDoubleQuote
'        .TextFileConsecutiveDelimiter = False
 '       .TextFileTabDelimiter = False
  '      .TextFileSemicolonDelimiter = False
   '     .TextFileCommaDelimiter = False
    '    .TextFileSpaceDelimiter = False
     '   .TextFileOtherDelimiter = "|"
      '  .TextFileColumnDataTypes = Array(2)
       ' .TextFileTrailingMinusNumbers = True
        '.Refresh BackgroundQuery:=False
'    End With
'End Sub
'Sub Macro2()
'
' Macro2 Macro
'
'
 '   Range("A21").Select
  '  ActiveWorkbook.RefreshAll
'End Sub

Private Sub ImportarCV()

Dim tmp
Dim nom
Set hoja = Workbooks("metacps.xlsm").Sheets("Hoja1")
hoja.Activate
origDIR = Application.ActiveWorkbook.Path
'MsgBox (origDIR)
'Sheets("Hoja1").Activate

If NomAr = "" Then
NomAr = Application.GetOpenFilename("Archivos de texto, *.txt")
Else
NomAr = origDIR & NomAr
End If
tmp = NomAr
'MsgBox (tmp)
If tmp <> False Then
    Sheets("Hoja1").UsedRange.Clear
    nom = Dir(NomAr)
    nom = Mid$(nom, 1, Len(nom) - 4)
    
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & NomAr, Destination:=Range("$A$1"))
        .Name = nom
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "|"
        .TextFileColumnDataTypes = Array(2)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    End If
    'ActiveWorkbook.RefreshAll
NomAr = ""
End Sub

Private Sub guardar(ar As String)

  Cells.Replace What:=" ", Replacement:="<SP>", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
Dim r As Range
Dim nom As String
nom = Environ("HOMEPATH") & "\Documents\iMacros\Datasources\" & ar & ".csv"

'MsgBox (nom)

'Dim sFileSaveName As Variant
'Dim r As Integer
'r = 0

'guardarDialog:
'If r > 0 Then
'sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=nom)
'nom = sFileSaveName
'End If

On Error GoTo 0
'If nom <> "" And nom <> "False" Then
    If Len(Dir(nom)) Then
    'comentar temporalmente descomentar en versión final
        Select Case MsgBox(Dir(nom) & " ya existe. ¿Quiere sobreescribirlo?", vbYesNoCancel + vbInformation)
            Case vbYes
                'Application.DisplayAlerts = False
                'ActiveWorkbook.SaveAs Filename:=nom, FileFormat:=xlCSV 'descomentar para guardar sin adjuntar
                'adjuntar en lugar de guardar
                Set r = Range("2:2").SpecialCells(xlCellTypeConstants)
                'Append2CSV nom, r
                dumprangetocsv nom, r
                Application.DisplayAlerts = True
            Case vbNo
                'r = 1
                'GoTo guardarDialog
                ActiveWorkbook.Close
            Case vbCancel
                 ActiveWorkbook.Close savechanges:=False
                 Exit Sub
        End Select
    Else
        ActiveWorkbook.SaveAs Filename:=nom, FileFormat:=xlCSV
        'Descomentar para guardar y cerrar
        ''ActiveWorkbook.Close Savechanges = True
        MsgBox ("Documento guardado")
    End If
'Else
    'ActiveWorkbook.Close Savechanges:=False
'End If

'Descomentar para guardar solamente
'ActiveWorkbook.SaveAs Filename:=nom, FileFormat:=xlCSV
'Descomentar para guardar y cerrar
''ActiveWorkbook.Close Savechanges = True
'ActiveWorkbook.Close
ActiveWorkbook.Close (False)
''MsgBox ("Documento guardado")
'Workbooks("cv_a_csv").Activate

'vaciar variables
Set gcell = Nothing
Set wb = Nothing
NomAr = ""
Set hoja = Nothing
'origXML = ""
'origDIR = ""
'nomsec = ""
errHandler:
    Exit Sub
    'If Err <> 0 Then
        'Exit Sub
     '   MsgBox Err.Description
    'End If
    On Error GoTo 0

End Sub

Private Sub nuevahoja()

'Crea nuevo libro con una hoja
Set wb = Workbooks.Add(xlWBATWorksheet)
'copiar hoja existente
'Sheets("Hoja2").Activate
'Sheets("Hoja2").UsedRange.Clear
'ThisWorkbook.Sheets("Hoja2").Copy

End Sub

Sub Autores()

'Dim gcell As Range
'Dim wb As Workbook
NomAr = "\meta_aut.txt"
ImportarCV

'Range("A1").EntireRow.Insert
'Cells(1, "A").Value = "datosGeneralesrfc"
'Cells(1, "B").Value = "datosGeneralesestadoCivilVO"
'Set gcell = ActiveSheet.Cells.Find("RFC", MatchCase:=True)
'Cells(2, "A").Value = Trim(gcell.Value)
'Set gcell = ActiveSheet.Cells.Find("estado civil", MatchCase:=False)
'Cells(2, "B").Value = Trim(gcell.Value)

'Método 2 (funcional)
'Sheets("Hoja2").Activate
'Sheets("Hoja2").UsedRange.Clear
'Range("A1").Value = "datosGeneralesrfc"
'Range("B1").Value = "datosGeneralesestadoCivilVO"
'Set gcell = Sheets("Hoja1").Cells.Find("RFC", MatchCase:=True)
'Cells(2, "A").Value = Trim(gcell.Value)
'Set gcell = Sheets("Hoja1").Cells.Find("estado civil", MatchCase:=False)
'Cells(2, "B").Value = Trim(gcell.Value)
'Columns("A:B").Select
'Selection.EntireColumn.AutoFit
'ThisWorkbook.Sheets("Hoja2").Copy
''Dim nom As String
'nom = Environ("HOMEPATH") & "\Documents\iMacros\Datasources\DatosI.csv"
''MsgBox (nom)
'ActiveWorkbook.SaveAs Filename:=nom, FileFormat:=xlCSV
''ActiveWorkbook.Close Savechanges = True
''MsgBox ("Documento guardado")

'Método 3
'Crea nuevo libro con una hoja
'Set wb = Workbooks.Add(xlWBATWorksheet)

niv = "nivel de estudios"
Set gcell = hoja.Rows.Find(niv, MatchCase:=False)
    If gcell Is Nothing Then
    MsgBox (niv & " no encontrado")
    niv = DoM
    Else
    niv = Trim(LCase(Cells(gcell.Row, gcell.Column + 1).Value))
   End If

nuevahoja
'buscar nombre del documento para msgbox.caption
Dim autorForm As DatosForm
Set gcell = hoja.Cells.Find("*.doc")
Dim docnum
docnum = gcell.Value

'Sheets("Hoja2").Activate
'Sheets("Hoja2").UsedRange.Clear
Range("A1").Value = "nombre"
Range("B1").Value = "apellido"
Range("C1").Value = "correo"
Range("D1").Value = "rev1"
Range("E1").Value = "rev1ape"

Dim ar
Dim aut
Dim col
Dim capt
Dim suf
Dim lastcol

Select Case niv
Case "doctorado"
Range("F1").Value = "rev2"
Range("G1").Value = "rev2ape"
Range("H1").Value = "rev3"
Range("I1").Value = "rev3ape"

'Set gcell = Workbooks("metacps.xlsm").Sheets("Hoja1").Rows.Find("autor/a", MatchCase:=False)
'Cells(2, "A").Value = Trim(hoja.Cells(gcell.Row, "B").Value)
'Cells(2, "B").Value = Trim(hoja.Cells(gcell.Row, "C").Value)
'Set gcell = hoja.Rows.Find("tutor principal", MatchCase:=False)
'Cells(2, "D").Value = Trim(hoja.Cells(gcell.Row, "B").Value)
'Cells(2, "E").Value = Trim(hoja.Cells(gcell.Row, "C").Value)

ar = Array("autor/a", "tutor principal", "tutor adjunto", "tutor externo")
'ar(0, 0) = "autor/a"
'ar(1, 0) = "tutor principal"
'ar(1, 1) = "tutor principal"
'ar(2, 0) = "tutor adjunto"
'ar(2, 1) = "tutor adjunto"

'ar(3) = "tutor externo"
nomsec = "autorespsi"

Case "maestría"
ar = Array("autor/a", "tutor")
nomsec = "autM_cpsi"

End Select

For Each aut In ar
   Set autorForm = New DatosForm
    Set gcell = hoja.Rows.Find(aut, MatchCase:=False)
    If gcell Is Nothing Then
    MsgBox (aut & " no encontrado")
    Else
    Select Case aut
        Case "autor/a"
            col = Array("A", "B")
            capt = "Datos del autor/a"
            suf = " - Autor"
        'tutor maestría
        Case "tutor"
            col = Array("D", "E")
            capt = "Datos del tutor principal"
            suf = " - Tutor Principal"
        'tutores doctorado
        Case "tutor principal"
            col = Array("D", "E")
            capt = "Datos del tutor principal"
            suf = " - Tutor Principal"
        Case "tutor adjunto"
            col = Array("F", "G")
            capt = "Datos del tutor adjunto"
            suf = " - Tutor Adjunto"
        Case "tutor externo"
            col = Array("H", "I")
            capt = "Datos del tutor externo"
            suf = " - Tutor Externo"
    End Select
    Cells(2, col(0)).Value = Trim(hoja.Cells(gcell.Row, "B").Value)
    Cells(2, col(1)).Value = Trim(hoja.Cells(gcell.Row, "C").Value)
    With autorForm
    .Caption = docnum
    .lbCapt = capt
    .case_Sel = suf
    .Nom_in.Value = Cells(2, col(0)).Value
    .col1 = col(0)
    .Ape_in.Value = Cells(2, col(1)).Value
    .col2 = col(1)
    .Show
    End With
    End If
Next aut

lastcol = Cells(1, Columns.Count).End(xlToLeft).Column
Cells(1, lastcol + 1).Value = "correv"
Cells(1, lastcol + 2).Value = "id"

Set gcell = Workbooks("metacps.xlsm").Sheets("Hoja1").Cells.Find("*@*", MatchCase:=False)
Cells(2, "C").Value = Trim(gcell.Value)
Cells(2, lastcol + 1).Value = "correo11@example.com"
Set gcell = hoja.Cells.Find("id:", MatchCase:=False)
Cells(2, lastcol + 2).Value = Trim(hoja.Cells(gcell.Row, "B").Value)
'DatosForm.Show

Columns("A:K").Select
Selection.EntireColumn.AutoFit
'ThisWorkbook.Sheets("Hoja2").Copy

'Reemplazar por sub guardar
''Dim nom As String
'nom = Environ("HOMEPATH") & "\Documents\iMacros\Datasources\DatosI.csv"
''MsgBox (nom)
'ActiveWorkbook.SaveAs Filename:=nom, FileFormat:=xlCSV
'Descomentar para guardar y cerrar
''ActiveWorkbook.Close Savechanges = True
''MsgBox ("Documento guardado")
'Workbooks("cv_a_csv").Activate
guardar nomsec

End Sub

Sub Titulo2()

'Dim gcell As Range
'Dim wb As Workbook
Dim defval, val As String
Dim tmp, mail As String
Dim bPral As Boolean
Dim i, j As Long
'Dim corform As DatoscorForm
'Dim telform As DatostelForm
Dim resForm As TyRForm

ImportarCV
'Set wb = Workbooks.Add(xlWBATWorksheet)
'Set hoja = Workbooks("cv_a_csv").Sheets("Hoja1")
nuevahoja
'Sheets("Hoja2").Activate
'Sheets("Hoja2").UsedRange.Clear
'ThisWorkbook.Sheets("Hoja2").Copy
bPral = False
Range("A1").Value = "Titulo"
'Range("A2").Value = "COR"
Range("B1").Value = "Resumen"

Dim arrnotes

'funciona
'Set gcell = Workbooks("cv_a_csv").Sheets("Hoja1").Cells.Find("@")
''defval = gcell.Value
'j = 2
'arrnotes = Split(gcell, " ")
'For I = LBound(arrnotes) To UBound(arrnotes)
'MsgBox (arrnotes(I))
 '   If Trim(arrnotes(I)) <> "" Then
  '  tmp = arrnotes(I)
   '     If tmp Like "*@*.*" Then
        
        'mail = tmp
        'val = InputBox("Introduzca los valores correctos", "DatosII", tmp)
        'If user has clicked Cancel, set myValue to defaultValue
        'If val = "" Then val = tmp
        ''MsgBox (val & j)
        'Cells(j, "C").Value = Trim(val)
        'ofic = MsgBox("¿Es " & val & " una cuenta oficial?", vbYesNo + vbQuestion, "Característica de contacto")
        ''bloquear y hacer respuesta obligatoria
        'If ofic = vbYes Then
         '   Cells(j, "B").Value = "Oficial"
          '  Else
           ' Cells(j, "B").Value = "Personal"
        'End If
        'Cells(j, "D").Value = "NO"
        'If bpral = False Then
         '   pral = MsgBox("¿Es " & val & " su contacto principal?", vbYesNo + vbQuestion, "Medio de contacto principal")
          '  'bloquear y hacer respuesta obligatoria
           ' If pral = vbYes Then
            'Cells(j, "D").Value = "YES"
            'bpral = True
            'End If
           
'método II
'Dim hoja As Worksheet
'Set hoja = Workbooks("cv_a_csv").Sheets("Hoja1")
'Set gcell = hoja.Cells.Find("@")
With hoja.UsedRange
    Set gcell = .Cells.Find("@")
    j = 2
    If Not gcell Is Nothing Then
        Dim inicio
        inicio = gcell.Address
        'Do Until gcell Is Nothing
        Do
        arrnotes = Split(gcell, " ")
        For i = LBound(arrnotes) To UBound(arrnotes)
            'MsgBox (arrnotes(I))
            If Trim(arrnotes(i)) <> "" Then
                tmp = arrnotes(i)
                If tmp Like "*@*.*" Then
                    Cells(j, "A").Value = "COR"
                    Cells(j, "C").Value = tmp
                    Set corform = New DatoscorForm
                    With corform
                    .cont = j
                    .correo_in.Value = Trim(tmp)
                    .Show
        'DatoscorForm.cont = j
        'DatoscorForm.Show
                    End With
                    Set corform = Nothing
                    j = j + 1
                End If
            End If
        Next i
        Set gcell = .FindNext(gcell)
        Loop While Not gcell Is Nothing And gcell.Address <> inicio
    End If
End With
Set gcell = Nothing

'j = 2
'arrnotes = Split(gcell, " ")
'For I = LBound(arrnotes) To UBound(arrnotes)
'MsgBox (arrnotes(I))

'teléfono
Range("E1").Value = "medioContactoDTOmedioContactngchangedatamedioContactoDTOtelefononulldatamedioContactoDTOcorreonull"
'Range("E2").Value = "TEL"
'Range("E3").Value = "MOV"
Range("F1").Value = "medioContactoDTOcatCaracteristContactoVO"
Range("G1").Value = "medioContactoDTOtelefono "
Range("H1").Value = "medioContactoDTOesPrincipal"

Dim c
With hoja.UsedRange
    j = 2
    Set gcell = .Range("A1:G10").SpecialCells(xlCellTypeConstants)
    For Each c In gcell
    
    'For Each C In .Range("A1:G10")
    'MsgBox (C.Value)
    If c.Value Like "*####*" Then
        'MsgBox (C.Value)
        tmp = c.Value
        Cells(j, "G").Value = tmp
        Set telform = New DatostelForm
        With telform
            .cont = j
            .tel_in.Value = Trim(tmp)
            .Show
        End With
        Set telform = Nothing
        j = j + 1
    End If
    Next
End With
'Set gcell = Nothing
nomsec = "DatosII"
guardar nomsec
'nom = Environ("HOMEPATH") & "\Documents\iMacros\Datasources\DatosII.csv"
'MsgBox (nom)
'ActiveWorkbook.SaveAs Filename:=nom, FileFormat:=xlCSV

End Sub

Sub Titulo()

'Dim gcell As Range
Dim wb As Workbook
'Dim hoja As Worksheet
Set hoja = Workbooks("metacps.xlsm").Sheets("Hoja1")

Dim resForm As TyRForm

NomAr = "\meta_res.txt"

ImportarCV
nuevahoja

Range("A1").Value = "Titulo"
Range("A2").Value = Trim(hoja.Cells(1, "B").Value)
Range("B1").Value = "Resumen"
Range("B2").Value = Trim(hoja.Cells(1, "C").Value)
Set resForm = New TyRForm
resForm.Show

If niv = "" Then
    niv = DoM
End If

Select Case niv
    Case "doctorado"
        nomsec = "titulocpsi"
    Case "maestría"
        nomsec = "tituloM_cpsi"
End Select

guardar nomsec

End Sub

Sub Index()

Dim wb As Workbook
'Dim hoja As Worksheet
Set hoja = Workbooks("metacps.xlsm").Sheets("Hoja1")
Dim indfrm As indexForm

NomAr = "\meta_key.txt"

ImportarCV
nuevahoja

Set gcell = hoja.Cells.Find("palabras claves", MatchCase:=False)
Range("A1").Value = "Palabras clave"
Range("A2").Value = Trim(hoja.Cells(gcell.Row + 1, "A").Value)

Set gcell = hoja.Cells.Find("fuentes de financiamiento", MatchCase:=False)
Range("B1").Value = "Patrocinio"
If gcell Is Nothing Then
MsgBox ("Falta 'Patrocinio'")
Else
Range("B2").Value = Trim(hoja.Cells(gcell.Row + 1, "A").Value)
End If
Set indfrm = New indexForm
indfrm.Show

If niv = "" Then
    niv = DoM
End If

Select Case niv
    Case "doctorado"
        nomsec = "keycpsi"
    Case "maestría"
        nomsec = "keyM_cpsi"
End Select

guardar nomsec

End Sub

Sub callUF()

'Dim ofrm As GAform
Dim ofrm As AreaForm
origDIR = Application.ActiveWorkbook.Path
origXML = origDIR & "\areaminf.xml"
'MsgBox (CurDir)
    If LoadDataPass Then
      'Set ofrm = New GAform
       Set ofrm = New AreaForm
       ofrm.Show
       Unload ofrm
       Set ofrm = Nothing
    End If
    
lbl_Exit:
   Exit Sub
End Sub

Function LoadDataPass() As Boolean

Dim xmldoc As New MSXML2.DOMDocument60

   xmldoc.validateOnParse = True
   xmldoc.async = False
    If xmldoc.Load(origXML) Then
       LoadDataPass = True
   Else
       LoadDataPass = False
       Dim strErr As String
       Dim xPE As MSXML2.IXMLDOMParseError
       Set xPE = xmldoc.parseError
       With xPE
        strErr = "Error al cargar el archivo" & vbCrLf & _
            "Error #: " & .ErrorCode & ": " & xPE.reason & _
            "Linea #: " & .Line & vbCrLf & _
            "Posición: " & .linepos & vbCrLf & _
            "Orígen: " & .srcText
        End With
        
        MsgBox strErr, vbExclamation
    End If
    Set xPE = Nothing

lbl_Exit:
    Exit Function
    
End Function

Private Sub Append2CSV(CSVfile As String, CellRange As Range)

Dim tmpCSV As String
Dim f As Integer

f = FreeFile

Open CSVfile For Append As #f
    tmpCSV = Range2CSV(CellRange)
Write #f, tmpCSV
Close #f

End Sub

Function Range2CSV(list) As String

Dim tmp As String
Dim cr As Long
Dim r As Range
If TypeName(list) = "Range" Then
    cr = 1
    For Each r In list.Cells
        If r.Row = cr Then
            If tmp = vbNullString Then
            tmp = r.Value
            Else
            tmp = tmp & "," & r.Value
            End If
        Else
        cr = cr + 1
            If tmp = vbNullString Then
            tmp = r.Value
            Else
            tmp = tmp & Chr(10) & r.Value
            End If
        End If
    Next
End If
Range2CSV = tmp

End Function

Private Sub dumprangetocsv(CSVfile As String, source_range As Range)

Dim tmpCSV As String
Dim row_range As Range, mycell As Range
Dim f As Integer
f = FreeFile

Open CSVfile For Append As #f
    For Each row_range In source_range.Rows
        For Each mycell In row_range.Cells
            Write #f, mycell.Value,
        Next mycell
        Write #f,
    Next row_range
Close #f

End Sub

Function DoM() As String

'niv = "nivel de estudios"
'Set gcell = hoja.Rows.Find(niv, MatchCase:=False)
'    If gcell Is Nothing Then
'    MsgBox (niv & " no encontrado")
'    Else
'    niv = Trim(LCase(hoja.Cells(gcell.Row, gcell.Column + 1).Value))
'    End If
    Select Case MsgBox("Doctorado", vbYesNoCancel + vbInformation)
            Case vbYes
                niv = "doctorado"
            Case vbNo
                niv = "maestría"
    End Select

DoM = niv

End Function
Sub Meta()

Autores
Titulo
Index
niv = ""

End Sub
