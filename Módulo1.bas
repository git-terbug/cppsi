Attribute VB_Name = "Módulo1"

Public Sub CargarDoc()

    Dim dAbr As FileDialog, result As Integer, it As Variant
    Set dAbr = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    With dAbr
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Documentos de Word", "*.docx;*.doc"
        .Filters.Add "Todos los archivos", "*.*"
    End With
    Dim str As String
   
    If dAbr.Show = -1 Then
      'ActiveDocument.StoryRanges(wdMainTextStory).Delete
      cls = LimpiarDoc()
    str = Dir(dAbr.SelectedItems(1))
      Selection.InsertFile (dAbr.SelectedItems(1))
      ScratchMacro
      'Application.Documents.Open (dAbr.SelectedItems(1))
      'Dim tmp
      'Selection.WholeStory
      'tmp = test
      'ActiveDocument.SaveAs FileName:=ThisDocument.Path & "\" & tmp & "_" & "01" & ".txt", FileFormat:=wdFormatText
     'Application.Selection.Copy
       ' ActiveDocument.SaveAs2 "test.txt", 2
      'Application.ActiveWindow.Close
      
      'ActiveDocument.Range.InsertBefore str & vbCrLf
      marcar str
      formatear
      
    End If
    
End Sub

Private Sub ScratchMacro()

Dim aShape As Shape
'Dim oILS As InlineShape
Dim oRng As Range
'Dim oanchor As Range
Dim strTex As String
Dim j As Long
Dim k As Long

For Each oRng In ActiveDocument.StoryRanges
    j = oRng.ShapeRange.Count
    If j > 0 Then
    For k = j To 1 Step -1
'For Each aShape In ActiveDocument.Shapes
    Set aShape = oRng.ShapeRange(k)
    If aShape.Type = msoAutoShape Then
        strtext = Trim(aShape.TextFrame.TextRange.Text)
        If Len(strtext) > 0 Then
            'Set oanchor = aShape.anchor.Paragraphs(1).Range
            oRng.InsertBefore strtext
        End If
        aShape.Delete
        'Set oILS = aShape.ConvertToFrame
        'Set oRng = oILS.Range
        'oILS.Delete
'        Selection.TypeText strText
    End If
    Next k
    End If
Next oRng
End Sub

Private Sub marcar(nom As String)
    
    Dim r As Range
    Dim titulo
    Dim str
    Set r = ActiveDocument.Range
    
    r.Select
    With Selection.Find
    .ClearFormatting
    .Text = "*Campo"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchWildcards = True
    .MatchCase = False
    .Execute
    If .Found Then
    'While .Found
    Selection.Range.Case = wdUpperCase
    '.Execute
    'Wend
    Selection.Range.Case = wdTitleSentence
    '.Execute
    titulo = Selection.Text
    titulo = Mid(titulo, 1, Len(titulo) - 6)
    End If
    End With
    
    r.InsertBefore nom & vbCrLf
    
    r.Select
    With Selection.Find
    .Text = "[0-9]{1,}-*"
    .Forward = True
    .Wrap = wdFindContinue
    .MatchWildcards = True
    End With
    Selection.Find.Execute
    str = Selection.Text
    str = Mid(str, 1, Len(str) - 1)
    r.InsertBefore "Id: | " & str & vbCrLf
    r.InsertBefore "##aut" & vbCrLf
    
    'r.Select
    ''Selection.ClearParagraphAllFormatting
    ''Selection.ClearFormatting
    ''Selection.ClearParagraphStyle
    'Selection.Find.Replacement.ClearFormatting
    
    'With Selection.Find
     '   .ClearFormatting
      ' .Text = "(^13){2,}"
       ' '.Replacement.Text = "^s"
        '.Replacement.Text = "^13"
        '.Forward = True
        '.Wrap = wdFindContinue
        '.Format = False
        '.Execute Replace:=wdReplaceAll
    'End With
    ''Do While Selection.Find.Execute
    ''Selection.Delete
    ''Loop
    
    Dim elem
    str = Array("autor/a del resumen:", "correo electrónico:", "dependencia:", _
        "tutor principal del alumno:", "tutor adjunto del alumno:", "tutor externo del alumno:")
    For Each elem In str
    r.Select
    With Selection.Find
    .ClearFormatting
    .Wrap = wdFindContinue
    .Format = False
    .Text = elem
    .MatchCase = False
    .MatchWildcards = False
    .Execute
    If .Found Then
    Selection.InsertAfter " |"
    Else
    MsgBox ("No se encontró " & elem)
    End If
    End With
    Next
    
    r.Select
    With Selection.Find
    '.Text = "ESQUEMA DE PRESENTACIÓN DEL RESUMEN"
    .Text = "esquema de presentación del resumen"
    .MatchCase = False
    .Execute
    If .Found Then
    Selection.InsertAfter vbCrLf & "##res"
    Selection.InsertAfter vbCrLf & "| " & titulo & " |"
    End If
    End With
    
    r.Select
    With Selection.Find
    .Text = "palabras clave"
    .Execute
    End With
    Selection.InsertBefore "##key" & vbCrLf
    
    
End Sub

Private Sub formatear()

Dim r As Range
Set r = ActiveDocument.Range
    
r.Select
With Selection.Range.Find
    .MatchWildcards = True
    .ClearFormatting
    .Replacement.ClearFormatting
    .Forward = True
    .Wrap = wdFindContinue
    .Format = False
    .Text = "[ ^s^t]{1,}^13"
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll
    '.Text = "([!^13])([^13])([!^13])"
    '.Replacement.Text = "\1 \3"
    '.Execute Replace:=wdReplaceAll
    .Text = "[ ]{2,}"
    .Replacement.Text = " "
    .Execute Replace:=wdReplaceAll
    
    '.Text = Space(2)
    '.Replacement.Text = Space(1)
    '.Execute Replace:=wdReplaceAll
    
'    .Text = "([A-Za-z])( {2,})"
 '   .Replacement.Text = "\1 "
  '  .Execute Replace:=wdReplaceAll
  
    .Text = "(^13){2,}"
    .Replacement.Text = "^13"
    .Execute Replace:=wdReplaceAll
    
End With
End Sub

Sub DivDoc(delim As String, strNomAr As String)

    Dim doc As Document
    Dim arrNotes
    Dim i As Long
    'Dim X As Long
    Dim nsec As Collection
    Dim sec As String
    Dim lsec As String
    Dim Response As Integer
    'Dim r As String
    Dim r As Range
    lsec = ""
    Set nsec = New Collection
    Set r = ActiveDocument.Range
    r.Select
    With Selection.Find
        .ClearFormatting
        .Text = "##[a-z]{1,10}"
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
    End With
    Do While Selection.Find.Execute
    lsec = lsec & Selection.Text & vbCrLf
    nsec.Add Selection.Text
    Loop
    Response = MsgBox("Dividir el documento en " & nsec.Count & " secciones: " & vbCrLf & lsec & " ¿Desea continuar?", 4)
    
    arrNotes = Split(ActiveDocument.Range, delim)
    'contar etiquetas no vacias ##w
    'nsec = FindTags(r)
    'lsec = ""
    'Dim it As Variant
    'For Each it In nsec
     '   lsec = lsec & it & vbCrLf
    'Next it
    'Response = MsgBox("Dividir el documento en " & nsec.Count & " secciones: " & vbCrLf & lsec & " ¿Desea continuar?", 4)
'Response = MsgBox("Dividir el documento en " & UBound(arrNotes) + 1 & " secciones. ¿Desea continuar?", 4)
If Response = 7 Then Exit Sub
Dim ns As Long
ns = 1

For i = LBound(arrNotes) To UBound(arrNotes)
If Trim(arrNotes(i)) <> "" Then
Dim tmp As String
tmp = Left(arrNotes(i), 1)
'MsgBox (tmp)
   If tmp Like "[a-zA-Z]" Then
       'X = X + 1
    sec = Mid$(nsec.Item(ns), 3)
    'MsgBox ("sección: " & sec)
    Set doc = Documents.Add
    doc.Range = arrNotes(i)
    
    If sec = "res" Then
    doc.Range.Select
    'Selection.ClearFormatting
    With Selection.Find
        .ClearFormatting
        '.Text = "(^13){2,}"
        .Text = "^13"
        .Replacement.Text = "^s"
        '.Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    End If
    
    'doc.SaveAs ThisDocument.Path & "\" & strNomAr & Format(X, "000")
    'doc.SaveAs ThisDocument.Path & "\" & strNomAr & "_" & sec
    doc.SaveAs FileName:=ThisDocument.Path & "\" & strNomAr & "_" & sec & ".txt", FileFormat:=wdFormatText
    doc.Close True
 
     If ns < nsec.Count Then
     ns = ns + 1
     End If
    'Else
     'MsgBox ("borrar sección " & I)
    End If
End If
Next i
MsgBox ("Documentos guardados")
End Sub

Sub SepararDoc()

'delimiter & filename
DivDoc "##", "meta"
'deseas guardar el documento
AbrirExcel
cls = LimpiarDoc()
Selection.Font.Name = "Courier New"
Selection.Font.Bold = True
Selection.Font.Size = 16
Selection.TypeText ("Antes de continuar asegúrate de que las macros estén habilitadas" & vbCrLf _
                & vbCrLf & "Instrucciones: " & vbCrLf & "Presiona alt + f8" & vbCrLf & "Selecciona CargarDoc" _
                & vbCrLf & "Separa las secciones con ## y añade una etiqueta (ej: ##datos) o déjala en blanco para ignorarla" _
                & vbCrLf & "Presiona alt + f8" & vbCrLf & "Selecciona SepararDoc")
End Sub

Private Function LimpiarDoc()

Dim clSec As Section
Dim hd_ft As HeaderFooter

ActiveDocument.StoryRanges(wdMainTextStory).Delete
For Each clSec In ActiveDocument.Sections
        For Each hd_ft In clSec.Headers
            hd_ft.Range.Delete
        Next
        For Each hd_ft In clSec.Footers
            hd_ft.Range.Delete
        Next
    Next clSec

End Function

'Public Function FindTags(ByRef r As String) As Collection


 '   Dim oMatch As Object
  '  Set FindTags = New Collection
    
   ' With CreateObject("VBScript.RegExp")
    '    .Global = True
     '   .Pattern = "##\w+"
        
      '  For Each oMatch In .Execute(r)
       ' FindTags.Add oMatch.Value
        'Next
    'End With
'End Function

'Public Function IsLetter(strV As String) As Boolean

'Dim inPos As Integer

'IsLetter = True
'IsLetter = (Mid$(strV,1,1)
'Do Until IsLeter = False Or

'End Function

Private Sub AbrirExcel()

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oSheet As Excel.Worksheet
Dim oRng As Excel.Range
Dim ExcelNoCorre As Boolean
Dim NomExc As String

NomExc = ThisDocument.Path & "\metacps.xlsm"

On Error Resume Next
Set oXL = GetObject(, "Excel.Application")

If Err Then
    ExcelNoCorre = True
    Set oXL = New Excel.Application
End If

On Error GoTo Err_Handler

oXL.Visible = True
Set oWB = oXL.Workbooks.Open(FileName:=NomExc)

'Process each of the spreadsheets in the workbook
'For Each oSheet In oXL.ActiveWorkbook.Worksheets
   'put guts of your code here
   'get next sheet
'Next oSheet

'If ExcelWasNotRunning Then
 ' oXL.Quit
'End If

'Make sure you release object references.
'Set oRng = Nothing
'Set oSheet = Nothing
'Set oWB = Nothing
'Set oXL = Nothing

Exit Sub

Err_Handler:
MsgBox NomExc & "causó un error inesperado. " & Err.Description, vbCritical, "Error: " & Err.Number
If ExcelNoCorre Then
    oXL.Quit
End If
End Sub
