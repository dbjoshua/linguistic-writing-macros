Sub Ling_Matrix_Make_draft()
'
' InterlinearTestMacro Macro
'
'

' Declaring variables
Dim mtxRng As Range
Dim mtxLines As Table
Dim mtxLBr As Column
Dim mtxRBr As Column
' Dim cell As cell

If Selection.Paragraphs.Count < 2 Then
    Set mtxRng = Selection.Paragraphs(1).Range
    Else: 'Otherwise, assume user has selected whole ex.
         Set mtxRng = Selection.Range
End If

With mtxRng.Find
     .ClearFormatting
     .Text = ","
     .Replacement.ClearFormatting
     .Replacement.Text = ""
     .Execute Replace:=wdReplaceAll, Forward:=True, _
     Wrap:=wdFindStop
End With

With mtxRng.Find
     .ClearFormatting
     .Text = " "
     .Replacement.ClearFormatting
     .Replacement.Text = "^p"
     .Execute Replace:=wdReplaceAll, Forward:=True, _
     Wrap:=wdFindStop
End With

'Process matrix text
mtxTx = mtxRng.Text
mtxRng.Font.SmallCaps = wdToggle

'
' / Conversion de la selection en tableau (adapte de la macro de Susanna Cumming dans "LingWord.dot")
'
    Set mtxLines = mtxRng.ConvertToTable(Separator:=" ", AutoFit:=True, _
    AutoFitBehavior:=wdAutoFitContent, DefaultTableBehavior:=wdWord9TableBehavior)
'   Formattage du tableau
    With mtxLines
         .Borders.Enable = False
         .TopPadding = 0
         .BottomPadding = 0
         .LeftPadding = 0
         .RightPadding = 0
         .Spacing = 0
         .AllowPageBreaks = False
         .AllowAutoFit = True
    End With

'
' / Adding brackets columns
'

    Set mtxLBr = mtxLines.Columns.Add(BeforeColumn:=mtxLines.Columns(1))
    mtxLines.Columns(mtxLines.Columns.Count).Select
    Selection.InsertColumnsRight
    Set mtxRBr = mtxLines.Columns(mtxLines.Columns.Count)
    
'
' / Making brackets
'
    If mtxLines.Rows.Count = 1 Then
       With mtxLines
            .cell(1, 1).Range.InsertAfter "["
            .cell(1, mtxLines.Columns.Count).Range.InsertAfter "]"
       End With
     ElseIf mtxLines.Rows.Count = 2 Then
       With mtxLines
            .cell(1, 1).Range.Select
            Selection.InsertSymbol Font:="Charis SIL", CharacterNumber:=9121, Unicode _
        :=True ' Coin superieur de crochet gauche
            .cell(mtxLines.Rows.Count, 1).Range.Select
            Selection.InsertAfter "|"
            .cell(1, mtxLines.Columns.Count).Range.Select
            Selection.InsertAfter "|"
            .cell(mtxLines.Rows.Count, mtxLines.Columns.Count).Range.Select
            Selection.InsertAfter "|"
       End With
    End If
    
    
    
End Sub
