Sub bold_keywords()
' Created by Vera
' This vba script colorizes keywords in a worksheet labeled
' condensed list using the words in the 1st column
' of a tab labeled keywords

Dim lRow As Long
Dim lCol As Long
    sheet_id = "Annex I"
    ref_row = 2
    ref_col = 2
    'Find the last non-blank cell in column of column condensed list tab
    lRow = Sheets(sheet_id).Cells(Rows.Count, ref_col).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Sheets(sheet_id).Cells(ref_row, Columns.Count).End(xlToLeft).Column
    
    ' set a range of cells that will be formatted
    Dim rng As Range
    Dim rng_col As Variant
    
    ' loop through the columns you will colorize
    For c = 1 To 4
        Set rng = Worksheets(sheet_id).Range(Sheets(sheet_id).Cells(ref_row + 2, ref_col + c), _
        Sheets(sheet_id).Cells(lRow, ref_col + c))
        
        rng.Select
        Size = rng.Count
        Call colorize(rng)
    Next c
    Debug.Print ("ran bold_keywords")
End Sub
Sub colorize(rng As Range)
' Color code words in cells that offer clues to the standard titles

'array of keywords
Dim key_words() As Variant
Dim i As Integer

'terms = Worksheets("keywords").Range("A1:A4")
lRowTerms = Worksheets("keywords").Cells(Rows.Count, 1).End(xlUp).Row

ReDim key_words(lRowTerms)

'store the keywords in an array
For i = 1 To lRowTerms
    word = Worksheets("keywords").Cells(i, 1)
    key_words(i) = word
Next

Dim rCell As Range, sToFind As String, iSeek As Long

' Colorize the cell
    For Each st In rng
        wd = st.Value
        wd = LCase(wd)
        
        For j = 1 To UBound(key_words)
                sToFind = key_words(j)
                iSeek = InStr(1, wd, sToFind)
            Do While iSeek > 0
                st.Characters(iSeek, Len(sToFind)).Font.Bold = True
                st.Characters(iSeek, Len(sToFind)).Font.Color = vbBlue
                iSeek = InStr(iSeek + 1, st.Value, sToFind)
            Loop
        Next j
    Next st

End Sub



