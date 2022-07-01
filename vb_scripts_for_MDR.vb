Sub bold_keywords()
' Created by Vera
' This vba script colorizes keywords in a worksheet labeled
' condensed list using the words in the 1st column
' of a tab labeled keywords

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column of column condensed list tab
    lRow = Sheets("Condensed List").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Sheets("Condensed List").Cells(2, Columns.Count).End(xlToLeft).Column
    
    Debug.Print ("Last Row: " & lRow & vbNewLine & _
            "Last Column: " & lCol)
    ' set a range of cells that will be formatted
    Dim rng As Range
    Set rng = Worksheets("Condensed List").Range(Sheets("Condensed List").Cells(2, 4), Sheets("Condensed List").Cells(lRow, 4))
    rng.Select
    Size = rng.Count
' Color code words in cells that offer clues to the standard titles

'array of keywords
Dim key_words() As Variant
Dim i As Integer

'terms = Worksheets("keywords").Range("A1:A4")
lRowTerms = Worksheets("keywords").Cells(Rows.Count, 1).End(xlUp).Row

Debug.Print (lRowTerms)
ReDim key_words(lRowTerms)

'store the keywords in an array
For i = 1 To lRowTerms
    word = Worksheets("keywords").Cells(i, 1)
    'Debug.Print ("wd " & word)
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
