'Sub InsertRows()
'
'End Sub
'    Dim I As Long, J As Integer, Nb As Integer
'
'    For I = Range("A65536").End(xlUp).Row To 2 Step -1
'
'             Nb = 4
'
'        For J = 1 To Nb - 1
'            Rows(I + J).Insert xlDown
'            Rows(I).Copy
'            Rows(I + J).PasteSpecial    '
'        Next
'
'    Next
'
'    Range("A1").Select
'    Application.CutCopyMode = False
'End Sub

Sub Insertdates()
    Dim I As Long, J As Integer, Nb As Integer
    I = Range("A65536").End(xlUp).Row
    cval = Cells(I, 1)
    cval = cval + 1
    final_date = Cells(2, 1) + 365
    I = I + 1
    While (cval <= final_date)
        For J = 0 To 23
            Cells(I, 1) = cval
            I = I + 1
        Next
        cval = cval + 1
    Wend
    Range("A1").Select
    Application.CutCopyMode = False
End Sub
Sub repeatRows()
    Dim I As Long, J As Integer, Nb As Integer
    I = 2
    T = 2 + 48
    cval = Cells(T, 1)
    While (cval <> "")
        Range(Cells(I, 2), Cells(I, 5)).Copy
        Range(Cells(T, 2), Cells(T, 5)).PasteSpecial
        I = I + 1
        T = T + 1
        cval = Cells(T, 1)
    Wend
    Range("A1").Select
    Application.CutCopyMode = False
End Sub
