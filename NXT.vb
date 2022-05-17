Option Explicit

Sub Dung_La_Trai_Dat_Tron_aHihi()

    Dim dict As Object
    Dim Data As Variant, Result As Variant
    Dim sKey As String, sID As String
    Dim i As Long, j As Long, k As Long
    Dim issueDate As Date
   
    Const startDate As Date = #8/1/2021#
    Const endDate As Date = #8/31/2021#

    Data = ThisWorkbook.Worksheets("DATABASE").Range("A1").CurrentRegion.Value
    If Not IsArray(Data) Then Exit Sub
   
    ReDim Result(1 To UBound(Data, 1), 1 To UBound(Data, 2))
    Set dict = CreateObject("Scripting.Dictionary")
    k = 1
    '//Tieu de
    For j = LBound(Data, 2) To UBound(Data, 2)
        Result(k, j) = Data(1, j)
    Next j
   
    '//Kiem tra nhap-xuat den thoi diem tinh den startDate
    For i = 2 To UBound(Data, 1)
        sID = Data(i, 2)
        issueDate = Data(i, 7)
        If issueDate < startDate Then
            If Not dict.Exists(sID) Then dict.Add sID, issueDate
            sKey = sID & "|" & Data(i, 6)
            If Not dict.Exists(sKey) Then dict.Add sKey, issueDate
        End If
    Next i
   
    For i = 2 To UBound(Data, 1)
        sID = Data(i, 2)
        issueDate = Data(i, 7)
        '// Xet trong khoang startDate den endDate
        If (issueDate >= startDate) And (issueDate <= endDate) Then
            k = k + 1
            For j = LBound(Data, 2) To UBound(Data, 2)
                Result(k, j) = Data(i, j)
            Next j
        ElseIf (issueDate <= endDate) Then
            If dict.Exists(sID) Then
                sKey = sID & "|X"
                '// Neu co nhap ma chua co xuat
                If Not dict.Exists(sKey) Then
                    k = k + 1
                    For j = LBound(Data, 2) To UBound(Data, 2)
                        Result(k, j) = Data(i, j)
                    Next j
                End If
            End If
        End If
    Next i
   
    ThisWorkbook.Worksheets("Result").Range("A1").Resize(k, UBound(Result, 2)).Value = Result
   
    MsgBox k
   
End Sub



