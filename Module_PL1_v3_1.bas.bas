Attribute VB_Name = "Module1_PL1_v3_0"

Type ProcStackType
    procName As Variant
    procNum  As Variant
End Type


Public Sub MclSetC()
Dim resultSheet As Worksheet
Dim analyzeSheet As Worksheet
Dim testSheet As Worksheet
Dim currentRow As Variant
Dim sourceRng As Range
Dim deleteTestRng As Range
Dim startRow As Long
Dim maxRow As Long
Dim row As Long
Dim pos As Integer
Dim currentStatus As Integer
Dim items As Variant
Dim itemCnt As Integer
Dim outline As Variant
Dim tmp As String

    Set resultSheet = Worksheets("比較結果")
    Set analyzeSheet = Worksheets("UT Case ID 採番シート")
    Set testSheet = Worksheets("検討")

    resultSheet.Activate
    startRow = 3
    maxRow = ActiveSheet.Range("E3").End(xlDown).row
    Set sourceRng = resultSheet.Range(GetColRange("D", startRow, maxRow))
    Set deleteTestRng = testSheet.Range(GetColRange("A", startRow, maxRow))
    
    row = 1
    currentStatus = 0
    For Each currentRow In sourceRng
        outline = ""
        tmp = Trim(currentRow)
        items = Split(tmp)
        itemCnt = UBound(items)
        For idx = 0 To itemCnt
            If Trim(items(idx)) = "" Then
            Else
                 currentStatus = ChangeStatus(currentStatus, items(idx), outline)
            End If
        Next
        Range(”解析テーブル[比較結果_変更後ソース_コメント文除去]”).Row(row) = outline
        row = row + 1
    Next currentRow
End Sub

Function ChangeStatus(ByRef status As Integer, ByRef item As Variant, ByRef outline As Variant) As Integer
     If status = 0 Then
        ChangeStatus = Normalline(item, outline)
     ElseIf status = 1 Then
        ChangeStatus = SearchCommentEnd(item, outline)
     ElseIf status = 2 Then
        ChangeStatus = SearchDoubleQuote(item, outline)
     Else
        ChangeStatus = SearchSingleQuote(item, outline)
     End If
End Function

Function Normalline(ByRef item As Variant, ByRef outline As Variant) As Integer
Dim pos As Integer
    Normalline = 0
    If Len(item) = 0 Then
        Exit Function
    End If
    pos = InStr(item, "/*")
    If pos > 0 Then
        outline = outline & Left(item, pos - 1) & " "
        Normalline = ChangeStatus(1, Mid(item, pos + 2), outline)
        Exit Function
    End If
    pos = InStr(item, """")
    If pos > 0 Then
        outline = outline & Left(item, pos - 1) & " "
        Normalline = ChangeStatus(2, Mid(item, pos + 1), outline)
        Exit Function
    End If
    pos = InStr(item, "'")
    If pos > 0 Then
        outline = outline & Left(item, pos - 1) & " "
        Normalline = ChangeStatus(3, Mid(item, pos + 1), outline)
        Exit Function
    End If
    outline = outline & item & " "
End Function

Function SearchCommentEnd(ByRef item As Variant, ByRef outline As Variant) As Integer
Dim pos As Integer
    SearchCommentEnd = 1
    pos = InStr(item, "*/")
    If pos > 0 Then
        item = Mid(item, pos + 2)
        SearchCommentEnd = ChangeStatus(0, item, outline)
        Exit Function
    End If
End Function

Function SearchDoubleQuote(ByRef item As Variant, ByRef outline As Variant) As Integer
Dim pos As Integer
    SearchDoubleQuote = 2
    pos = InStr(item, """")
    If pos > 0 Then
        item = Mid(item, pos + 1)
        SearchDoubleQuote = ChangeStatus(0, item, outline)
        Exit Function
    End If
End Function

Function SearchSingleQuote(ByRef item As Variant, ByRef outline As Variant) As Integer
Dim pos As Integer
    SearchSingleQuote = 3
    pos = InStr(item, "'")
    If pos > 0 Then
        item = Mid(item, pos + 1)
        SearchSingleQuote = ChangeStatus(0, item, outline)
        Exit Function
    End If
End Function

Public Sub GetProc()

End Sub

Function GetColRange(ByVal colStr As String, ByVal startRow As Long, ByVal endRow As Long) As String
    GetColRange = colStr & Trim(Str(startRow)) & ":" & colStr & Trim(Str(endRow))
End Function

