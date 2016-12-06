Private LotteryCells(1 To 6) As Object
Private WinningCells(1 To 7) As Object
Private Counter As Object
Const Store_Row As Integer = 13
Const Store_Col As Integer = 1


''''''''''''''''''''''''''
'       Auto_Open        '
''''''''''''''''''''''''''

Private Sub Auto_Open()
    Dim i As Integer

    ' 設定 LotteryCells
    For i = LBound(LotteryCells) To UBound(LotteryCells)
        Set LotteryCells(i) = cells(8, i + 1)
    Next

    ' 設定 WinningCells
    For i = LBound(WinningCells) To UBound(WinningCells)
        Set WinningCells(i) = cells(3, i + 1)
    Next

    ' 設定計數器
    Set Counter = cells(11, 1)
End Sub


''''''''''''''''''''''''''
'        投注號碼        '
''''''''''''''''''''''''''

Public Sub 投注號碼_產生_click()
    Generate_Lottery_Cells LotteryCells
End Sub

Public Sub 投注號碼_檢查_click()
    Dim result As String
    result = Check_Lottery_Cells(LotteryCells)
    MsgBox IIf(result = "", "投注號碼 檢查正確", result)
End Sub

Public Sub 投注號碼_清除_click()
    For Each cell In LotteryCells
        cell.value = ""
    Next
End Sub

Public Sub 投注號碼_儲存_click()
    Dim result As String
    result = Check_Lottery_Cells(LotteryCells)
    If Check_Lottery_Cells(LotteryCells) <> "" Then
        MsgBox result
        Exit Sub
    End If

    Dim ArrayCells(0 To 6) As Object
    Dim i, j As Integer
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        Set ArrayCells(i) = cells(Counter.value + Store_Row, i + Store_Col)
    Next

    ' 設定編號
    ArrayCells(0) = Counter.value + 1

    ' 將 LotteryCells 的值複製到 ArrayCells
    For i = LBound(LotteryCells) To UBound(LotteryCells)
        ArrayCells(i).value = LotteryCells(i).value
    Next

    ' 由小到大排序 (氣泡排序)
    For i = 1 To 6
        For j = i + 1 To 6
            If ArrayCells(i).value > ArrayCells(j).value Then
                Dim tmp As Integer
                tmp = ArrayCells(i).value
                ArrayCells(i).value = ArrayCells(j).value
                ArrayCells(j).value = tmp
            End If
        Next
    Next

    ' 計數器遞增
    Counter.value = Counter.value + 1
End Sub


''''''''''''''''''''''''''
'        開獎號碼        '
''''''''''''''''''''''''''

Public Sub 開獎號碼_產生_click()
    Reset_Color
    Generate_Lottery_Cells WinningCells
End Sub

Public Sub 開獎號碼_檢查_click()
    Dim result As String
    result = Check_Lottery_Cells(WinningCells)
    MsgBox IIf(result = "", "開獎號碼 檢查正確", result)
End Sub

Public Sub 開獎號碼_清除_click()
    Reset_Color
    For Each cell In WinningCells
        cell.value = ""
    Next
End Sub

' Unused Function
Public Sub 開獎號碼_儲存_click()
    If Check_Lottery_Cells(WinningCells) = False Then
        Exit Sub
    End If

    Dim ArrayCells(0 To 7) As Object
    Dim i, j As Integer
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        Set ArrayCells(i) = cells(Counter.value + Store_Row, i + Store_Col)
    Next

    ' 設定編號
    ArrayCells(0) = Counter.value + 1

    ' 將 WinningCells 的值複製到 ArrayCells
    For i = LBound(WinningCells) To UBound(WinningCells)
        ArrayCells(i).value = WinningCells(i).value
    Next

    ' 由小到大排序 (氣泡排序)
    For i = 1 To 6
        For j = i + 1 To 6
            If ArrayCells(i).value > ArrayCells(j).value Then
                Dim tmp As Integer
                tmp = ArrayCells(i).value
                ArrayCells(i).value = ArrayCells(j).value
                ArrayCells(j).value = tmp
            End If
        Next
    Next

    ' 計數器遞增
    Counter.value = Counter.value + 1
End Sub


''''''''''''''''''''''''''
'          對獎          '
''''''''''''''''''''''''''

' Unused Function
Public Sub 對獎_click()
    If Check_Lottery_Cells(LotteryCells) = False Or Check_Lottery_Cells(WinningCells) = False Then
        Exit Sub
    End If

    Dim winning As Integer
    Dim special As Boolean
    winning = 0
    special = False

    ' 計算中獎數目和特別號
    Dim i, j As Integer
    For i = LBound(LotteryCells) To UBound(LotteryCells)
        For j = LBound(WinningCells) To UBound(WinningCells)
            ' 號碼相同
            If LotteryCells(i).value = WinningCells(j).value Then
                Dim color As Long
                If j = UBound(WinningCells) Then
                    ' 特別號 (紅色)
                    special = True
                    color = RGB(255, 0, 0)
                Else
                    ' 普通號 (黃色)
                    winning = winning + 1
                    color = RGB(255, 255, 0)
                End If

                ' 標示顏色
                LotteryCells(i).Interior.color = color
                WinningCells(j).Interior.color = color

                ' 跳出迴圈
                Exit For
            End If
        Next
    Next

    ' 獎項
    If winning = 6 Then
        MsgBox "頭獎"
    ElseIf winning = 5 And special = True Then
        MsgBox "貳獎"
    ElseIf winning = 5 And special = False Then
        MsgBox "參獎"
    ElseIf winning = 4 And special = True Then
        MsgBox "肆獎"
    ElseIf winning = 4 And special = False Then
        MsgBox "伍獎"
    ElseIf winning = 3 And special = True Then
        MsgBox "陸獎"
    ElseIf winning = 2 And special = True Then
        MsgBox "柒獎"
    ElseIf winning = 3 And special = False Then
        MsgBox "普獎"
    Else
        MsgBox "沒中獎"
    End If
End Sub


''''''''''''''''''''''''''
'        儲存號碼        '
''''''''''''''''''''''''''

Public Sub 儲存號碼_對獎_click()
    Debug.Print "儲存號碼_對獎"
End Sub

Public Sub 儲存號碼_檢查_click()
    Dim result As Boolean
    result = True

    Dim i As Integer
    For i = 1 To Counter.value

        ' 設定 ArrayCells
        Dim ArrayCells(1 To 6) As Object
        For j = LBound(ArrayCells) To UBound(ArrayCells)
            Set ArrayCells(j) = cells(Store_Row - 1 + i, Store_Col + j)
        Next

        ' 檢查 ArrayCells，以粉紅色標示不合格的標號
        If Check_Lottery_Cells(ArrayCells) <> "" Then
            result = False
            For Each cell In ArrayCells
                cell.Interior.color = RGB(255, 217, 236)
            Next
        End If
    Next

    If result = True Then
        MsgBox "儲存號碼 全部檢查正確"
    End If
End Sub

Public Sub 儲存號碼_清除_click()
    If Counter.value = 0 Then
        Exit Sub
    End If

    ' 清除儲存格
    Reset_Color
    Range(cells(Store_Row, Store_Col), cells(Counter.value + Store_Row - 1, Store_Col + 7)) = ""

    ' 計數器歸零
    Counter.value = 0
End Sub


''''''''''''''''''''''''''
'        Function        '
''''''''''''''''''''''''''

Private Function Generate_Lottery_Cells(ByRef ArrayCells() As Object)
    ' 產生 nums 陣列 = {1, 2, ..., 49}
    Dim i, nums(1 To 49) As Integer
    For i = LBound(nums) To UBound(nums)
        nums(i) = i
    Next

    ' 隨機產生六個大樂透號碼，寫到 ArrayCells
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        ' 隨機產生一個 i 到 49 之間的亂數 n
        Dim n As Integer
        n = Int(Rnd * (49 - i + 1) + i)

        ' 將 nums 的第 i 項的值與第 n 項交換
        Dim tmp As Integer
        tmp = nums(i)
        nums(i) = nums(n)
        nums(n) = tmp

        ' 寫入 ArrayCells
        ArrayCells(i).value = nums(i)
    Next
End Function

Private Function Check_Lottery_Cells(ByRef ArrayCells() As Object) As String
    Check_Lottery_Cells = ""

    Dim i As Integer
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        Dim value
        value = ArrayCells(i).value

        Dim name As String
        name = IIf(i = 7, "特別碼", "第 " & i & " 碼")

        ' 1. 檢查空白欄位
        If IsEmpty(value) Then
            Check_Lottery_Cells = name & "不可空值"
            Exit Function
        End If

        ' 2. 檢查型別
        If Not IsNumeric(value) Then
            Check_Lottery_Cells = name & " ( " & value & " ) 必須為整數型別"
            Exit Function
        End If

        ' 3. 檢查整數
        If Round(value) <> value Then
            Check_Lottery_Cells = name & " ( " & value & " ) 必須為整數型別"
            Exit Function
        End If

        ' 4. 超出範圍 1 ~ 49
        If 1 > value Or value > 49 Then
            Check_Lottery_Cells = name & " ( " & value & " ) 超出範圍 1～49"
            Exit Function
        End If

        ' 5. 檢查重複
        Dim j As Integer
        For j = i + 1 To UBound(ArrayCells)
            If value = ArrayCells(j).value Then
                Check_Lottery_Cells = name & " ( " & value & " ) 號碼重複"
                Exit Function
            End If
        Next
    Next
End Function

Function Reset_Color()
    Range(cells(Store_Row, Store_Col), cells(Counter.value + Store_Row - 1, Store_Col + 7)).Interior.ColorIndex = xlNone
End Function
