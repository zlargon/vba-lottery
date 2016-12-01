Private LotteryCells(1 To 6) As Object
Private WinningCells(1 To 7) As Object
Private Counter As Object

''''''''''''''''''''''''''
'       Auto_Open        '
''''''''''''''''''''''''''

Private Sub Auto_Open()
    Set LotteryCells(1) = cells(3, 2)
    Set LotteryCells(2) = cells(3, 3)
    Set LotteryCells(3) = cells(3, 4)
    Set LotteryCells(4) = cells(3, 5)
    Set LotteryCells(5) = cells(3, 6)
    Set LotteryCells(6) = cells(3, 7)

    Set WinningCells(1) = cells(11, 2)
    Set WinningCells(2) = cells(11, 3)
    Set WinningCells(3) = cells(11, 4)
    Set WinningCells(4) = cells(11, 5)
    Set WinningCells(5) = cells(11, 6)
    Set WinningCells(6) = cells(11, 7)
    Set WinningCells(7) = cells(11, 8)

    Set Counter = cells(17, 1)
End Sub


''''''''''''''''''''''''''
'        投注號碼        '
''''''''''''''''''''''''''

Public Sub 投注號碼_產生_click()
    generateLotteryCells ArrayCells:=LotteryCells
End Sub

Public Sub 投注號碼_檢查_click()
    If Check_Lottery_Cells(LotteryCells) = False Then
        Exit Sub
    End If

    MsgBox "投注號碼 檢查正確"
End Sub

Public Sub 投注號碼_清除_click()
    For Each cell In LotteryCells
        cell.value = ""
    Next
End Sub


''''''''''''''''''''''''''
'        開獎號碼        '
''''''''''''''''''''''''''

Public Sub 開獎號碼_產生_click()
    generateLotteryCells ArrayCells:=WinningCells
End Sub

Public Sub 開獎號碼_檢查_click()
    If Check_Lottery_Cells(WinningCells) = False Then
        Exit Sub
    End If

    MsgBox "開獎號碼 檢查正確"
End Sub

Public Sub 開獎號碼_清除_click()
    For Each cell In WinningCells
        cell.value = ""
    Next
End Sub

Public Sub 開獎號碼_儲存_click()
    Debug.Print "開獎號碼_儲存"
End Sub


''''''''''''''''''''''''''
'          對獎          '
''''''''''''''''''''''''''

Public Sub 對獎_click()
    If Check_Lottery_Cells(WinningCells) = False Then
        Exit Sub
    End If

    Debug.Print "對獎"
End Sub


''''''''''''''''''''''''''
'        儲存號碼        '
''''''''''''''''''''''''''

Public Sub 儲存號碼_清除_click()
    Debug.Print "儲存號碼_清除"

    Counter.value = 0
End Sub


''''''''''''''''''''''''''
'        Function        '
''''''''''''''''''''''''''

Private Function generateLotteryCells(ByRef ArrayCells() As Object)
    ' 產生 nums 陣列 = {1, 2, ..., 49}
    Dim i, nums(1 To 49) As Integer
    For i = 1 To 49
        nums(i) = i
    Next

    ' 隨機產生六個大樂透號碼，寫到 ArrayCells
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        ' 隨機產生一個 i 到 49 之間的亂數 n
        Dim n As Integer
        n = Int(Rnd * (49 - i + 1)) + i

        ' 將 nums 的第 i 項的值與第 n 項交換
        Dim tmp As Integer
        tmp = nums(i)
        nums(i) = nums(n)
        nums(n) = tmp

        ' 寫入 ArrayCells
        ArrayCells(i).value = nums(i)
    Next
End Function

Private Function Check_Lottery_Cells(ByRef ArrayCells() As Object) As Boolean
    Check_Lottery_Cells = False

    Dim i As Integer
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        Dim value
        value = ArrayCells(i).value

        Dim name As String
        name = IIf(i = 7, "特別碼", "第 " & i & " 碼")

        ' 1. 檢查空白欄位
        If IsEmpty(value) Then
            MsgBox name & "不可空值"
            Exit Function
        End If

        ' 2. 檢查型別
        If Not IsNumeric(value) Then
            MsgBox name & " ( " & value & " ) 必須為整數型別"
            Exit Function
        End If

        ' 3. 檢查整數
        If Round(value) <> value Then
            MsgBox name & " ( " & value & " ) 必須為整數型別"
            Exit Function
        End If

        ' 4. 超出範圍 1 ~ 49
        If 1 > value Or value > 49 Then
            MsgBox name & " ( " & value & " ) 超出範圍 1～49"
            Exit Function
        End If

        ' 5. 檢查重複
        Dim j As Integer
        For j = i + 1 To UBound(ArrayCells)
            If value = ArrayCells(j).value Then
                MsgBox name & " ( " & value & " ) 號碼重複"
                Exit Function
            End If
        Next
    Next

    Check_Lottery_Cells = True
End Function

Private Function checkWinningPrize(ByVal no As Integer)
    Dim winning As Integer
    Dim special As Boolean
    winning = 0
    special = False

    ' 計算中獎數目和特別號
    Dim i, j As Integer
    For i = 1 To 6
        Dim value As Integer
        value = cells(no + 1, i + 1).value

        For j = LBound(WinningCells) To UBound(WinningCells)
            If WinningCells(j).value = value Then
                ' special number
                If j = UBound(WinningCells) Then
                    special = True
                    Exit For
                End If

                ' winning number
                winning = winning + 1
                Exit For
            End If
        Next
    Next

    ' 獎項
    Dim title As String
    title = "--"
    If winning = 6 Then
        title = "頭獎"
    ElseIf winning = 5 And special = True Then
        title = "貳獎"
    ElseIf winning = 5 And special = False Then
        title = "參獎"
    ElseIf winning = 4 And special = True Then
        title = "肆獎"
    ElseIf winning = 4 And special = False Then
        title = "伍獎"
    ElseIf winning = 3 And special = True Then
        title = "陸獎"
    ElseIf winning = 2 And special = True Then
        title = "柒獎"
    ElseIf winning = 3 And special = False Then
        title = "柒獎"
    End If
    cells(no + 1, "I").value = title
End Function
