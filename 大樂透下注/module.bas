' Module_大樂透下注.bas

Private LotteryCells(1 To 6) As Object

Private Sub init()
    Static isInit As Boolean
    If isInit = True Then
        Exit Sub
    End If

    Set LotteryCells(1) = Cells(2, 2)
    Set LotteryCells(2) = Cells(2, 3)
    Set LotteryCells(3) = Cells(2, 4)
    Set LotteryCells(4) = Cells(2, 5)
    Set LotteryCells(5) = Cells(2, 6)
    Set LotteryCells(6) = Cells(2, 7)

    isInit = True
End Sub

Public Sub 自動選號_click()
    init
    ' 產生 nums 陣列 = {1, 2, ..., 49}
    Dim i, nums(1 To 49) As Integer
    For i = 1 To 49
        nums(i) = i
    Next

    ' 隨機產生六個大樂透號碼，寫到 LotteryCells(1 to 6)
    For i = 1 To 6
        ' 隨機產生一個 i 到 49 之間的亂數 n
        Dim n As Integer
        n = Int(Rnd * (49 - i + 1)) + i

        ' 將 nums 的第 i 項的值與第 n 項交換
        Dim tmp As Integer
        tmp = nums(i)
        nums(i) = nums(n)
        nums(n) = tmp

        ' 寫入 LotteryCells
        LotteryCells(i).value = nums(i)
    Next
End Sub

Public Sub 儲存號碼_click()
    init
    Debug.Print ("儲存號碼")
End Sub

Public Sub 清除選號_click()
    init
    Debug.Print ("清除號碼")
End Sub

Private Function 檢查號碼() As Boolean
    init
    檢查號碼 = False
End Function
