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
    Debug.Print ("自動選號")
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
