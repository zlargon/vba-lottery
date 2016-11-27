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
    If 檢查號碼 = False Then
        Exit Sub
    End If

    Debug.Print ("儲存號碼")
End Sub

Public Sub 清除選號_click()
    init
    For Each cell In LotteryCells
        cell.value = ""
    Next
End Sub

Private Function 檢查號碼() As Boolean
    init
    Dim title As String
    title = "檢查號碼"
    檢查號碼 = False

    Dim i As Integer
    For i = 1 To 6
        Dim value
        value = LotteryCells(i)

        ' 1. 檢查空白欄位
        If IsEmpty(value) Then
            MsgBox title:=title, prompt:="第 " & i & " 碼為空值"
            Exit Function
        End If

        ' 2. 檢查型別
        If Not IsNumeric(value) Then
            MsgBox title:=title, prompt:="第 " & i & " 碼 ( " & value & " ) 必須為整數型別"
            Exit Function
        End If

        ' 3. 檢查整數
        If Round(value) <> value Then
            MsgBox title:=title, prompt:="第 " & i & " 碼 ( " & value & " ) 必須為整數型別"
            Exit Function
        End If

        ' 4. 超出範圍 1 ~ 49
        If 1 > value Or value > 49 Then
            MsgBox title:=title, prompt:="第 " & i & " 碼 ( " & value & " ) 超出範圍 1～49"
            Exit Function
        End If

        ' 5. 檢查重複
        Dim j As Integer
        For j = i + 1 To 6
            If value = LotteryCells(j).value Then
                MsgBox title:=title, prompt:="第 " & i & "、" & j & " 碼 ( " & value & " ) 號碼重複"
                Exit Function
            End If
        Next
    Next

    檢查號碼 = True
End Function

