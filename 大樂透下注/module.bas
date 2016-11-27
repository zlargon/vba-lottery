' Module.bas

Private LotteryCells(1 To 6) As Object
Private WinningCells(1 To 7) As Object
Private Counter As Object

Private Sub init()
    Static isInit As Boolean
    If isInit = True Then
        Exit Sub
    End If

    Set LotteryCells(1) = cells(2, 2)
    Set LotteryCells(2) = cells(2, 3)
    Set LotteryCells(3) = cells(2, 4)
    Set LotteryCells(4) = cells(2, 5)
    Set LotteryCells(5) = cells(2, 6)
    Set LotteryCells(6) = cells(2, 7)

    Set WinningCells(1) = cells(12, 2)
    Set WinningCells(2) = cells(12, 3)
    Set WinningCells(3) = cells(12, 4)
    Set WinningCells(4) = cells(12, 5)
    Set WinningCells(5) = cells(12, 6)
    Set WinningCells(6) = cells(12, 7)
    Set WinningCells(7) = cells(12, 8)

    Set Counter = cells(8, 2)

    isInit = True
End Sub

Public Sub 自動選號_click()
    init
    generateLotteryCells ArrayCells:=LotteryCells
End Sub

Public Sub 儲存號碼_click()
    init
    If checkLotteryCells(LotteryCells) = False Then
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

Public Sub 產生中獎號碼_click()
    init
    generateLotteryCells ArrayCells:=WinningCells
End Sub

Public Sub 開始兌獎_click()
    init
    If checkLotteryCells(WinningCells) = False Then
        Exit Sub
    End If

    Debug.Print ("開始兌獎")
End Sub

Public Sub 重置中獎號碼_click()
    init
    For Each cell In WinningCells
        cell.value = ""
    Next

    ' TODO: 清除中獎獎項、金額
End Sub

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

Private Function checkLotteryCells(ByRef ArrayCells() As Object) As Boolean
    checkLotteryCells = False

    Dim title As String
    title = "檢查號碼"

    Dim i As Integer
    For i = LBound(ArrayCells) To UBound(ArrayCells)
        Dim value
        value = ArrayCells(i)

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
        For j = i + 1 To UBound(ArrayCells)
            If value = ArrayCells(j).value Then
                MsgBox title:=title, prompt:="第 " & i & "、" & j & " 碼 ( " & value & " ) 號碼重複"
                Exit Function
            End If
        Next
    Next

    checkLotteryCells = True
End Function

