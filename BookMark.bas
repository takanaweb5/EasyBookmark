Attribute VB_Name = "BookMark"
Option Explicit

Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer

Const C_Color = &HFFFFCC
Const C_PatternColorIndex = 29
Private lngColor  As Long
    
'*****************************************************************************
'[概要] 選択されセルにBookmarkを設定/解除する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetBookmark()
On Error GoTo ErrHandle
    Dim objRange As Range
    
    'Rangeオブジェクトが選択されているか判定
    If TypeOf Selection Is Range Then
        Set objRange = Selection
    Else
        Exit Sub
    End If
        
    With objRange.Cells(1).Interior
        If .Pattern = xlSolid And _
           .PatternColorIndex = C_PatternColorIndex Then
            '書式のクリア
            objRange.Interior.ColorIndex = xlNone
            Exit Sub
        End If
    End With
    
    With objRange.Interior
        '[Ctrl]Keyが押下されていれば、セルの色を選択
        If GetKeyState(vbKeyControl) < 0 Then
            '20:カラーパターンの水色(デフォルト色)
            If Application.Dialogs(xlDialogPatterns).Show(, , 20) = False Then
                Exit Sub
            End If
            If .ColorIndex = xlColorIndexNone Then
                lngColor = C_Color
            Else
                lngColor = .Color
            End If
        Else
            If lngColor = 0 Then
                lngColor = C_Color
            End If
        End If
        .Pattern = xlSolid
        .Color = lngColor
        .PatternColorIndex = C_PatternColorIndex
    End With
ErrHandle:
End Sub

'*****************************************************************************
'[概要] 次のBookmarkに移動
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub NextOrPrevBookmark()
    If CommandBars.ActionControl.Caption = "次のブックマーク" Then
        '[Shift]または[Ctrl]Keyが押下されていれば、前方に移動
        If GetKeyState(vbKeyShift) < 0 Or GetKeyState(vbKeyControl) < 0 Then
            Call JumpBookmark(xlPrevious)
        Else
            Call JumpBookmark(xlNext)
        End If
    Else
        '[Shift]または[Ctrl]Keyが押下されていれば、後方に移動
        If GetKeyState(vbKeyShift) < 0 Or GetKeyState(vbKeyControl) < 0 Then
            Call JumpBookmark(xlNext)
        Else
            Call JumpBookmark(xlPrevious)
        End If
    End If
End Sub

'*****************************************************************************
'[概要] 次のBookmarkに移動
'[引数] 検索方向
'[戻値] なし
'*****************************************************************************
Private Sub JumpBookmark(ByVal xlDirection As XlSearchDirection)
On Error GoTo ErrHandle
    Dim objCell      As Range
    Dim objNextCell  As Range
    Dim objSheetCell As Range
    
    Call SetFindFormat
    
    '****************************************
    'アクティブシート内の検索
    '****************************************
    Dim blnFind  As Boolean
    Set objCell = ActiveCell
    Set objNextCell = FindNextFormat(objCell, xlDirection)
    If Not (objNextCell Is Nothing) Then
        Set objSheetCell = objNextCell
        If TypeOf Selection Is Range Then
            If xlDirection = xlNext Then
                If objNextCell.Row > objCell.Row Or _
                  (objNextCell.Row = objCell.Row And objNextCell.Column > objCell.Column) Then
                    blnFind = True
                End If
            Else
                If objNextCell.Row < objCell.Row Or _
                  (objNextCell.Row = objCell.Row And objNextCell.Column < objCell.Column) Then
                    blnFind = True
                End If
            End If
        Else
            blnFind = True
        End If
    End If
    
    If blnFind = True Then
        Call objNextCell.Select
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '****************************************
    '隣のシートの検索
    '****************************************
    Dim i As Long
    Dim j As Long
    Dim lngSheetCnt As Long
    Dim lngStartIdx As Long
    
    lngSheetCnt = ActiveWorkbook.Worksheets.Count
    j = ActiveSheet.Index
    
    For i = 2 To lngSheetCnt
        If xlDirection = xlNext Then
            j = j + 1
            If j > lngSheetCnt Then
                j = 1
            End If
            Set objCell = ActiveWorkbook.Worksheets(j).Cells(Rows.Count, Columns.Count)
        Else
            j = j - 1
            If j < 1 Then
                j = lngSheetCnt
            End If
            Set objCell = ActiveWorkbook.Worksheets(j).Cells(1, 1)
        End If
        
        Set objNextCell = FindNextFormat(objCell, xlDirection)
        If Not (objNextCell Is Nothing) Then
            Call objNextCell.Worksheet.Select
            Call objNextCell.Select
            Application.FindFormat.Clear
            Exit Sub
        End If
    Next

    If Not (objSheetCell Is Nothing) Then
        Call objSheetCell.Select
    End If
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] Bookmark検索用のセル書式を設定する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SetFindFormat()
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
    If TypeOf Selection Is Range Then
        '選択されているセルが1つだけか判定
        If Not IsOnlyCell(Selection) Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    With ActiveCell.Interior
        If .Pattern = xlSolid And _
           .PatternColorIndex = C_PatternColorIndex Then
            Application.FindFormat.Interior.Color = .Color
        End If
    End With
End Sub

'*****************************************************************************
'[概要] Rangeが(結合された)単一のセルかどうか
'[引数] 判定するRange
'[戻値] True:単一のセル、False:複数のセル
'*****************************************************************************
Private Function IsOnlyCell(ByRef objRange As Range) As Boolean
    IsOnlyCell = (objRange.Address(0, 0) = objRange(1, 1).MergeArea.Address(0, 0))
End Function

'*****************************************************************************
'[概要] すべてのBookmarkを選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub SelectAllBookmarks()
On Error GoTo ErrHandle
    Dim objRange As Range
    
    Call SetFindFormat
    Set objRange = GetBookmarks(ActiveWorkbook.ActiveSheet)
    If Not (objRange Is Nothing) Then
        Call objRange.Select
    End If
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] すべてのBookmarkをクリア
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ClearBookmarks()
On Error GoTo ErrHandle
    Dim i As Long
    Dim j1 As Long
    Dim j2 As Long
    Dim objRange As Range
    
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
    'すべてのブックマークの数を計算
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objRange = GetBookmarks(ActiveWorkbook.Worksheets(i))
        If Not (objRange Is Nothing) Then
            j1 = j1 + objRange.Cells.Count
        End If
    Next
    
    If j1 = 0 Then
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '選択セルと同色のブックマークの数を計算
    Call SetFindFormat
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objRange = GetBookmarks(ActiveWorkbook.Worksheets(i))
        If Not (objRange Is Nothing) Then
            j2 = j2 + objRange.Cells.Count
        End If
    Next
    
    '****************************************
    '実行確認
    '****************************************
    If j1 = j2 Then
        If MsgBox(j1 & " 個のブックマークを削除します" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    Else
        If MsgBox(j1 & " 個のブックマークのうち" & vbLf & "選択されたセルと同じ色の " & j2 & " 個のブックマークを削除します" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    End If
    
    '****************************************
    'すべてのブックマークを削除
    '****************************************
    Application.ScreenUpdating = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objRange = GetBookmarks(ActiveWorkbook.Worksheets(i), True)
        If Not (objRange Is Nothing) Then
            objRange.Interior.ColorIndex = xlNone
        End If
    Next
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] 対象シートのすべてのBookmarkを取得
'[引数] 対象シート、blnMerge:マージセルの時マージエリアを取得する
'[戻値] Bookmarkが設定されたセルすべて
'*****************************************************************************
Private Function GetBookmarks(ByRef objSheet As Worksheet, Optional blnMerge As Boolean = False) As Range
    Dim objCell  As Range
    Dim objRange As Range
    
    Set objCell = GetLastCell(objSheet)
    Do While (True)
        Set objCell = FindNextFormat(objCell, xlNext)
        If objCell Is Nothing Then
            Exit Function
        End If
        
        If blnMerge Then
            Set objRange = objCell.MergeArea
        Else
            Set objRange = objCell
        End If
        
        If GetBookmarks Is Nothing Then
            Set GetBookmarks = objRange
        ElseIf Intersect(GetBookmarks, objRange) Is Nothing Then
            Set GetBookmarks = Application.Union(GetBookmarks, objRange)
        Else
            Exit Function
        End If
    Loop
End Function

'*****************************************************************************
'[概要] 次の書式のセルに移動
'[引数] 検索開始セル、検索方向
'[戻値] 次の書式のセル
'*****************************************************************************
Private Function FindNextFormat(ByRef objCell As Range, _
                                ByVal xlDirection As XlSearchDirection) As Range
    Dim objUsedRange As Range
    With objCell.Worksheet
        Set objUsedRange = .Range(.Range("A1"), .Cells.SpecialCells(xlLastCell))
        Set objUsedRange = Application.Union(objUsedRange, objCell)
    End With
    Set FindNextFormat = objUsedRange.Find("", objCell, _
                  xlFormulas, xlPart, xlByRows, xlDirection, False, False, True)
End Function

'*****************************************************************************
'[概要] 次を検索
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub FindNext()
On Error GoTo ErrHandle
    Dim objCell As Range
    Dim xlDirection As XlSearchDirection
    
    '[Shift]または[Ctrl]Keyが押下されていれば、前方検索
    If GetKeyState(vbKeyShift) < 0 Or GetKeyState(vbKeyControl) < 0 Then
        xlDirection = xlPrevious
    Else
        xlDirection = xlNext
    End If
    
    Set objCell = FindJump(ActiveCell, xlDirection)
    
    If Not (objCell Is Nothing) Then
        Call objCell.Select
'    Else
'        Call ShowFindDialog
    End If

ErrHandle:
    Call ActiveCell.Worksheet.Select
End Sub

'*****************************************************************************
'[概要] 次を検索
'[引数] 検索開始セル、検索方向
'[戻値] 次のセル
'*****************************************************************************
Private Function FindJump(ByRef objNowCell As Range, ByVal xlDirection As XlSearchDirection) As Range
On Error GoTo ErrHandle
    Dim objCell As Range
        
    If xlDirection = xlNext Then
        Set objCell = Cells.FindNext(objNowCell)
    Else
        Set objCell = Cells.FindPrevious(objNowCell)
    End If
    If Not (objCell Is Nothing) Then
        '自分自身のセルを選択する意味不明なバグ対応
        If objCell.Address = objNowCell.Address Then
            If xlDirection = xlNext Then
                Set objCell = Cells.FindNext(objNowCell)
            Else
                Set objCell = Cells.FindPrevious(objNowCell)
            End If
        End If
        If objCell.Value <> "" Then
            Set FindJump = objCell
        End If
    End If
ErrHandle:
End Function

'*****************************************************************************
'[概要] 「検索と置換」のダイアログを表示する
'[引数] なし
'[戻値] なし
'*****************************************************************************
Private Sub ShowFindDialog()
    Dim objTmpBar As CommandBar
    
    Set objTmpBar = CommandBars.Add(, msoBarPopup, , True)
    With objTmpBar.Controls.Add(, 1849)
        Call ActiveCell.Select  'Shapeが選択されていると例外になるため
        .Execute
    End With
    Call objTmpBar.Delete
End Sub

'*****************************************************************************
'[概要] 使用されている最後のセルを取得する
'[引数] 対象のシート
'[戻値] 最後のセル
'*****************************************************************************
Private Function GetLastCell(ByRef objSheet As Worksheet) As Range
    Set GetLastCell = objSheet.Cells.SpecialCells(xlLastCell)
End Function
