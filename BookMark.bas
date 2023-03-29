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
Public Sub SetBookmark()
On Error GoTo ErrHandle
    Dim objRange As Range
    Dim blnSet   As Boolean
    
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
            If Application.Dialogs(xlDialogPatterns).Show = False Then
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
Public Sub NextOrPrevBookmark()
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
    
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
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
    Next i

    If Not (objSheetCell Is Nothing) Then
        Call objSheetCell.Select
    End If
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] すべてのBookmarkを選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub SelectAllBookmarks()
On Error GoTo ErrHandle
    Dim objCell  As Range
    Dim objRange As Range
    
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
    'アクティブシート上のすべてのBookmarkを取得
    Set objCell = ActiveSheet.Cells(Rows.Count, Columns.Count)
    Do While (True)
        Set objCell = FindNextFormat(objCell, xlNext)
        If objCell Is Nothing Then
            Exit Do
        ElseIf objRange Is Nothing Then
            Set objRange = objCell
        ElseIf Intersect(objRange, objCell) Is Nothing Then
            Set objRange = Union(objRange, objCell)
        Else
            'すべてのBookmarkを選択
            Call objRange.Select
            Exit Do
        End If
    Loop
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] すべてのBookmarkをクリア
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub ClearBookmarks()
On Error GoTo ErrHandle
    Dim i As Long
    Dim j As Long
    Dim objCell  As Range
    Dim objRange As Range
    
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
    '****************************************
    'ブックマークの数を計算
    '****************************************
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objRange = Nothing
        Set objCell = ActiveWorkbook.Worksheets(i).Cells(1, 1)
        Do While (True)
            Set objCell = FindNextFormat(objCell, xlNext)
            If objCell Is Nothing Then
                Exit Do
            ElseIf objRange Is Nothing Then
                Set objRange = objCell
                j = j + 1
            ElseIf Intersect(objRange, objCell) Is Nothing Then
                Set objRange = Union(objRange, objCell)
                j = j + 1
            Else
                Exit Do
            End If
        Loop
    Next i
    
    If j = 0 Then
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '****************************************
    '実行確認
    '****************************************
    If MsgBox(j & "個のブックマークを削除します" & vbLf & "よろしいですか？", vbOKCancel + vbQuestion) = vbCancel Then
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '****************************************
    'すべてのブックマークを削除
    '****************************************
    Application.ScreenUpdating = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objCell = ActiveWorkbook.Worksheets(i).Cells(1, 1)
        Do While (True)
            Set objCell = FindNextFormat(objCell, xlNext)
            If objCell Is Nothing Then
                Exit Do
            Else
                '書式のクリア
                With objCell.Interior
                    If .Color = lngColor Then
                        .ColorIndex = xlNone
                    Else
                        .PatternColorIndex = xlAutomatic
                    End If
                End With
            End If
        Loop
    Next i
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[概要] 次の書式のセルに移動
'[引数] 検索開始セル
'            検索方向
'            検索文字列（省略可）
'[戻値] 次の書式のセル
'*****************************************************************************
Private Function FindNextFormat(ByRef objNowCell As Range, _
                                ByVal xlDirection As XlSearchDirection, _
                       Optional ByVal strFind As String = "") As Range
    Set FindNextFormat = objNowCell.Worksheet.Cells.Find(strFind, objNowCell, _
                  xlFormulas, xlPart, xlByRows, xlDirection, False, False, True)
End Function

'*****************************************************************************
'[概要] 次を検索
'[引数] なし
'[戻値] なし
'*****************************************************************************
Public Sub FindNext()
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
'[概要] すべての検索対象セルを選択
'[引数] なし
'[戻値] なし
'*****************************************************************************
'Public Sub SelectAllSearchCell()
'On Error GoTo ErrHandle
'    Dim blnFindFormat As Boolean
'    Dim blnFindstr    As Boolean
'    Dim strFind    As String
'    Dim hWnd       As Long
'    Dim objCell    As Range
'    Dim objRange   As Range
'
'    blnFindFormat = CheckFindFormat()
'    blnFindstr = Not CheckFindStrIsBlank()
'
'    '検索文字列も検索書式も設定されていない時
'    If blnFindstr = False And blnFindFormat = False Then
'        '「検索と置換」のダイアログを表示し、閉じられるまでループ
'        hWnd = ShowFindDialog
'        Do While (GetDialogHandle() <> 0)
'            strFind = GetFindStr(hWnd)
'            DoEvents
'        Loop
'        blnFindFormat = CheckFindFormat()
'    ElseIf blnFindstr = True And blnFindFormat = True Then
'        strFind = GetFindStr()
'    End If
'
'    'アクティブシート上のすべての検索対象セルを取得
'    Set objCell = ActiveSheet.Cells(Rows.Count, Columns.Count)
'    Do While (True)
'        If blnFindFormat = True Then
'            '検索書式設定あり
'            Set objCell = FindNextFormat(objCell, xlNext, strFind)
'        Else
'            '検索書式設定なし
'            Set objCell = FindJump(objCell, xlNext)
'        End If
'
'        If objCell Is Nothing Then
'            Exit Do
'        ElseIf objRange Is Nothing Then
'            Set objRange = objCell
'        ElseIf Intersect(objRange, objCell) Is Nothing Then
'            Set objRange = Union(objRange, objCell)
'        Else
'            'すべての検索対象セルを選択
'            Call objRange.Select
'            Exit Do
'        End If
'    Loop
'ErrHandle:
'    Call ActiveCell.Worksheet.Select
'End Sub

'*****************************************************************************
'[概要] 検索対象文字列が空白かどうか
'[引数] なし
'[戻値] True:検索対象文字列設定なし
'*****************************************************************************
Private Function CheckFindStrIsBlank() As Boolean
    Dim objCell As Range
    Set objCell = Cells.FindNext(ActiveCell)
    If objCell Is Nothing Then
        CheckFindStrIsBlank = True
    Else
        If objCell.Value = "" Then
            CheckFindStrIsBlank = True
        Else
            CheckFindStrIsBlank = False
        End If
    End If
End Function

'*****************************************************************************
'[概要] 次を検索
'[引数] 検索開始セル
'            検索方向
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

