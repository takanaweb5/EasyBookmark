Attribute VB_Name = "BookMark"
Option Explicit

Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer

Const C_Color = &HFFFFCC
Const C_PatternColorIndex = 29
Private lngColor  As Long
    
'*****************************************************************************
'[�T�v] �I������Z����Bookmark��ݒ�/��������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetBookmark()
On Error GoTo ErrHandle
    Dim objRange As Range
    
    'Range�I�u�W�F�N�g���I������Ă��邩����
    If TypeOf Selection Is Range Then
        Set objRange = Selection
    Else
        Exit Sub
    End If
        
    With objRange.Cells(1).Interior
        If .Pattern = xlSolid And _
           .PatternColorIndex = C_PatternColorIndex Then
            '�����̃N���A
            objRange.Interior.ColorIndex = xlNone
            Exit Sub
        End If
    End With
    
    With objRange.Interior
        '[Ctrl]Key����������Ă���΁A�Z���̐F��I��
        If GetKeyState(vbKeyControl) < 0 Then
            '20:�J���[�p�^�[���̐��F(�f�t�H���g�F)
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
'[�T�v] ����Bookmark�Ɉړ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub NextOrPrevBookmark()
    If CommandBars.ActionControl.Caption = "���̃u�b�N�}�[�N" Then
        '[Shift]�܂���[Ctrl]Key����������Ă���΁A�O���Ɉړ�
        If GetKeyState(vbKeyShift) < 0 Or GetKeyState(vbKeyControl) < 0 Then
            Call JumpBookmark(xlPrevious)
        Else
            Call JumpBookmark(xlNext)
        End If
    Else
        '[Shift]�܂���[Ctrl]Key����������Ă���΁A����Ɉړ�
        If GetKeyState(vbKeyShift) < 0 Or GetKeyState(vbKeyControl) < 0 Then
            Call JumpBookmark(xlNext)
        Else
            Call JumpBookmark(xlPrevious)
        End If
    End If
End Sub

'*****************************************************************************
'[�T�v] ����Bookmark�Ɉړ�
'[����] ��������
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub JumpBookmark(ByVal xlDirection As XlSearchDirection)
On Error GoTo ErrHandle
    Dim objCell      As Range
    Dim objNextCell  As Range
    Dim objSheetCell As Range
    
    Call SetFindFormat
    
    '****************************************
    '�A�N�e�B�u�V�[�g���̌���
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
    '�ׂ̃V�[�g�̌���
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
'[�T�v] Bookmark�����p�̃Z��������ݒ肷��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub SetFindFormat()
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
    If TypeOf Selection Is Range Then
        '�I������Ă���Z����1����������
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
'[�T�v] Range��(�������ꂽ)�P��̃Z�����ǂ���
'[����] ���肷��Range
'[�ߒl] True:�P��̃Z���AFalse:�����̃Z��
'*****************************************************************************
Private Function IsOnlyCell(ByRef objRange As Range) As Boolean
    IsOnlyCell = (objRange.Address(0, 0) = objRange(1, 1).MergeArea.Address(0, 0))
End Function

'*****************************************************************************
'[�T�v] ���ׂĂ�Bookmark��I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
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
'[�T�v] ���ׂĂ�Bookmark���N���A
'[����] �Ȃ�
'[�ߒl] �Ȃ�
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
    
    '���ׂẴu�b�N�}�[�N�̐����v�Z
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
    
    '�I���Z���Ɠ��F�̃u�b�N�}�[�N�̐����v�Z
    Call SetFindFormat
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objRange = GetBookmarks(ActiveWorkbook.Worksheets(i))
        If Not (objRange Is Nothing) Then
            j2 = j2 + objRange.Cells.Count
        End If
    Next
    
    '****************************************
    '���s�m�F
    '****************************************
    If j1 = j2 Then
        If MsgBox(j1 & " �̃u�b�N�}�[�N���폜���܂�" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    Else
        If MsgBox(j1 & " �̃u�b�N�}�[�N�̂���" & vbLf & "�I�����ꂽ�Z���Ɠ����F�� " & j2 & " �̃u�b�N�}�[�N���폜���܂�" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
            Application.FindFormat.Clear
            Exit Sub
        End If
    End If
    
    '****************************************
    '���ׂẴu�b�N�}�[�N���폜
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
'[�T�v] �ΏۃV�[�g�̂��ׂĂ�Bookmark���擾
'[����] �ΏۃV�[�g�AblnMerge:�}�[�W�Z���̎��}�[�W�G���A���擾����
'[�ߒl] Bookmark���ݒ肳�ꂽ�Z�����ׂ�
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
'[�T�v] ���̏����̃Z���Ɉړ�
'[����] �����J�n�Z���A��������
'[�ߒl] ���̏����̃Z��
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
'[�T�v] ��������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub FindNext()
On Error GoTo ErrHandle
    Dim objCell As Range
    Dim xlDirection As XlSearchDirection
    
    '[Shift]�܂���[Ctrl]Key����������Ă���΁A�O������
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
'[�T�v] ��������
'[����] �����J�n�Z���A��������
'[�ߒl] ���̃Z��
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
        '�������g�̃Z����I������Ӗ��s���ȃo�O�Ή�
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
'[�T�v] �u�����ƒu���v�̃_�C�A���O��\������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Private Sub ShowFindDialog()
    Dim objTmpBar As CommandBar
    
    Set objTmpBar = CommandBars.Add(, msoBarPopup, , True)
    With objTmpBar.Controls.Add(, 1849)
        Call ActiveCell.Select  'Shape���I������Ă���Ɨ�O�ɂȂ邽��
        .Execute
    End With
    Call objTmpBar.Delete
End Sub

'*****************************************************************************
'[�T�v] �g�p����Ă���Ō�̃Z�����擾����
'[����] �Ώۂ̃V�[�g
'[�ߒl] �Ō�̃Z��
'*****************************************************************************
Private Function GetLastCell(ByRef objSheet As Worksheet) As Range
    Set GetLastCell = objSheet.Cells.SpecialCells(xlLastCell)
End Function
