Attribute VB_Name = "BookMark"
Option Explicit
'�L�[����������Ă��邩�ǂ������肷��
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal lngVirtKey As Long) As Integer
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal clpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Long, ByVal hwndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Const C_ColorIndex = 20
Const C_PatternColorIndex = 29

Private lngColorIndex As Long
    
'*****************************************************************************
'[�T�v] �I������Z����Bookmark��ݒ�/��������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SetBookmark()
On Error GoTo ErrHandle
    Dim objRange As Range
    Dim blnSet   As Boolean
    
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
            If Application.Dialogs(xlDialogPatterns).Show(, , C_ColorIndex) = False Then
                Exit Sub
            End If
            If .ColorIndex = xlColorIndexNone Then
                lngColorIndex = C_ColorIndex
            Else
                lngColorIndex = .ColorIndex
            End If
        Else
            If lngColorIndex = 0 Then
                lngColorIndex = C_ColorIndex
            End If
        End If
        .Pattern = xlSolid
        .ColorIndex = lngColorIndex
        .PatternColorIndex = C_PatternColorIndex
    End With
ErrHandle:
End Sub

'*****************************************************************************
'[�T�v] ����Bookmark�Ɉړ�
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub NextOrPrevBookmark()
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
    
    Application.FindFormat.Clear
    With Application.FindFormat.Interior
        .Pattern = xlSolid
        .PatternColorIndex = C_PatternColorIndex
    End With
    
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
    Next i

    If Not (objSheetCell Is Nothing) Then
        Call objSheetCell.Select
    End If
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[�T�v] ���ׂĂ�Bookmark��I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
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
    
    '�A�N�e�B�u�V�[�g��̂��ׂĂ�Bookmark���擾
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
            '���ׂĂ�Bookmark��I��
            Call objRange.Select
            Exit Do
        End If
    Loop
ErrHandle:
    Application.FindFormat.Clear
End Sub

'*****************************************************************************
'[�T�v] ���ׂĂ�Bookmark���N���A
'[����] �Ȃ�
'[�ߒl] �Ȃ�
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
    '�u�b�N�}�[�N�̐����v�Z
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
    '���s�m�F
    '****************************************
    If MsgBox(j & "�̃u�b�N�}�[�N���폜���܂�" & vbLf & "��낵���ł����H", vbOKCancel + vbQuestion) = vbCancel Then
        Application.FindFormat.Clear
        Exit Sub
    End If
    
    '****************************************
    '���ׂẴu�b�N�}�[�N���폜
    '****************************************
    Application.ScreenUpdating = False
    For i = 1 To ActiveWorkbook.Worksheets.Count
        Set objCell = ActiveWorkbook.Worksheets(i).Cells(1, 1)
        Do While (True)
            Set objCell = FindNextFormat(objCell, xlNext)
            If objCell Is Nothing Then
                Exit Do
            Else
                '�����̃N���A
                With objCell.Interior
                    If .ColorIndex = lngColorIndex Then
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
'[�T�v] ���̏����̃Z���Ɉړ�
'[����] �����J�n�Z��
'            ��������
'            ����������i�ȗ��j
'[�ߒl] ���̏����̃Z��
'*****************************************************************************
Private Function FindNextFormat(ByRef objNowCell As Range, _
                                ByVal xlDirection As XlSearchDirection, _
                       Optional ByVal strFind As String = "") As Range
    Set FindNextFormat = objNowCell.Worksheet.Cells.Find(strFind, objNowCell, _
                  xlFormulas, xlPart, xlByRows, xlDirection, False, False, True)
End Function

'*****************************************************************************
'[�T�v] ��������
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub FindNext()
On Error GoTo ErrHandle
    Dim objCell As Range
    Dim xlDirection As XlSearchDirection
    
    '[Shift]�܂���[Ctrl]Key����������Ă���΁A�O������
    If GetKeyState(vbKeyShift) < 0 Or GetKeyState(vbKeyControl) < 0 Then
        xlDirection = xlPrevious
    Else
        xlDirection = xlNext
    End If
    
    '�����Ώۂ̏������ݒ肳��Ă��邩����
    If CheckFindFormat() = True Then
        '�����Ώە����񂪐ݒ肳��Ă��邩����
        If CheckFindStrIsBlank() = True Then
            Set objCell = FindNextFormat(ActiveCell, xlDirection, "")
        Else
            Set objCell = FindNextFormat(ActiveCell, xlDirection, GetFindStr())
        End If
    Else
        Set objCell = FindJump(ActiveCell, xlDirection)
    End If
    If Not (objCell Is Nothing) Then
        Call objCell.Select
    Else
        Call ShowFindDialog
    End If

ErrHandle:
    Call ActiveCell.Worksheet.Select
End Sub

'*****************************************************************************
'[�T�v] ���ׂĂ̌����ΏۃZ����I��
'[����] �Ȃ�
'[�ߒl] �Ȃ�
'*****************************************************************************
Public Sub SelectAllSearchCell()
On Error GoTo ErrHandle
    Dim blnFindFormat As Boolean
    Dim blnFindstr    As Boolean
    Dim strFind    As String
    Dim hWnd       As Long
    Dim objCell    As Range
    Dim objRange   As Range
    
    blnFindFormat = CheckFindFormat()
    blnFindstr = Not CheckFindStrIsBlank()
    
    '����������������������ݒ肳��Ă��Ȃ���
    If blnFindstr = False And blnFindFormat = False Then
        '�u�����ƒu���v�̃_�C�A���O��\�����A������܂Ń��[�v
        hWnd = ShowFindDialog
        Do While (GetDialogHandle() <> 0)
            strFind = GetFindStr(hWnd)
            DoEvents
        Loop
        blnFindFormat = CheckFindFormat()
    ElseIf blnFindstr = True And blnFindFormat = True Then
        strFind = GetFindStr()
    End If
    
    '�A�N�e�B�u�V�[�g��̂��ׂĂ̌����ΏۃZ�����擾
    Set objCell = ActiveSheet.Cells(Rows.Count, Columns.Count)
    Do While (True)
        If blnFindFormat = True Then
            '���������ݒ肠��
            Set objCell = FindNextFormat(objCell, xlNext, strFind)
        Else
            '���������ݒ�Ȃ�
            Set objCell = FindJump(objCell, xlNext)
        End If
        
        If objCell Is Nothing Then
            Exit Do
        ElseIf objRange Is Nothing Then
            Set objRange = objCell
        ElseIf Intersect(objRange, objCell) Is Nothing Then
            Set objRange = Union(objRange, objCell)
        Else
            '���ׂĂ̌����ΏۃZ����I��
            Call objRange.Select
            Exit Do
        End If
    Loop
ErrHandle:
    Call ActiveCell.Worksheet.Select
End Sub

'*****************************************************************************
'[�T�v] �����Ώە����񂪋󔒂��ǂ���
'[����] �Ȃ�
'[�ߒl] True:�����Ώە�����ݒ�Ȃ�
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
'[�T�v] ��������
'[����] �����J�n�Z��
'            ��������
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
'[�T�v] �u�����ƒu���v�̃_�C�A���O���猟���Ώۂ̕�������擾����
'[����] �u�����ƒu���v�̃_�C�A���O�̃n���h��
'[�ߒl] ����������
'*****************************************************************************
Private Function GetFindStr(Optional ByVal hWnd As Long = 0) As String
    Dim hFindWnd  As Long
    
    If hWnd = 0 Then
        hFindWnd = ShowFindDialog()
        GetFindStr = GetFindStrFromEdit(hFindWnd)
        Call CloseDialog(hFindWnd)
    Else
        GetFindStr = GetFindStrFromEdit(hWnd)
    End If
End Function

'*****************************************************************************
'[�T�v] �u�����ƒu���v�̃_�C�A���O�����
'[����] �u�����ƒu���v�̃_�C�A���O�̃n���h��
'[�ߒl] ����������
'*****************************************************************************
Private Sub CloseDialog(ByVal hWnd As Long)
    Const WM_SYSCOMMAND = &H112
    Const SC_CLOSE = &HF060& '�E�B���h�E���I������
    Call SendMessage(hWnd, WM_SYSCOMMAND, SC_CLOSE, 0)
End Sub

'*****************************************************************************
'[�T�v] �u�����ƒu���v�̃_�C�A���O�̃n���h�����擾
'[����] �Ȃ�
'[�ߒl] �u�����ƒu���v�̃_�C�A���O�̃n���h��
'*****************************************************************************
Private Function GetDialogHandle() As Long
    GetDialogHandle = FindWindowA("bosa_sdm_XL9", "�����ƒu��")
End Function

'*****************************************************************************
'[�T�v] �u�������镶����v�̃R���g���[�����猟���Ώە�������擾
'[����] �u�����ƒu���v�̃_�C�A���O�̃n���h��
'[�ߒl] �����Ώە�����
'*****************************************************************************
Private Function GetFindStrFromEdit(ByVal hFindWnd As Long) As String
    Dim strFindStr As String
    Dim hChildWnd  As Long
    Dim lngLen     As Long
    
    hChildWnd = FindWindowEx(hFindWnd, 0, "EDTBX", "")
    strFindStr = String(1024, Chr(0))
    lngLen = GetWindowText(hChildWnd, strFindStr, 1024)
    GetFindStrFromEdit = Left$(strFindStr, lngLen)
End Function

'*****************************************************************************
'[�T�v] FindFormat���ݒ肳��Ă��邩����
'[����] �Ȃ�
'[�ߒl] True:�ݒ肠��
'*****************************************************************************
Private Function CheckFindFormat() As Boolean
    
    CheckFindFormat = True
    
    With Application.FindFormat
        With .Interior
            If .Color = 0 And _
               IsNull(.Pattern) Then
            Else
                Exit Function
            End If
        End With
        
        With .Font
            If IsNull(.Name) And _
               IsNull(.Size) And _
               IsNull(.FontStyle) And _
               IsNull(.Background) And _
               IsNull(.Color) And _
               IsNull(.Bold) And _
               IsNull(.Italic) And _
               IsNull(.Strikethrough) And _
               IsNull(.Subscript) And _
               IsNull(.Superscript) And _
               IsNull(.Underline) Then
            Else
                Exit Function
            End If
        End With
        
        With .Borders
            If IsNull(.Item(1).LineStyle) And _
               IsNull(.Item(2).LineStyle) And _
               IsNull(.Item(3).LineStyle) And _
               IsNull(.Item(4).LineStyle) And _
               IsNull(.Item(5).LineStyle) And _
               IsNull(.Item(6).LineStyle) Then
            Else
                Exit Function
            End If
        End With
        
        If IsNull(.AddIndent) And _
           IsNull(.FormulaHidden) And _
           IsNull(.HorizontalAlignment) And _
           IsNull(.IndentLevel) And _
           IsNull(.Locked) And _
           IsNull(.MergeCells) And _
           IsNull(.NumberFormat) And _
           IsNull(.NumberFormatLocal) And _
           IsNull(.Orientation) And _
           IsNull(.ShrinkToFit) And _
           IsNull(.VerticalAlignment) And _
           IsNull(.WrapText) Then
        Else
            Exit Function
        End If
    End With

    CheckFindFormat = False

End Function

'*****************************************************************************
'[�T�v] �u�����ƒu���v�̃_�C�A���O��\������
'[����] �Ȃ�
'[�ߒl] �_�C�A���O�̃E�B���h�E�n���h��
'*****************************************************************************
Private Function ShowFindDialog() As Long
    Dim objTmpBar As CommandBar
    
    Set objTmpBar = CommandBars.Add(, msoBarPopup, , True)
    With objTmpBar.Controls.Add(, 1849)
        Call ActiveCell.Select  'Shape���I������Ă���Ɨ�O�ɂȂ邽��
        .Execute
    End With
    Call objTmpBar.Delete

    ShowFindDialog = GetDialogHandle()
End Function

