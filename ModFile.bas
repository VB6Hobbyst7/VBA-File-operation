Attribute VB_Name = "ModFile"
Option Explicit
Sub SaveSheetAsBook(TargetSheet As Worksheet, Optional SaveName$, Optional SavePath$, _
                           Optional MessageIruNaraTrue As Boolean = False)
'�w��̃V�[�g��ʃu�b�N�ŕۑ����� �֐������
'20210719�쐬
                           
    Call �V�[�g��ʃu�b�N�ŕۑ�(TargetSheet, SaveName, SavePath, MessageIruNaraTrue)
                           
End Sub
Sub �V�[�g��ʃu�b�N�ŕۑ�(TargetSheet As Worksheet, Optional SaveName$, Optional SavePath$, _
                           Optional MessageIruNaraTrue As Boolean = False)
'�w��̃V�[�g��ʃu�b�N�ŕۑ�����
'20210719�쐬
                           
    '���͈����̒���
    If SaveName = "" Then
        SaveName = TargetSheet.Name
    End If
    
    If SavePath = "" Then
        SavePath = TargetSheet.Parent.Path
    End If
    
    '�ʃu�b�N�ŕۑ�
    TargetSheet.Copy
    ActiveWorkbook.SaveAs SavePath & "\" & SaveName
    ActiveWorkbook.Close
    
    If MessageIruNaraTrue Then
        MsgBox ("�V�[�g���u" & TargetSheet.Name & "�v��" & vbLf & _
               "�u" & SavePath & "�v��" & vbLf & _
               "�t�@�C�����u" & SaveName & ".xlsx�v�ŕۑ����܂����B")
    End If
    
End Sub

Function GetSheetByName(SheetName&) As Worksheet
'�w��̖��O�̃V�[�g�����[�N�V�[�g�I�u�W�F�N�g�Ƃ��Ď擾���� �֐����

    Set GetSheetByName = �w�薼�̃V�[�g�擾(SheetName)

End Function
Function �w�薼�̃V�[�g�擾(SheetName$) As Worksheet
'�w��̖��O�̃V�[�g�����[�N�V�[�g�I�u�W�F�N�g�Ƃ��Ď擾����
'20210715�쐬

    Dim Output As Worksheet
    On Error Resume Next
    Set Output = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    
    If Output Is Nothing Then
        MsgBox ("�u" & SheetName & "�v�V�[�g������܂���I�I")
        End
    End If
    
    Set �w�薼�̃V�[�g�擾 = Output

End Function

Function InputCSV(CSVPath$)
'CSV�t�@�C����ǂݍ���Ŕz��`���ŕԂ�
'20210706�쐬

    '���͒l�m�F
    Dim Dummy
    If Dir(CSVPath, vbDirectory) = "" Then
        Dummy = MsgBox(CSVPath & "�̃t�@�C���͑��݂��܂���", vbOKOnly + vbCritical)
        Exit Function
    End If
    
    Dim intFree As Integer
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    Dim TmpStr$, TmpSplit
    Dim StrList
    Dim Output

    intFree = FreeFile '��ԍ����擾
    Open CSVPath For Input As #intFree 'CSV�t�@�B�����I�[�v��
    
    K = 0
    ReDim StrList(1 To 1)
    Do Until EOF(intFree)
        Line Input #intFree, TmpStr '1�s�ǂݍ���
        K = K + 1
        ReDim Preserve StrList(1 To K)
        StrList(K) = TmpStr
        
        M = WorksheetFunction.Max(UBound(Split(TmpStr, ",")) + 1, M)
    Loop
        
    Close #intFree
    N = K
    ReDim Output(1 To N, 1 To M)
    
    For I = 1 To N
        TmpStr = StrList(I)
        TmpSplit = Split(TmpStr, ",")
        
        For J = 0 To UBound(TmpSplit)
            Output(I, J + 1) = TmpSplit(J)
        Next J
    Next I
        
    InputCSV = Output
    
End Function
Function InputFromBook(BookFolderPath$, BookName$, SheetName$, StartCellAddress$, Optional EndCellAddress$)
'�u�b�N���J���Ȃ��Ńf�[�^���擾����
'ExecuteExcel4Macro���g�p����̂ŁAExcel�̃o�[�W�����A�b�v�̎��ɒ���
'20210720

'BookFolderPath�E�E�E�w��u�b�N�̃t�H���_�p�X
'BookName�E�E�E�w��u�b�N�̖��O �g���q�܂�
'SheetName�E�E�E�w��u�b�N�̎擾�ΏۂƂȂ�V�[�g�̖��O
'StartCellAddress�E�E�E�擾�͈͂̍ŏ��̃Z���A�h���X(��:"A1")
'EndCellAddress�E�E�E�擾�͈͂̍Ō�̃Z��(��F"B3")�i�ȗ��Ȃ�StartCellAddress�Ɠ����j
    
    Dim Rs&, Re&, Cs&, Ce& '�n�[�s,��ԍ�����яI�[�s,��ԍ�(Long�^)
    Dim strRC$
    With Range(StartCellAddress)
        Rs = .Row
        Cs = .Column
    End With
    
    If EndCellAddress = "" Then
        Re = Rs
        Ce = Cs
    Else
        With Range(EndCellAddress)
            Re = .Row
            Ce = .Column
        End With
    End If
    
    '�n�_�A�I�_�̔��]���Ă���ꍇ�̏���
    Dim Dummy&
    If Re < Rs Then
        Dummy = Rs
        Re = Rs
        Rs = Dummy
    End If
    
    If Ce < Cs Then
        Dummy = Cs
        Ce = Cs
        Cs = Dummy
    End If

    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Output
    
    If Rs = Re And Cs = Ce Then
        '1�̃Z����������擾����ꍇ�͂��̒l��Ԃ�
        strRC = "R" & Rs & "C" & Cs
        Output = ExecuteExcel4Macro("'" & BookFolderPath & "\[" & BookName & "]" & SheetName & "'!" & strRC)
    Else
        '�����Z������擾����ꍇ�͔z��ŕԂ�
        ReDim Output(1 To Re - Rs + 1, 1 To Ce - Cs + 1)
        
        For I = Rs To Re
            For J = Cs To Ce
                strRC = "R" & I & "C" & J
                Output(I, J) = ExecuteExcel4Macro("'" & BookFolderPath & "\[" & BookName & "]" & SheetName & "'!" & strRC)
            Next J
        Next I
    End If
    
    InputFromBook = Output
    
End Function
Private Sub SelectFileTest()
'SelectFile�̎��s�T���v��
'20210720

    Dim FolderPath$
    Dim strFileName$
    Dim strExtentions$
    FolderPath = "" 'ActiveWorkbook.Path
    strFileName = "" '"Excel�u�b�N"   '����������������������������������������������
    strExtentions = "" '"*.xls; *.xlsx; *.xlsm" '����������������������������������������������
    
    Dim FilePath$
    FilePath = SelectFile(FolderPath, strFileName, strExtentions)
    
End Sub
Function SelectFile(Optional FolderPath$, Optional strFileName$ = "", Optional strExtentions$ = "")
'�t�@�C����I������_�C�A���O��\�����ăt�@�C����I��������
'�I�������t�@�C���̃t���p�X��Ԃ�
'20210720

'FolderPath�E�E�E�ŏ��ɊJ���t�H���_ �w�肵�Ȃ��ꍇ�̓J�����g�t�H���_�p�X
'strFileName�E�E�E�I������t�@�C���̖��O  ��FExcel�u�b�N
'strExtentions�E�E�E�I������t�@�C���̊g���q�@��F"*.xls; *.xlsx; *.xlsm"

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir '�J�����g�t�H���_
    End If
    
    Dim Output$
    
    With FD
        With .Filters
            .Clear
            .Add strFileName, strExtentions, 1
        End With
        .InitialFileName = FolderPath & "\"
        If .Show = True Then
            Output = .SelectedItems(1)
        Else
            MsgBox ("�t�@�C�����I������Ȃ������̂ŏI�����܂�")
            End
        End If
    End With
    
    SelectFile = Output
    
End Function
Private Sub SelectFolderTest()
'SelectFolder�̎��s�T���v��
'20210720

    Dim FolderPath$
    FolderPath = ActiveWorkbook.Path
    
    Dim FilePath$
    FilePath = SelectFolder(FolderPath)
    
End Sub
Function SelectFolder(Optional FolderPath$)
'�t�H���_��I������_�C�A���O��\�����ăt�@�C����I��������
'�I�������t�H���_�̃t���p�X��Ԃ�
'20210720

'FolderPath�E�E�E�ŏ��ɊJ���t�H���_ �w�肵�Ȃ��ꍇ�̓J�����g�t�H���_�p�X

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir '�J�����g�t�H���_
    End If
    
    Dim Output$
    
    With FD
        With .Filters
            .Clear
        End With
        .InitialFileName = FolderPath & "\"
        If .Show = True Then
            Output = .SelectedItems(1)
        Else
            MsgBox ("�t�H���_���I������Ȃ������̂ŏI�����܂�")
            End
        End If
    End With
    
    SelectFolder = Output
    
End Function
Function GetFileDateTime(FilePath$)
'�t�@�C���̃^�C���X�^���v���擾����B
'�֐��v���o���p
'20210720

'FilePath�E�E�E�^�C���X�^���v���擾����t�@�C���̃t���p�X

    GetFileDateTime = FileDateTime(FilePath)
    
End Function
Sub MakeFolder(FolderPath$)
'�t�H���_���쐬����
'20210720

'FilePath�E�E�E�쐬����t�H���_�̃t���p�X

    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FilePath
    End If
End Sub
Sub GetRowCountTextFileTest()
    
    Dim FilePath$
    FilePath = ActiveWorkbook.Path & "\" & "TestText.txt"
    
    Dim RowCount&
    RowCount = GetRowCountTextFile(FilePath)
    
End Sub
Function GetRowCountTextFile(FilePath$)
'�e�L�X�g�t�@�C���ACSV�t�@�C���̍s�����擾����
'20210720

    '�t�@�C���̑��݊m�F
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("�u" & FilePath & "�v������܂���" & vbLf & _
                "�I�����܂�")
        End
    End If
    
    Dim Output&
    With CreateObject("Scripting.FileSystemObject")
        Output = .OpenTextFile(FilePath, 8).Line
    End With
    
    GetRowCountTextFile = Output
    
End Function
Function GetCurrentFolder()
'�J�����g�t�H���_�̃p�X���擾
'�֐��v���o���p
'20210720

    GetCurrentFolder = CurDir
    
End Function
Sub SetCurrentFolder(FolderPath$)
'�w��t�H���_�p�X���J�����g�t�H���_��ݒ�
'�t�H���_�p�X���l�b�g���[�N�h���C�u��̃t�H���_�������I�ɔ��肵��
'�l�b�g���[�N�h���C�u��̃t�H���_���J�����g�t�H���_�ɐݒ�ł���
'20210720

    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("�u" & FolderPath & "�v������܂���" & vbLf & _
                "�I�����܂�")
        End
    End If
    
    If Mid(FolderPath, 1, 2) = "\\" Then
        '�l�b�g���[�N�h���C�u�̏ꍇ
        Call SetCurrentFolderNetworkDrive(FolderPath)
    Else
        
        '�J�����g�h���C�u���قȂ�ꍇ�͐�ɐݒ肷��K�v������
        If Mid(FolderPath, 1, 1) <> Mid(CurDir, 1, 1) Then
            ChDrive Mid(FolderPath, 1, 1)
        End If
        
        '�J�����g�t�H���_�ݒ�
        ChDir FolderPath
    End If
    
End Sub
Sub SetCurrentFolderNetworkDrive(NetworkFolderPath$)
'�l�b�g���[�N�h���C�u��̃t�H���_�p�X���J�����g�t�H���_�ɐݒ肷��
'20210720

    With CreateObject("WScript.Shell")
        .CurrentDirectory = NetworkFolderPath
    End With
    
End Sub
Private Sub GetExtensionTest()
    
    Dim Dummy
    Dummy = GetExtension(ActiveWorkbook.Path & "\" & ActiveWorkbook.Name)
    
End Sub
Function GetExtension(FilePath$)
'�t�@�C���̊g���q���擾����
'20210720

    Dim Output$
    With CreateObject("Scripting.FileSystemObject")
        Output = .GetExtensionName(FilePath)
    End With
    GetExtension = Output
    
End Function
