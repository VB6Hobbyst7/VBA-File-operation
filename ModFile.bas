Attribute VB_Name = "ModFile"
Option Explicit
Function InputCSV(CSVPath$)
'CSV�t�@�C����ǂݍ���Ŕz��`���ŕԂ� �֐����
    InputCSV = CSV�Ǎ�(CSVPath)

End Function
Function CSV�Ǎ�(CSVPath$)
'CSV�t�@�C����ǂݍ���Ŕz��`���ŕԂ�
'20210706�쐬

    '���͒l�m�F
    Dim Dummy
    If Dir(CSVPath, vbDirectory) = "" Then
        Dummy = MsgBox(CSVPath & "�̃t�@�C���͑��݂��܂���", vbOKOnly + vbCritical)
        Exit Sub
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
        
    CSV�Ǎ� = Output
    
End Function
