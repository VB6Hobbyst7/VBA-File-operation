Attribute VB_Name = "ModLibrary"
Option Explicit

'�e�v���V�[�W���p�̃��C�u�����v���V�[�W��
Sub TestLib���͔z��������p�ɕϊ�()
    
    '�J�n�ԍ���1�ȊO��2�����z��(1�����v�f����1)
    Dim Hairetu
    ReDim Hairetu(0 To 0, -1 To 2)
    Hairetu(0, -1) = 1
    Hairetu(0, 0) = 2
    Hairetu(0, 1) = 3
    Hairetu(0, 2) = 4
    
    Dim Dummy
    Dummy = Lib���͔z��������p�ɕϊ�(Hairetu)
    Call DPH(Hairetu, , "�ϊ��O")
    Call DPH(Dummy, , "�ϊ���")
    
    Debug.Print "------------------------------"
    
    '�J�n�ԍ���1�ȊO��1�����z��
    ReDim Hairetu(-1 To 3)
    Hairetu(-1) = "A"
    Hairetu(0) = "B"
    Hairetu(1) = "C"
    Hairetu(2) = "D"
    Hairetu(3) = "E"
    Dummy = Lib���͔z��������p�ɕϊ�(Hairetu)
    Call DPH(Hairetu, , "�ϊ��O")
    Call DPH(Dummy, , "�ϊ���")
    
    Debug.Print "------------------------------"
    
    '���l�������͕�����
    Hairetu = "B"
    Dummy = Lib���͔z��������p�ɕϊ�(Hairetu)
'    Call DPH(Hairetu, , "�ϊ��O")
    Call DPH(Dummy, , "�ϊ���")
    
    Debug.Print "------------------------------"
    
    '�J�n�ԍ���1�ȊO��2�����z��
    ReDim Hairetu(0 To 1, 1 To 2)
    Hairetu(0, 1) = "A"
    Hairetu(0, 2) = "B"
    Hairetu(1, 1) = "C"
    Hairetu(1, 2) = "D"
    Dummy = Lib���͔z��������p�ɕϊ�(Hairetu)
    Call DPH(Hairetu, , "�ϊ��O")
    Call DPH(Dummy, , "�ϊ���")
      
    Debug.Print "------------------------------"
    
    '�J�n�ԍ���1��2�����z��
    ReDim Hairetu(1 To 2, 1 To 2)
    Hairetu(1, 1) = "A"
    Hairetu(1, 2) = "B"
    Hairetu(2, 1) = "C"
    Hairetu(2, 2) = "D"
    Dummy = Lib���͔z��������p�ɕϊ�(Hairetu)
    Call DPH(Hairetu, , "�ϊ��O")
    Call DPH(Dummy, , "�ϊ���")
     
End Sub

Function Lib���͔z��������p�ɕϊ�(InputHairetu)
'���͂����z��������p�ɕϊ�����
'1�����z��2�����z��
'���l��������2�����z��(1,1)
'�v�f�̊J�n�ԍ���1�ɂ���
'20210721

    Dim Output
    Dim I%, J%, K%, M%, N% '�����グ�p(Integer�^)
    Dim Base1%, Base2%
    If IsArray(InputHairetu) = False Then
        '�z��łȂ��ꍇ(���l��������)
        ReDim Output(1 To 1, 1 To 1)
        Output(1, 1) = InputHairetu
    Else
        On Error Resume Next
        M = UBound(InputHairetu, 2)
        On Error GoTo 0
        If M = 0 Then
            '1�����z��
            Output = WorksheetFunction.Transpose(InputHairetu)
        Else
            '2�����z��
            Base1 = LBound(InputHairetu, 1)
            Base2 = LBound(InputHairetu, 2)
            
            If Base1 <> 1 Or Base2 <> 1 Then
                N = UBound(InputHairetu, 1)
                If N = Base1 Then
                    '(1,M)�z��
                    ReDim Output(1 To 1, 1 To M - Base2 + 1)
                    For I = 1 To M - Base2 + 1
                        Output(1, I) = InputHairetu(Base1, Base2 + I - 1)
                    Next I
                Else
                    Output = WorksheetFunction.Transpose(InputHairetu)
                    Output = WorksheetFunction.Transpose(Output)
                End If
            Else
                Output = InputHairetu
            End If
        End If
    End If
    
    Lib���͔z��������p�ɕϊ� = Output
    
End Function

