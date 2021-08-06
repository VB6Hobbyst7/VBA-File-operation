Attribute VB_Name = "ModLibrary"
Option Explicit

'各プロシージャ用のライブラリプロシージャ
Sub TestLib入力配列を処理用に変換()
    
    '開始番号が1以外の2次元配列(1次元要素数が1)
    Dim Hairetu
    ReDim Hairetu(0 To 0, -1 To 2)
    Hairetu(0, -1) = 1
    Hairetu(0, 0) = 2
    Hairetu(0, 1) = 3
    Hairetu(0, 2) = 4
    
    Dim Dummy
    Dummy = Lib入力配列を処理用に変換(Hairetu)
    Call DPH(Hairetu, , "変換前")
    Call DPH(Dummy, , "変換後")
    
    Debug.Print "------------------------------"
    
    '開始番号が1以外の1次元配列
    ReDim Hairetu(-1 To 3)
    Hairetu(-1) = "A"
    Hairetu(0) = "B"
    Hairetu(1) = "C"
    Hairetu(2) = "D"
    Hairetu(3) = "E"
    Dummy = Lib入力配列を処理用に変換(Hairetu)
    Call DPH(Hairetu, , "変換前")
    Call DPH(Dummy, , "変換後")
    
    Debug.Print "------------------------------"
    
    '数値もしくは文字列
    Hairetu = "B"
    Dummy = Lib入力配列を処理用に変換(Hairetu)
'    Call DPH(Hairetu, , "変換前")
    Call DPH(Dummy, , "変換後")
    
    Debug.Print "------------------------------"
    
    '開始番号が1以外の2次元配列
    ReDim Hairetu(0 To 1, 1 To 2)
    Hairetu(0, 1) = "A"
    Hairetu(0, 2) = "B"
    Hairetu(1, 1) = "C"
    Hairetu(1, 2) = "D"
    Dummy = Lib入力配列を処理用に変換(Hairetu)
    Call DPH(Hairetu, , "変換前")
    Call DPH(Dummy, , "変換後")
      
    Debug.Print "------------------------------"
    
    '開始番号が1の2次元配列
    ReDim Hairetu(1 To 2, 1 To 2)
    Hairetu(1, 1) = "A"
    Hairetu(1, 2) = "B"
    Hairetu(2, 1) = "C"
    Hairetu(2, 2) = "D"
    Dummy = Lib入力配列を処理用に変換(Hairetu)
    Call DPH(Hairetu, , "変換前")
    Call DPH(Dummy, , "変換後")
     
End Sub

Function Lib入力配列を処理用に変換(InputHairetu)
'入力した配列を処理用に変換する
'1次元配列→2次元配列
'数値か文字列→2次元配列(1,1)
'要素の開始番号を1にする
'20210721

    Dim Output
    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Base1%, Base2%
    If IsArray(InputHairetu) = False Then
        '配列でない場合(数値か文字列)
        ReDim Output(1 To 1, 1 To 1)
        Output(1, 1) = InputHairetu
    Else
        On Error Resume Next
        M = UBound(InputHairetu, 2)
        On Error GoTo 0
        If M = 0 Then
            '1次元配列
            Output = WorksheetFunction.Transpose(InputHairetu)
        Else
            '2次元配列
            Base1 = LBound(InputHairetu, 1)
            Base2 = LBound(InputHairetu, 2)
            
            If Base1 <> 1 Or Base2 <> 1 Then
                N = UBound(InputHairetu, 1)
                If N = Base1 Then
                    '(1,M)配列
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
    
    Lib入力配列を処理用に変換 = Output
    
End Function

