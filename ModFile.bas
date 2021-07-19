Attribute VB_Name = "ModFile"
Option Explicit
Function InputCSV(CSVPath$)
'CSVファイルを読み込んで配列形式で返す 関数代替
    InputCSV = CSV読込(CSVPath)

End Function
Function CSV読込(CSVPath$)
'CSVファイルを読み込んで配列形式で返す
'20210706作成

    '入力値確認
    Dim Dummy
    If Dir(CSVPath, vbDirectory) = "" Then
        Dummy = MsgBox(CSVPath & "のファイルは存在しません", vbOKOnly + vbCritical)
        Exit Sub
    End If
    
    Dim intFree As Integer
    Dim I&, J&, K&, M&, N& '数え上げ用(Long型)
    Dim TmpStr$, TmpSplit
    Dim StrList
    Dim Output

    intFree = FreeFile '空番号を取得
    Open CSVPath For Input As #intFree 'CSVファィルをオープン
    
    K = 0
    ReDim StrList(1 To 1)
    Do Until EOF(intFree)
        Line Input #intFree, TmpStr '1行読み込み
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
        
    CSV読込 = Output
    
End Function
