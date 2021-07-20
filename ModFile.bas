Attribute VB_Name = "ModFile"
Option Explicit
Sub SaveSheetAsBook(TargetSheet As Worksheet, Optional SaveName$, Optional SavePath$, _
                           Optional MessageIruNaraTrue As Boolean = False)
'指定のシートを別ブックで保存する 関数名代替
'20210719作成
                           
    Call シートを別ブックで保存(TargetSheet, SaveName, SavePath, MessageIruNaraTrue)
                           
End Sub
Sub シートを別ブックで保存(TargetSheet As Worksheet, Optional SaveName$, Optional SavePath$, _
                           Optional MessageIruNaraTrue As Boolean = False)
'指定のシートを別ブックで保存する
'20210719作成
                           
    '入力引数の調整
    If SaveName = "" Then
        SaveName = TargetSheet.Name
    End If
    
    If SavePath = "" Then
        SavePath = TargetSheet.Parent.Path
    End If
    
    '別ブックで保存
    TargetSheet.Copy
    ActiveWorkbook.SaveAs SavePath & "\" & SaveName
    ActiveWorkbook.Close
    
    If MessageIruNaraTrue Then
        MsgBox ("シート名「" & TargetSheet.Name & "」を" & vbLf & _
               "「" & SavePath & "」に" & vbLf & _
               "ファイル名「" & SaveName & ".xlsx」で保存しました。")
    End If
    
End Sub

Function GetSheetByName(SheetName&) As Worksheet
'指定の名前のシートをワークシートオブジェクトとして取得する 関数代替

    Set GetSheetByName = 指定名のシート取得(SheetName)

End Function
Function 指定名のシート取得(SheetName$) As Worksheet
'指定の名前のシートをワークシートオブジェクトとして取得する
'20210715作成

    Dim Output As Worksheet
    On Error Resume Next
    Set Output = ThisWorkbook.Sheets(SheetName)
    On Error GoTo 0
    
    If Output Is Nothing Then
        MsgBox ("「" & SheetName & "」シートがありません！！")
        End
    End If
    
    Set 指定名のシート取得 = Output

End Function

Function InputCSV(CSVPath$)
'CSVファイルを読み込んで配列形式で返す
'20210706作成

    '入力値確認
    Dim Dummy
    If Dir(CSVPath, vbDirectory) = "" Then
        Dummy = MsgBox(CSVPath & "のファイルは存在しません", vbOKOnly + vbCritical)
        Exit Function
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
        
    InputCSV = Output
    
End Function
Function InputFromBook(BookFolderPath$, BookName$, SheetName$, StartCellAddress$, Optional EndCellAddress$)
'ブックを開かないでデータを取得する
'ExecuteExcel4Macroを使用するので、Excelのバージョンアップの時に注意
'20210720

'BookFolderPath・・・指定ブックのフォルダパス
'BookName・・・指定ブックの名前 拡張子含む
'SheetName・・・指定ブックの取得対象となるシートの名前
'StartCellAddress・・・取得範囲の最初のセルアドレス(例:"A1")
'EndCellAddress・・・取得範囲の最後のセル(例："B3")（省略ならStartCellAddressと同じ）
    
    Dim Rs&, Re&, Cs&, Ce& '始端行,列番号および終端行,列番号(Long型)
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
    
    '始点、終点の反転している場合の処理
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

    Dim I%, J%, K%, M%, N% '数え上げ用(Integer型)
    Dim Output
    
    If Rs = Re And Cs = Ce Then
        '1つのセルだけから取得する場合はその値を返す
        strRC = "R" & Rs & "C" & Cs
        Output = ExecuteExcel4Macro("'" & BookFolderPath & "\[" & BookName & "]" & SheetName & "'!" & strRC)
    Else
        '複数セルから取得する場合は配列で返す
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
'SelectFileの実行サンプル
'20210720

    Dim FolderPath$
    Dim strFileName$
    Dim strExtentions$
    FolderPath = "" 'ActiveWorkbook.Path
    strFileName = "" '"Excelブック"   '←←←←←←←←←←←←←←←←←←←←←←←
    strExtentions = "" '"*.xls; *.xlsx; *.xlsm" '←←←←←←←←←←←←←←←←←←←←←←←
    
    Dim FilePath$
    FilePath = SelectFile(FolderPath, strFileName, strExtentions)
    
End Sub
Function SelectFile(Optional FolderPath$, Optional strFileName$ = "", Optional strExtentions$ = "")
'ファイルを選択するダイアログを表示してファイルを選択させる
'選択したファイルのフルパスを返す
'20210720

'FolderPath・・・最初に開くフォルダ 指定しない場合はカレントフォルダパス
'strFileName・・・選択するファイルの名前  例：Excelブック
'strExtentions・・・選択するファイルの拡張子　例："*.xls; *.xlsx; *.xlsm"

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir 'カレントフォルダ
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
            MsgBox ("ファイルが選択されなかったので終了します")
            End
        End If
    End With
    
    SelectFile = Output
    
End Function
Private Sub SelectFolderTest()
'SelectFolderの実行サンプル
'20210720

    Dim FolderPath$
    FolderPath = ActiveWorkbook.Path
    
    Dim FilePath$
    FilePath = SelectFolder(FolderPath)
    
End Sub
Function SelectFolder(Optional FolderPath$)
'フォルダを選択するダイアログを表示してファイルを選択させる
'選択したフォルダのフルパスを返す
'20210720

'FolderPath・・・最初に開くフォルダ 指定しない場合はカレントフォルダパス

    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFolderPicker)
    
    If FolderPath = "" Then
        FolderPath = CurDir 'カレントフォルダ
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
            MsgBox ("フォルダが選択されなかったので終了します")
            End
        End If
    End With
    
    SelectFolder = Output
    
End Function
Function GetFileDateTime(FilePath$)
'ファイルのタイムスタンプを取得する。
'関数思い出し用
'20210720

'FilePath・・・タイムスタンプを取得するファイルのフルパス

    GetFileDateTime = FileDateTime(FilePath)
    
End Function
Sub MakeFolder(FolderPath$)
'フォルダを作成する
'20210720

'FilePath・・・作成するフォルダのフルパス

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
'テキストファイル、CSVファイルの行数を取得する
'20210720

    'ファイルの存在確認
    If Dir(FilePath, vbDirectory) = "" Then
        MsgBox ("「" & FilePath & "」がありません" & vbLf & _
                "終了します")
        End
    End If
    
    Dim Output&
    With CreateObject("Scripting.FileSystemObject")
        Output = .OpenTextFile(FilePath, 8).Line
    End With
    
    GetRowCountTextFile = Output
    
End Function
Function GetCurrentFolder()
'カレントフォルダのパスを取得
'関数思い出し用
'20210720

    GetCurrentFolder = CurDir
    
End Function
Sub SetCurrentFolder(FolderPath$)
'指定フォルダパスをカレントフォルダを設定
'フォルダパスがネットワークドライブ上のフォルダか自動的に判定して
'ネットワークドライブ上のフォルダもカレントフォルダに設定できる
'20210720

    If Dir(FolderPath, vbDirectory) = "" Then
        MsgBox ("「" & FolderPath & "」がありません" & vbLf & _
                "終了します")
        End
    End If
    
    If Mid(FolderPath, 1, 2) = "\\" Then
        'ネットワークドライブの場合
        Call SetCurrentFolderNetworkDrive(FolderPath)
    Else
        
        'カレントドライブが異なる場合は先に設定する必要がある
        If Mid(FolderPath, 1, 1) <> Mid(CurDir, 1, 1) Then
            ChDrive Mid(FolderPath, 1, 1)
        End If
        
        'カレントフォルダ設定
        ChDir FolderPath
    End If
    
End Sub
Sub SetCurrentFolderNetworkDrive(NetworkFolderPath$)
'ネットワークドライブ上のフォルダパスをカレントフォルダに設定する
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
'ファイルの拡張子を取得する
'20210720

    Dim Output$
    With CreateObject("Scripting.FileSystemObject")
        Output = .GetExtensionName(FilePath)
    End With
    GetExtension = Output
    
End Function
