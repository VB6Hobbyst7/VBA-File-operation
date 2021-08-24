# VBA-Renew-Modules
- License: The MIT license

- Copyright (c) 2021 YujiFukami

- 開発テスト環境 Excel: Microsoft® Excel® 2019 32bit 

- 開発テスト環境 OS: Windows 10 Pro

その他、実行環境など報告していただくと感謝感激雨霰。

# 説明
ファイル操作関係のプロシージャがたくさん入ったモジュール「ModFile」

# 使い方

## 設定
実行サンプル「実行サンプル ファイル操作.xlsm」の中の設定は以下の通り。


### 設定1（使用モジュール）

-  ModFile.bas

### 設定2（参照ライブラリ）

- Microsoft Scripting Runtime
	プロシージャ「GetFiles」にてFileSystemObjectを使用するため
- Microsoft XML, v6.0
	プロシージャ「OutputXML」にて必要

## 現在「ModFile.bas」にて使用できるプロシージャ一覧
- SaveSheetAsBook	…シートをブックで保存
- GetSheetByName	…シート名指定でシートオブジェクト取得
- InputCSV		…CSVファイル読込	
- InputBook		…ブックの値読込
- SelectFile		…ファイル選択
- SelectFolder		…フォルダ選択
- GetFileDataTime	…ファイルのタイムスタンプ取得
- MakeFolder		…フォルダ作成
- GetRowCountTextFile	…テキストファイルの行数取得
- GetCurrentFolder	…カレントフォルダ取得
- SetCurrentFolder	…カレントフォルダ設定
- GetExtension		…ファイルの拡張子取得
- OpenFolder		…フォルダを開く
- OpenFile		…ファイルを開く
- OpenApplication	…アプリケーションを開く
- OutputCSV		…CSVファイル出力
- OutputText		…TXTファイル出力
- InputText		…TXTファイル読込
- GetFiles		…フォルダ内のファイル一覧取得
- GetSubFolders		…フォルダ内のサブフォルダ一覧取得
- OutputPDF		…指定シートのPDF出力
- OutputXML		…XMLデータ出力
- GetFileName		…ファイルパスからファイル名取得

