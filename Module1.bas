Attribute VB_Name = "Module1"
'================================================================
' Excel VBAコード（改修版：v4）
' CAM・WS・NCフォルダ作成＋ファイルコピー＋ショートカット作成マクロ
'================================================================

Dim キャンセルメッセージ As String

Sub メイン処理()
    キャンセルメッセージ = ""  ' 初期化

    ' --- ユーザーフォームを表示して入力値を名前付き範囲へ書き込み ---
    frmInput.Show
    If frmInput.Cancelled Then Exit Sub

    ' --- 入力値の検証 ---
    If Range("NC_機械番号").Value = "" And Range("コード番").Value = "" Then
        MsgBox "NC_機械番号を入力してください。", vbExclamation
        Exit Sub
    End If

    Call 出力セルの値を削除

    If Not CAMの処理 Then Exit Sub

    Call WS手順書の処理

    Call WS表紙の処理

    Call NCの処理

    ' --- スケジュールシートのテーブルへ記録を追記 ---
    Call スケジュール記録追記

    If キャンセルメッセージ <> "" Then
        MsgBox "処理が完了しました。" & vbCrLf & vbCrLf & キャンセルメッセージ & vbCrLf & "のファイルが既に存在したため新規ファイル保存をキャンセルしました。", vbInformation
    Else
        MsgBox "処理が完了しました。", vbInformation
    End If
End Sub

Private Sub 出力セルの値を削除()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ws.Range("ショートカットフォルダ名").Value = ""
    ws.Range("CAM_新規フォルダ名").Value = ""
    ws.Range("CAM_品番番号フォルダ名").Value = ""
    ws.Range("CAM_新規ファイル名").Value = ""
    ws.Range("WS_新規フォルダ名").Value = ""
    ws.Range("WS_品番番号フォルダ名").Value = ""
    ws.Range("WS_新規ファイル名").Value = ""
    ws.Range("WS_表紙ファイル名").Value = ""
    ws.Range("NC_新規フォルダ名").Value = ""
    ws.Range("NC_品番番号フォルダ名").Value = ""
End Sub

Private Function CAMの処理() As Boolean
    On Error GoTo エラー処理
    CAMの処理 = False

    Dim フォルダパス As String, 新規フォルダ名 As String, 品番フォルダ名 As String

    If Range("CAM有効").Value <> "有効" Then CAMの処理 = True: Exit Function

    If Range("CAM_図番フォルダ作成").Value = "有効" Then
        Dim 親フォルダ As String
        Dim 図番文字列 As String
        Dim 対象フォルダ As String
        Dim 見つかった As Collection
        Dim ファイルシステム As Object
        Dim フォルダ As Object

        Set 見つかった = New Collection
        Set ファイルシステム = CreateObject("Scripting.FileSystemObject")

        図番文字列 = Range("図番").Value
        If 図番文字列 = "" Then
            図番文字列 = Range("指示番号").Value
        End If

        親フォルダ = Range("CAM_保存先").Value

        If ファイルシステム.FolderExists(親フォルダ) Then
            For Each フォルダ In ファイルシステム.GetFolder(親フォルダ).SubFolders
                If InStr(1, フォルダ.Name, 図番文字列, vbTextCompare) > 0 Then
                    見つかった.Add フォルダ.Name
                End If
            Next
        End If

        If 見つかった.Count = 1 Then
            Range("CAM_新規フォルダ名").Value = 見つかった(1)
            フォルダパス = 親フォルダ & "\" & 見つかった(1)
        ElseIf 見つかった.Count > 1 Then
            MsgBox "「3Dモデルデータ」に" & vbCrLf & _
                   """" & 図番文字列 & """ フォルダが複数存在しています" & vbCrLf & _
                   "処理を中止します。", vbCritical
            Exit Function
        Else
            If Range("図番").Value <> "" Then
                新規フォルダ名 = Range("図番").Value & "_" & Range("製品名").Value
            Else
                新規フォルダ名 = Range("指示番号").Value & "_" & Range("製品名").Value
            End If
            Range("CAM_新規フォルダ名").Value = 新規フォルダ名
            フォルダパス = 親フォルダ & "\" & 新規フォルダ名
            Call フォルダ作成(フォルダパス)
        End If
    End If

    If Range("CAM_品番番号フォルダ作成").Value = "有効" Then
        品番フォルダ名 = "No" & Range("部品番号").Value & "_" & Range("製品名").Value
        Range("CAM_品番番号フォルダ名").Value = 品番フォルダ名
        Call フォルダ作成(フォルダパス & "\" & 品番フォルダ名)
        Call フォルダ作成(フォルダパス & "\" & 品番フォルダ名 & "\設計者の支援データ")
    Else
        品番フォルダ名 = ""
    End If

    Dim 新規ファイル名 As String, 元ファイル名 As String
    元ファイル名 = Range("CAM_テンプレートファイルパス").Value

    If Range("コード番").Value <> "" Then
        新規ファイル名 = Range("コード番").Value & "_" & Range("図番").Value & "_No" & Range("部品番号").Value & "_" & Range("CAM_元ファイル名").Value
    ElseIf Range("図番").Value <> "" Then
        新規ファイル名 = Range("図番").Value & "_" & Range("指示番号").Value & "_No" & Range("部品番号").Value & "_" & Range("CAM_元ファイル名").Value
    Else
        新規ファイル名 = Range("指示番号").Value & "_No" & Range("部品番号").Value & "_" & Range("CAM_元ファイル名").Value
    End If

    新規ファイル名 = Replace(新規ファイル名, "Ｏ", "O")
    新規ファイル名 = Replace(新規ファイル名, "０", "O")
    新規ファイル名 = Replace(新規ファイル名, "__", "_")
    新規ファイル名 = Replace(新規ファイル名, "___", "_")
    Range("CAM_新規ファイル名").Value = 新規ファイル名

    Dim テンプレートファイル As String, 拡張子 As String
    Dim コピー先フォルダ As String, コピー先フルパス As String
    テンプレートファイル = Range("CAM_テンプレートファイルパス").Value
    拡張子 = Mid(テンプレートファイル, InStrRev(テンプレートファイル, "."))

    If 品番フォルダ名 <> "" Then
        コピー先フォルダ = フォルダパス & "\" & 品番フォルダ名
    Else
        コピー先フォルダ = フォルダパス
    End If

    コピー先フルパス = コピー先フォルダ & "\" & 新規ファイル名 & 拡張子

    If Dir(コピー先フルパス) <> "" Then
        キャンセルメッセージ = キャンセルメッセージ & "CAM、"
    Else
        FileCopy テンプレートファイル, コピー先フルパス
    End If

    Call ショートカットの作成("CAM")

    CAMの処理 = True
    Exit Function

エラー処理:
    MsgBox "CAMの処理中にエラーが発生しました：" & Err.Description, vbCritical
End Function

Private Sub フォルダ作成(フォルダパス As String)
    If Dir(フォルダパス, vbDirectory) = "" Then MkDir フォルダパス
End Sub

Private Sub ショートカットの作成(種別 As String)
    Dim ショートカット作成先 As String
    Dim ショートカットフォルダ As String
    Dim ショートカット名 As String
    Dim シェル As Object, ショートカット As Object
    Dim targetFolder As String

    Dim コード番 As String, 図番 As String, 指示番号 As String
    Dim 製品名 As String, 部品番号 As String, 部品名 As String
    Dim ショートカットフォルダ名 As String

    コード番 = Trim(Range("コード番").Value)
    図番 = Trim(Range("図番").Value)
    指示番号 = Trim(Range("指示番号").Value)
    製品名 = Trim(Range("製品名").Value)
    部品番号 = Trim(Range("部品番号").Value)
    部品名 = Trim(Range("部品名").Value)

    ショートカットフォルダ名 = ""
    If コード番 <> "" Then ショートカットフォルダ名 = コード番
    If 図番 <> "" Then ショートカットフォルダ名 = ショートカットフォルダ名 & IIf(ショートカットフォルダ名 <> "", "_", "") & 図番
    If 指示番号 <> "" Then ショートカットフォルダ名 = ショートカットフォルダ名 & IIf(ショートカットフォルダ名 <> "", "_", "") & 指示番号
    If 製品名 <> "" Then ショートカットフォルダ名 = ショートカットフォルダ名 & IIf(ショートカットフォルダ名 <> "", "_", "") & 製品名
    If 部品番号 <> "" Then ショートカットフォルダ名 = ショートカットフォルダ名 & IIf(ショートカットフォルダ名 <> "", "_", "") & "No" & 部品番号
    If 部品名 <> "" Then ショートカットフォルダ名 = ショートカットフォルダ名 & "_" & 部品名
    Range("ショートカットフォルダ名").Value = ショートカットフォルダ名

    ショートカット作成先 = Range("ショートカット作成先フォルダ").Value
    If Dir(ショートカット作成先, vbDirectory) = "" Then
        MsgBox "ショートカット作成先フォルダが見つかりません。", vbExclamation
        Exit Sub
    End If
    ショートカットフォルダ = ショートカット作成先 & "\" & ショートカットフォルダ名
    If Dir(ショートカットフォルダ, vbDirectory) = "" Then MkDir ショートカットフォルダ

    Select Case 種別
        Case "CAM"
            targetFolder = Range("CAM_保存先").Value & "\" & Range("CAM_新規フォルダ名").Value
        Case "WS"
            targetFolder = IIf(コード番 <> "", Range("WS_異種品保存先").Value, Range("WS_通常保存先").Value) _
                           & "\" & Range("WS_新規フォルダ名").Value
        Case "NC"
            targetFolder = IIf(コード番 <> "", Range("NC_異種品保存先").Value, Range("NC_通常保存先").Value) _
                           & "\" & Range("NC_新規フォルダ名").Value
        Case Else
            MsgBox "不正な種別です。", vbExclamation
            Exit Sub
    End Select

    Set シェル = CreateObject("WScript.Shell")
    If Dir(targetFolder, vbDirectory) <> "" Then
        ショートカット名 = ショートカットフォルダ & "\" & 種別 & "_" & ショートカットフォルダ名 & ".lnk"

        If 種別 = "NC" And Range("NC_機械番号").Value = "5軸" _
            And (InStr(Range("図番").Value, "Ｏ") > 0 Or InStr(Range("図番").Value, "０") > 0) Then
            ショートカット名 = ショートカットフォルダ & "\" & 種別 & "5軸" & "_" & ショートカットフォルダ名 & ".lnk"
        End If

        If Len(ショートカット名) >= 180 Then
            Dim 本体名 As String
            本体名 = Left(ショートカット名, Len(ショートカット名) - 4)
            Do While Len(本体名 & ".lnk") >= 180
                本体名 = Left(本体名, Len(本体名) - 1)
            Loop
            ショートカット名 = 本体名 & ".lnk"
        End If

        Set ショートカット = シェル.CreateShortcut(ショートカット名)
        ショートカット.TargetPath = targetFolder
        ショートカット.Save
    End If
    Set ショートカット = Nothing
    Set シェル = Nothing
End Sub

Private Sub WS手順書の処理()
    Dim フォルダパス As String, 新規フォルダ名 As String, 品番フォルダ名 As String
    Dim 新規ファイル名 As String, 元ファイル名 As String, ファイルフルパス As String
    Dim 拡張子 As String
    Dim 作業ブック As Workbook, 作業シート As Worksheet, 呼出シート As Worksheet

    If Range("WS有効").Value <> "有効" Then Exit Sub

    If Range("コード番").Value <> "" Then
        新規フォルダ名 = Range("コード番").Value
    ElseIf Range("図番").Value <> "" Then
        新規フォルダ名 = Range("図番").Value
    Else
        新規フォルダ名 = Range("指示番号").Value
    End If
    Range("WS_新規フォルダ名").Value = 新規フォルダ名

    If Range("コード番").Value <> "" Then
        フォルダパス = Range("WS_異種品保存先").Value & "\" & 新規フォルダ名
    Else
        フォルダパス = Range("WS_通常保存先").Value & "\" & 新規フォルダ名
    End If

    Call フォルダ作成(フォルダパス)

    品番フォルダ名 = Format(Range("部品番号").Value, "00")
    Range("WS_品番番号フォルダ名").Value = 品番フォルダ名
    Call フォルダ作成(フォルダパス & "\" & 品番フォルダ名)

    新規ファイル名 = "T" & Replace(新規フォルダ名, "-", "") & IIf(品番フォルダ名 <> "", "-" & 品番フォルダ名, "")
    Range("WS_新規ファイル名").Value = 新規ファイル名

    元ファイル名 = Range("WS_テンプレートファイルパス").Value
    拡張子 = Mid(元ファイル名, InStrRev(元ファイル名, "."))

    If 品番フォルダ名 <> "" Then
        ファイルフルパス = フォルダパス & "\" & 品番フォルダ名 & "\" & 新規ファイル名 & 拡張子
    Else
        ファイルフルパス = フォルダパス & "\" & 新規ファイル名 & 拡張子
    End If

    If Dir(ファイルフルパス) <> "" Then
        キャンセルメッセージ = キャンセルメッセージ & "WS手順書、"
    Else
        FileCopy 元ファイル名, ファイルフルパス

        Set 呼出シート = ThisWorkbook.Sheets(1)

        On Error Resume Next
        Set 作業ブック = Workbooks.Open(ファイルフルパス)
        If 作業ブック Is Nothing Then
            MsgBox "ファイルのオープンに失敗しました: " & ファイルフルパス, vbExclamation
            Exit Sub
        End If
        On Error GoTo 0

        Set 作業シート = 作業ブック.Sheets(1)
        Call データ転記(作業シート, 呼出シート)

        作業ブック.Save
        作業ブック.Close False
    End If

    Call ショートカットの作成("WS")
End Sub

Private Sub WS表紙の処理()
    Dim フォルダパス As String, 新規フォルダ名 As String, 品番フォルダ名 As String
    Dim 新規ファイル名 As String, 元ファイル名 As String, ファイルフルパス As String
    Dim 拡張子 As String
    Dim 作業ブック As Workbook, 作業シート As Worksheet, 呼出シート As Worksheet

    If Range("WS有効").Value <> "有効" Then Exit Sub
    If Range("コード番").Value = "" Then Exit Sub
    If Range("WS_異種品表紙作成").Value <> "有効" Then Exit Sub

    新規フォルダ名 = Range("コード番").Value
    フォルダパス = Range("WS_異種品保存先").Value & "\" & 新規フォルダ名
    品番フォルダ名 = "00"
    Call フォルダ作成(フォルダパス & "\" & 品番フォルダ名)

    新規ファイル名 = "S" & Replace(新規フォルダ名, "-", "") & IIf(品番フォルダ名 <> "", "-" & 品番フォルダ名, "")
    Range("WS_表紙ファイル名").Value = 新規ファイル名

    元ファイル名 = Range("WS_表紙テンプレートファイルパス").Value
    拡張子 = Mid(元ファイル名, InStrRev(元ファイル名, "."))
    ファイルフルパス = フォルダパス & "\" & 品番フォルダ名 & "\" & 新規ファイル名 & 拡張子

    If Dir(ファイルフルパス) <> "" Then
        キャンセルメッセージ = キャンセルメッセージ & "WS表紙、"
    Else
        FileCopy 元ファイル名, ファイルフルパス

        Set 呼出シート = ThisWorkbook.Sheets(1)

        On Error Resume Next
        Set 作業ブック = Workbooks.Open(ファイルフルパス)
        If 作業ブック Is Nothing Then
            MsgBox "ファイルのオープンに失敗しました: " & ファイルフルパス, vbExclamation
            Exit Sub
        End If
        On Error GoTo 0

        Set 作業シート = 作業ブック.Sheets(1)
        Call データ転記(作業シート, 呼出シート)

        Set 作業シート = 作業ブック.Sheets(4)
        作業シート.Range("MC日付").Value = Date
        作業シート.Range("MC年度").Value = 呼出シート.Range("WS_異種品表紙作成者").Value
        作業シート.Range("MC状態").Value = "受注"

        作業ブック.Save
        作業ブック.Close False
    End If
End Sub

Private Sub データ転記(作業シート As Worksheet, 呼出シート As Worksheet)
    If 名前範囲確認("コード番", 呼出シート) Then
        作業シート.Range("コード番").Value = 呼出シート.Range("コード番").Value
    End If

    If 名前範囲確認("図番", 呼出シート) Then
        If 呼出シート.Range("図番").Value <> "" Then
            作業シート.Range("図番").Value = 呼出シート.Range("図番").Value
        Else
            作業シート.Range("図番").Value = 呼出シート.Range("指示番号").Value
            作業シート.Range("図番").Offset(0, -1).Value = "指示番号"
        End If
    End If

    If 名前範囲確認("製品名", 呼出シート) Then
        作業シート.Range("製品名").Value = 呼出シート.Range("製品名").Value
    End If

    If 名前範囲確認("部品名", 呼出シート) Then
        作業シート.Range("部品名").Value = "No" & 呼出シート.Range("部品番号").Value & " " & 呼出シート.Range("部品名").Value
    End If
End Sub

' 拡張子の重複を修正する関数
Private Function 修正拡張子(パス As String) As String
    If InStrRev(パス, ".xlsm.xlsm") > 0 Then
        パス = Replace(パス, ".xlsm.xlsm", ".xlsm")
    ElseIf InStrRev(パス, ".xlsx.xlsx") > 0 Then
        パス = Replace(パス, ".xlsx.xlsx", ".xlsx")
    End If
    修正拡張子 = パス
End Function

' 名前付き範囲が存在するか確認する関数
Private Function 名前範囲確認(範囲名 As String, 対象シート As Worksheet) As Boolean
    On Error Resume Next
    名前範囲確認 = Not Intersect(対象シート.Range(範囲名), 対象シート.Cells) Is Nothing
    On Error GoTo 0
End Function

Private Sub NCの処理()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.Range("NC有効").Value <> "有効" Then Exit Sub

    Dim コード番 As String, 図番 As String, 指示番号 As String, NC新規フォルダ名 As String
    コード番 = ws.Range("コード番").Value
    図番 = ws.Range("図番").Value
    指示番号 = ws.Range("指示番号").Value

    If コード番 <> "" Then
        NC新規フォルダ名 = コード番
    ElseIf 図番 <> "" Then
        NC新規フォルダ名 = 図番
        If ws.Range("NC_機械番号").Value = "5軸" Then
            If InStr(図番, "Ｏ") > 0 Or InStr(図番, "０") > 0 Then
                NC新規フォルダ名 = Replace(NC新規フォルダ名, "Ｏ", "O")
                NC新規フォルダ名 = Replace(NC新規フォルダ名, "０", "O")
            End If
        End If
    ElseIf 指示番号 <> "" Then
        NC新規フォルダ名 = 指示番号
    Else
        MsgBox "コード番・図番・指示番号のいずれかを入力してください。", vbExclamation
        Exit Sub
    End If
    ws.Range("NC_新規フォルダ名").Value = NC新規フォルダ名

    Dim NC異種品保存先 As String, NC通常保存先 As String, 保存先フォルダ As String
    NC異種品保存先 = ws.Range("NC_異種品保存先").Value
    NC通常保存先 = ws.Range("NC_通常保存先").Value
    If コード番 <> "" Then
        保存先フォルダ = NC異種品保存先
    Else
        保存先フォルダ = NC通常保存先
    End If

    Dim フォルダパス As String
    フォルダパス = 保存先フォルダ & "\" & NC新規フォルダ名

    If Dir(フォルダパス, vbDirectory) = "" Then
        MkDir フォルダパス
    End If

    If ws.Range("NC_品番番号フォルダ作成").Value = "有効" Then
        Dim 部品番号 As String, NC品番番号フォルダ名 As String
        部品番号 = ws.Range("部品番号").Value

        If IsNumeric(部品番号) And Len(部品番号) = 1 Then
            NC品番番号フォルダ名 = "0" & 部品番号
        Else
            NC品番番号フォルダ名 = 部品番号
        End If
        ws.Range("NC_品番番号フォルダ名").Value = NC品番番号フォルダ名

        Dim 品番フォルダパス As String
        品番フォルダパス = フォルダパス & "\" & NC品番番号フォルダ名

        If Dir(品番フォルダパス, vbDirectory) = "" Then
            MkDir 品番フォルダパス
        End If
    End If

    Call ショートカットの作成("NC")
End Sub

Private Sub スケジュール記録追記()
    ' プロジェクトのスケジュールシートのデータ範囲テーブルに1行追記
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("スケジュール")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set tbl = ws.ListObjects(1)
    On Error GoTo 0

    If tbl Is Nothing Then Exit Sub

    Set newRow = tbl.ListRows.Add

    On Error Resume Next
    newRow.Range(tbl.ListColumns("登録日").Index).Value = Date
    newRow.Range(tbl.ListColumns("コード番号").Index).Value = Range("コード番").Value
    newRow.Range(tbl.ListColumns("図番").Index).Value = Range("図番").Value
    newRow.Range(tbl.ListColumns("指示番号").Index).Value = Range("指示番号").Value
    newRow.Range(tbl.ListColumns("品名").Index).Value = Range("製品名").Value
    newRow.Range(tbl.ListColumns("部品番号").Index).Value = Range("部品番号").Value
    newRow.Range(tbl.ListColumns("部品名").Index).Value = Range("部品名").Value
    On Error GoTo 0
End Sub
