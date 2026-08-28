Attribute VB_Name = "Module8_クイック調整ボタン"
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
' CAM/加工バー クイック調整ボタン
'
' Module3（ガントチャート描画）から分離した、ガント上のセルをクリックして
' 開始日・所要日数を±1/±0.1ずらすためのボタン用エントリーポイント一式。
' 実際の行再描画（単一行再描画）・祝日判定・完了日時計算はModule3側の
' Public関数を呼び出す（m_arrCalendarもModule3側のPublic変数を参照）。
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
Option Explicit

'================================================================
' CAM/加工バー クイック調整ボタン（開始日・所要日数）
' ボタンに割り当てて使用する8個のエントリーポイント
'================================================================
Sub 開始日調整_プラス1()
    Call 開始日所要日数調整("開始日", 1)
End Sub

Sub 開始日調整_プラス0_1()
    Call 開始日所要日数調整("開始日", 0.1)
End Sub

Sub 開始日調整_マイナス0_1()
    Call 開始日所要日数調整("開始日", -0.1)
End Sub

Sub 開始日調整_マイナス1()
    Call 開始日所要日数調整("開始日", -1)
End Sub

Sub 所要日数調整_プラス1()
    Call 開始日所要日数調整("所要日数", 1)
End Sub

Sub 所要日数調整_プラス0_1()
    Call 開始日所要日数調整("所要日数", 0.1)
End Sub

Sub 所要日数調整_マイナス0_1()
    Call 開始日所要日数調整("所要日数", -0.1)
End Sub

Sub 所要日数調整_マイナス1()
    Call 開始日所要日数調整("所要日数", -1)
End Sub

'================================================================
' 選択セル→行番号・CAM/加工の判定
'================================================================
Private Function 対象行と種別を判定(ByVal セル As Range, ByRef 種別 As String, ByRef 行番号 As Long) As Boolean

    対象行と種別を判定 = False
    種別 = ""
    行番号 = 0


    On Error GoTo エラー処理

    Dim データ範囲 As ListObject
    Dim rng_カレンダー As Range
    Set データ範囲 = セル.Worksheet.ListObjects("データ範囲")
    Set rng_カレンダー = セル.Worksheet.Range("カレンダー範囲")

    If データ範囲.DataBodyRange Is Nothing Then
        Exit Function
    End If

    '▼行番号算出
    Dim 行番号候補 As Long
    行番号候補 = セル.Row - データ範囲.DataBodyRange.Row + 1
    If 行番号候補 < 1 Or 行番号候補 > データ範囲.DataBodyRange.Rows.Count Then
        Exit Function
    End If

    '▼列→カレンダー日付変換
    Dim 列オフセット As Long
    列オフセット = セル.Column - rng_カレンダー.Column + 1
    If 列オフセット < 1 Or 列オフセット > rng_カレンダー.Columns.Count Then
        Exit Function
    End If

    Dim 対象日 As Double
    対象日 = CDbl(rng_カレンダー.Cells(1, 列オフセット).Value)

    '▼CAM/加工の日程範囲取得（対象行のみ）
    Dim rng_行 As Range
    Set rng_行 = データ範囲.DataBodyRange.Rows(行番号候補)

    Dim CAM開始 As Double, CAM終了 As Double
    Dim 加工開始 As Double, 加工終了 As Double
    CAM開始 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns("CAM開始日時").Index).Value) Then CAM開始 = CDbl(rng_行.Cells(1, データ範囲.ListColumns("CAM開始日時").Index).Value)
    CAM終了 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns("CAM終了日時").Index).Value) Then CAM終了 = CDbl(rng_行.Cells(1, データ範囲.ListColumns("CAM終了日時").Index).Value)
    加工開始 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns("加工開始日時").Index).Value) Then 加工開始 = CDbl(rng_行.Cells(1, データ範囲.ListColumns("加工開始日時").Index).Value)
    加工終了 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns("加工終了日時").Index).Value) Then 加工終了 = CDbl(rng_行.Cells(1, データ範囲.ListColumns("加工終了日時").Index).Value)


    '▼CAM優先で範囲判定（共通塗り分け_高速版と同じ許容誤差）
    If CAM開始 > 0 And CAM終了 > 0 Then
        If 対象日 + 0.005 >= CAM開始 And 対象日 < CAM終了 - 0.005 Then
            種別 = "CAM"
            行番号 = 行番号候補
            対象行と種別を判定 = True
            Exit Function
        End If
    End If

    If 加工開始 > 0 And 加工終了 > 0 Then
        If 対象日 + 0.005 >= 加工開始 And 対象日 < 加工終了 - 0.005 Then
            種別 = "加工"
            行番号 = 行番号候補
            対象行と種別を判定 = True
            Exit Function
        End If
    End If

    Exit Function

エラー処理:
    Debug.Print "[ERROR][対象行と種別を判定]", Err.Number, Err.Description
    対象行と種別を判定 = False

End Function

'================================================================
' 開始日／所要日数 調整エンジン
'================================================================
Private Sub 開始日所要日数調整(ByVal 種別対象 As String, ByVal デルタ As Double)

    Application.ScreenUpdating = False
    On Error GoTo エラー処理

    Call 祝日辞書初期化   ' 未構築なら構築、構築済みならスキップ

    '▼[追加] 複数行選択に対応：選択範囲内の各セルから対象行/種別を判定し、
    ' 同じ行を重複処理しないよう行番号単位でDictionary管理する
    Dim dic処理済み行 As Object
    Set dic処理済み行 = CreateObject("Scripting.Dictionary")

    Dim areaRng As Range
    Dim c As Range
    Dim 種別 As String
    Dim 行番号 As Long

    For Each areaRng In Selection.Areas
        For Each c In areaRng.Cells
            If Not dic処理済み行.Exists(c.Row) Then
                If 対象行と種別を判定(c, 種別, 行番号) Then
                    dic処理済み行.Add c.Row, True
                    Call 単一行調整実行(種別対象, デルタ, 種別, 行番号)
                End If
            End If
        Next c
    Next areaRng

    GoTo 終了処理

終了処理:
    Application.ScreenUpdating = True
    Exit Sub

エラー処理:
    Application.ScreenUpdating = True
    Debug.Print "[ERROR][開始日所要日数調整]", Err.Number, Err.Description

End Sub

'================================================================
' 対象行1件分の開始日/所要日数調整の実処理（開始日所要日数調整から行ごとに呼ばれる）
'================================================================
Private Sub 単一行調整実行(ByVal 種別対象 As String, ByVal デルタ As Double, ByVal 種別 As String, ByVal 行番号 As Long)

    On Error GoTo エラー処理

    Dim データ範囲 As ListObject
    Set データ範囲 = Worksheets("プロジェクトのスケジュール").ListObjects("データ範囲")

    Dim rng_行 As Range
    Set rng_行 = データ範囲.DataBodyRange.Rows(行番号)

    Dim col開始日 As String, col開始時刻 As String, col所要日数 As String
    Dim col開始日時 As String, col終了日時 As String, col進捗日時 As String
    Dim col進捗率 As String, col例外 As String

    If 種別 = "CAM" Then
        col開始日 = "CAM開始日": col開始時刻 = "CAM開始時刻": col所要日数 = "CAM所要日数"
        col開始日時 = "CAM開始日時": col終了日時 = "CAM終了日時": col進捗日時 = "CAM進捗日時"
        col進捗率 = "CAM進捗率": col例外 = "例外_CAM"
    Else
        col開始日 = "加工開始日": col開始時刻 = "加工開始時刻": col所要日数 = "加工所要日数"
        col開始日時 = "加工開始日時": col終了日時 = "加工終了日時": col進捗日時 = "加工進捗日時"
        col進捗率 = "加工進捗率": col例外 = "例外_加工"
    End If

    Dim rng_休日一覧 As Range
    On Error Resume Next
    Set rng_休日一覧 = Range("休日一覧")
    On Error GoTo 0

    If 種別対象 = "開始日" Then

        Dim 開始日セル As Range, 開始時刻セル As Range
        Set 開始日セル = rng_行.Cells(1, データ範囲.ListColumns(col開始日).Index)
        Set 開始時刻セル = rng_行.Cells(1, データ範囲.ListColumns(col開始時刻).Index)

        '▼[修正] IsNumeric(Empty)はTrueを返すVBAの仕様があるため、
        ' 空欄セルを「有効な0」と誤認しないようIsEmptyも合わせてチェックする。
        ' さらに「2026/07/21」のようにテキスト形式で日付が入力されているケース
        ' （休日一覧で判明したのと同じパターン）にも対応するためIsDate/CDateもフォールバックする。
        Dim 開始日値 As Double, 開始時刻値 As Double
        開始日値 = 0
        If Not IsEmpty(開始日セル.Value) Then
            If IsNumeric(開始日セル.Value) Then
                開始日値 = CDbl(開始日セル.Value)
            ElseIf IsDate(開始日セル.Value) Then
                開始日値 = CDbl(CDate(開始日セル.Value))
            End If
        End If
        開始時刻値 = 0
        If Not IsEmpty(開始時刻セル.Value) Then
            If IsNumeric(開始時刻セル.Value) Then
                開始時刻値 = CDbl(開始時刻セル.Value)
            ElseIf IsDate(開始時刻セル.Value) Then
                開始時刻値 = CDbl(CDate(開始時刻セル.Value))
            End If
        End If


        '▼[修正] 開始日が未入力（0以下）の行は調整対象として不正なので、何もせず静かに中断する
        If 開始日値 <= 0 Then
            Exit Sub
        End If

        If Abs(デルタ) = 1 Then

            開始日値 = 開始日値 + デルタ

        Else

            '▼キャリー／ボロー処理（Round(...,4)で浮動小数点誤差を吸収）
            開始時刻値 = Round(開始時刻値 + デルタ, 4)
            If 開始時刻値 >= 1 Then
                開始時刻値 = Round(開始時刻値 - 1, 4)
                開始日値 = 開始日値 + 1
            ElseIf 開始時刻値 < 0 Then
                開始時刻値 = Round(開始時刻値 + 1, 4)
                開始日値 = 開始日値 - 1
            End If

        End If


        開始日セル.Value = 開始日値
        開始時刻セル.Value = 開始時刻値

        Dim 新開始日時 As Double
        新開始日時 = Round(開始日値 + 開始時刻値, 4)
        rng_行.Cells(1, データ範囲.ListColumns(col開始日時).Index).Value = 新開始日時


        '▼[レビュー指摘] 土日祝日への例外自動追記（開始日土日休日例外設定の該当行分を複製）
        Dim 例外セル As Range
        Set 例外セル = rng_行.Cells(1, データ範囲.ListColumns(col例外).Index)

        If 新開始日時 > 0 Then
            If Not 営業日であるか(CDbl(Int(新開始日時)), rng_休日一覧) Then
                Dim lDate As Long, sDate As String, sEntry As String, sExc As String
                lDate = Int(新開始日時)
                sDate = Format(CDate(lDate), "yyyy/m/d")
                sEntry = sDate & ">0+1"
                sExc = CStr(例外セル.Value)
                If InStr(sExc, sDate) = 0 Then
                    If sExc = "" Then
                        例外セル.Value = sEntry
                    Else
                        例外セル.Value = sExc & "," & sEntry
                    End If
                End If
            End If
        End If

    Else   ' 種別対象 = "所要日数"

        Dim 所要日数セル As Range
        Set 所要日数セル = rng_行.Cells(1, データ範囲.ListColumns(col所要日数).Index)

        Dim 所要日数値 As Double
        所要日数値 = 0: If IsNumeric(所要日数セル.Value) Then 所要日数値 = CDbl(所要日数セル.Value)


        所要日数値 = Round(所要日数値 + デルタ, 4)
        If 所要日数値 < 0.1 Then
            所要日数値 = 0.1
        End If

        所要日数セル.Value = 所要日数値


    End If

    '▼終了日時・進捗日時の再計算（Get完了日時を利用、例外文字列は上記反映後の最新値で再構築）
    Dim 例外セル参照 As Range
    Set 例外セル参照 = rng_行.Cells(1, データ範囲.ListColumns(col例外).Index)

    Dim 例外文字列 As String
    例外文字列 = CStr(例外セル参照.Value)

    Dim dic_例外 As Object
    Set dic_例外 = 例外データ構築(例外文字列, 例外セル参照)

    Dim flg_例外なし As Boolean
    flg_例外なし = (dic_例外 Is Nothing)

    Dim 現開始日時 As Double, 現所要日数 As Double, 現進捗率 As Double
    現開始日時 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns(col開始日時).Index).Value) Then 現開始日時 = CDbl(rng_行.Cells(1, データ範囲.ListColumns(col開始日時).Index).Value)
    現所要日数 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns(col所要日数).Index).Value) Then 現所要日数 = CDbl(rng_行.Cells(1, データ範囲.ListColumns(col所要日数).Index).Value)
    現進捗率 = 0: If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns(col進捗率).Index).Value) Then 現進捗率 = CDbl(rng_行.Cells(1, データ範囲.ListColumns(col進捗率).Index).Value)


    If 現開始日時 > 0 And 現所要日数 > 0 Then
        Dim 新終了日時 As Double, 新進捗日時 As Double
        新終了日時 = Get完了日時(現開始日時, 現所要日数, rng_休日一覧, dic_例外, flg_例外なし)
        新進捗日時 = Get完了日時(現開始日時, 現所要日数 * 現進捗率, rng_休日一覧, dic_例外, flg_例外なし)
        rng_行.Cells(1, データ範囲.ListColumns(col終了日時).Index).Value = 新終了日時
        rng_行.Cells(1, データ範囲.ListColumns(col進捗日時).Index).Value = 新進捗日時
    Else
    End If


    Call 単一行再描画(行番号)

    '▼ボタン連打対応：ガントチャート上の（調整後の）開始日時セルへアクティブセルを追従させる
    Dim 現在開始日時 As Double
    現在開始日時 = 0
    If IsNumeric(rng_行.Cells(1, データ範囲.ListColumns(col開始日時).Index).Value) Then _
        現在開始日時 = CDbl(rng_行.Cells(1, データ範囲.ListColumns(col開始日時).Index).Value)


    If 現在開始日時 > 0 Then
        Dim rng_カレンダー2 As Range, rng_ガント As Range
        Set rng_カレンダー2 = Worksheets("プロジェクトのスケジュール").Range("カレンダー範囲")
        Set rng_ガント = Worksheets("プロジェクトのスケジュール").Range("ガントチャート範囲")

        '▼[修正] カレンダーは1日1列ではなく0.1日刻みの列で構成されているため、
        ' Int()で「その日の先頭列」にマッチさせるのではなく、値そのものが一致する列を探す
        Dim 列i As Long, 対象列 As Long
        Dim 対象列_シート直読み As Long
        対象列 = 0
        対象列_シート直読み = 0
        For 列i = 1 To UBound(m_arrCalendar, 2)
            If Abs(m_arrCalendar(1, 列i) - 現在開始日時) < 0.05 Then
                If 対象列 = 0 Then 対象列 = 列i
            End If
            If Abs(CDbl(rng_カレンダー2.Cells(1, 列i).Value) - 現在開始日時) < 0.05 Then
                If 対象列_シート直読み = 0 Then 対象列_シート直読み = 列i
            End If
        Next 列i


        If 対象列 > 0 Then
            Dim セル候補 As Range
            Set セル候補 = rng_ガント.Cells(行番号, 対象列)
            '▼[修正] ActiveXボタン等クリック後はフォーカスがコントロール側に残り、
            ' 単なる.Selectでは画面上アクティブセルが追従しないことがあるため、
            ' ワークシートへ確実にフォーカスを戻すApplication.Gotoを使う
            Application.Goto Reference:=セル候補, Scroll:=False
        Else
        End If
    End If

    Exit Sub

エラー処理:
    Debug.Print "[ERROR][単一行調整実行]", Err.Number, Err.Description

End Sub

