Attribute VB_Name = "Module2_列表示非表示"
Private Const テーブル名 As String = "データ範囲"

Sub 列表示()
'
'
    Dim tbl As ListObject
    Set tbl = テーブル取得(ActiveSheet)
    If tbl Is Nothing Then
        Debug.Print "[エラー] テーブル「" & テーブル名 & "」が見つかりません。シート:", ActiveSheet.Name
        MsgBox "テーブル「" & テーブル名 & "」が見つかりません。処理を中断します。", vbExclamation
        Exit Sub
    End If

    '=== 追加：画面状態保存 ===
    Dim 保存スクロール行 As Long
    Dim 保存スクロール列 As Long
    Dim 保存アクティブセル As Range
    Call 画面状態保存(保存スクロール行, 保存スクロール列, 保存アクティブセル)

 ' 列表示 Macro
    Application.ScreenUpdating = False
    ActiveWindow.FreezePanes = False
    Columns("A:AL").Select
    Range("A4").Activate
    Selection.EntireColumn.Hidden = False

    Call 固定位置設定(tbl, 列も固定:=False)

    Application.ScreenUpdating = True

    '=== 追加：画面状態復元 ===
    Call 画面状態復元(保存スクロール行, 保存スクロール列, 保存アクティブセル, 中央寄せ:=True)

End Sub

Sub 列非表示()
'
'
    Dim tbl As ListObject
    Set tbl = テーブル取得(ActiveSheet)
    If tbl Is Nothing Then
        Debug.Print "[エラー] テーブル「" & テーブル名 & "」が見つかりません。シート:", ActiveSheet.Name
        MsgBox "テーブル「" & テーブル名 & "」が見つかりません。処理を中断します。", vbExclamation
        Exit Sub
    End If

    '=== 追加：画面状態保存 ===
    Dim 保存スクロール行 As Long
    Dim 保存スクロール列 As Long
    Dim 保存アクティブセル As Range
    Call 画面状態保存(保存スクロール行, 保存スクロール列, 保存アクティブセル)

 ' 列非表示 Macro
    Application.ScreenUpdating = False
    ActiveWindow.FreezePanes = False
    Range("B:B,D:O,Q:Q,T:T,Y:AA,AE:AE,AJ:AL").Select
    Selection.EntireColumn.Hidden = True

    Call 固定位置設定(tbl, 列も固定:=True)

    Application.ScreenUpdating = True

   '=== 追加：画面状態復元 ===
    Call 画面状態復元(保存スクロール行, 保存スクロール列, 保存アクティブセル, 中央寄せ:=False)

End Sub

'=== 共通：テーブル「データ範囲」を取得（見つからない場合はNothing） ===
Private Function テーブル取得(ByVal 対象シート As Worksheet) As ListObject
    On Error Resume Next
    Set テーブル取得 = 対象シート.ListObjects(テーブル名)
    On Error GoTo 0
End Function

'=== 共通：ウィンドウ枠の固定位置を設定 ===
'   固定行 = テーブルの見出し行の行番号（見出し行まで固定）
'   固定列 = テーブルの右端列の列番号（テーブル全体を固定、列も固定:=Falseの場合は固定しない）
'
'   ExcelはSplitRow/SplitColumnを設定する際、対象範囲に非表示の行・列が
'   挟まっていると「実際の行番号・列番号」ではなく「非表示を除いた見た目上の
'   カウント」で解釈してしまう（セルを選択してからFreezePanesを実行する方式でも
'   同様の現象が発生することを実機で確認済み）。
'   これを避けるため、固定範囲（1行目〜見出し行+1行目、1列目〜右端列+1列目）に
'   含まれる非表示行・非表示列を設定直前だけ一時的に表示状態へ戻し、
'   SplitRow/SplitColumnを数値で設定してから、元の非表示状態へ復元する。
Private Sub 固定位置設定(ByVal tbl As ListObject, Optional ByVal 列も固定 As Boolean = True)
    Dim 固定行 As Long
    Dim 固定列 As Long
    固定行 = tbl.HeaderRowRange.Row
    固定列 = 0
    If 列も固定 Then
        固定列 = tbl.Range.Column + tbl.Range.Columns.Count - 1
    End If

    Debug.Print "[固定位置設定] テーブル:", tbl.Name, "見出し行:", 固定行, "右端列:", 固定列, "列も固定:", 列も固定

    ' 固定範囲にある非表示行・非表示列を記憶して一時的に表示
    Dim 非表示列 As Range
    Dim 非表示行 As Range
    Dim i As Long

    If 列も固定 Then
        For i = 1 To 固定列 + 1
            If Columns(i).Hidden Then
                If 非表示列 Is Nothing Then
                    Set 非表示列 = Columns(i)
                Else
                    Set 非表示列 = Union(非表示列, Columns(i))
                End If
            End If
        Next i
    End If

    For i = 1 To 固定行 + 1
        If Rows(i).Hidden Then
            If 非表示行 Is Nothing Then
                Set 非表示行 = Rows(i)
            Else
                Set 非表示行 = Union(非表示行, Rows(i))
            End If
        End If
    Next i

    Dim 非表示列数 As Long
    Dim 非表示行数 As Long
    If 非表示列 Is Nothing Then 非表示列数 = 0 Else 非表示列数 = 非表示列.Areas.Count
    If 非表示行 Is Nothing Then 非表示行数 = 0 Else 非表示行数 = 非表示行.Areas.Count

    Debug.Print "[固定位置設定] 一時解除する非表示列数:", 非表示列数, "一時解除する非表示行数:", 非表示行数

    If Not 非表示列 Is Nothing Then 非表示列.EntireColumn.Hidden = False
    If Not 非表示行 Is Nothing Then 非表示行.EntireRow.Hidden = False

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SplitRow = 固定行
    ActiveWindow.SplitColumn = 固定列
    ActiveWindow.FreezePanes = True

    Debug.Print "[固定位置設定] 設定後 SplitRow:", ActiveWindow.SplitRow, _
                "SplitColumn:", ActiveWindow.SplitColumn, "FreezePanes:", ActiveWindow.FreezePanes

    ' 非表示状態を復元
    If Not 非表示列 Is Nothing Then 非表示列.EntireColumn.Hidden = True
    If Not 非表示行 Is Nothing Then 非表示行.EntireRow.Hidden = True
End Sub

'=== 共通：画面状態保存 ===
Private Sub 画面状態保存(ByRef 行 As Long, ByRef 列 As Long, ByRef セル As Range)
    Set セル = ActiveCell
    行 = ActiveWindow.ScrollRow
    列 = ActiveWindow.ScrollColumn

    Debug.Print "[画面保存] Row:", 行, "Col:", 列
End Sub

'=== 共通：画面状態復元 ===
'   中央寄せ:=True  … アクティブセルが画面中央に来るようスクロール列を再計算（列表示用）
'   中央寄せ:=False … 保存しておいたスクロール列をそのまま復元（列非表示用）
Private Sub 画面状態復元(ByVal 行 As Long, ByVal 列 As Long, ByVal セル As Range, Optional ByVal 中央寄せ As Boolean = False)
    Application.ScreenUpdating = False

    If 中央寄せ Then
        ActiveWindow.ScrollRow = 行
        Dim 表示列数 As Long
        表示列数 = ActiveWindow.VisibleRange.Columns.Count
        ActiveWindow.ScrollColumn = Application.Max(1, セル.Column - 表示列数 ¥ 2)
        セル.Select
    Else
        セル.Select
        ActiveWindow.ScrollRow = 行
        ActiveWindow.ScrollColumn = 列
    End If

    Application.ScreenUpdating = True
    DoEvents

    Debug.Print "[画面復元] 完了"
End Sub


' アクティブなシートに存在するすべてのテーブル(ListObject)の
' オートフィルタと、そのテーブルをデータソースとするピボットテーブルの
' スライサー選択状態をあわせて解除するマクロ
Sub フィルタクリア()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim clearedCount As Long

    Set ws = ActiveSheet
    clearedCount = 0

    If ws.ListObjects.Count = 0 Then
        MsgBox "アクティブシートにテーブルが存在しません。", vbInformation
        Exit Sub
    End If

    For Each tbl In ws.ListObjects
        If tbl.ShowAutoFilter Then
            If Not tbl.AutoFilter Is Nothing Then
                If tbl.AutoFilter.FilterMode Then
                    tbl.AutoFilter.ShowAllData
                    clearedCount = clearedCount + 1
                End If
            End If
        End If

        Call テーブル連動スライサークリア(tbl)
    Next tbl

    'MsgBox clearedCount & " 個のテーブルのフィルタを解除しました。", vbInformation
End Sub

'=== 共通：指定テーブルをデータソースとするピボットテーブルに紐づく
'          スライサーの選択状態を全解除（全アイテム選択に戻す） ===
Private Sub テーブル連動スライサークリア(ByVal tbl As ListObject)
    Dim ws2 As Worksheet
    Dim pt As PivotTable
    Dim sc As SlicerCache
    Dim scPt As PivotTable
    Dim isLinked As Boolean

    For Each ws2 In ThisWorkbook.Worksheets
        For Each pt In ws2.PivotTables
            isLinked = False
            On Error Resume Next
            isLinked = (pt.SourceData = tbl.Name)
            On Error GoTo 0

            If isLinked Then
                For Each sc In ThisWorkbook.SlicerCaches
                    For Each scPt In sc.PivotTables
                        If scPt.Name = pt.Name And scPt.Parent.Name = pt.Parent.Name Then
                            On Error Resume Next
                            sc.ClearManualFilter
                            On Error GoTo 0
                            Debug.Print "[フィルタクリア] スライサー解除:", sc.Name, "← ピボット:", pt.Name, "(" & pt.Parent.Name & ")"
                            Exit For
                        End If
                    Next scPt
                Next sc
            End If
        Next pt
    Next ws2
End Sub
