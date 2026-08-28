VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)

    Debug.Print "[加工担当者/機械連動] PivotTableUpdate発火 Target.Name=" & Target.Name

    If Target.Name <> "ピボットテーブル1" Then Exit Sub

    Call スライサー連動反映(Target, "加工担当者")
    Call スライサー連動反映(Target, "加工機械")

End Sub

Private Sub スライサー連動反映(ByVal Target As PivotTable, ByVal フィールド名 As String)

    On Error GoTo エラー処理

    Debug.Print "[スライサー連動] START フィールド名=" & フィールド名

    Dim pf As PivotField
    Set pf = Target.PivotFields(フィールド名)

    Dim tbl As ListObject
    Set tbl = Worksheets("プロジェクトのスケジュール").ListObjects("データ範囲")

    Dim n As Long, idx As Long
    Dim 全件数 As Long
    Dim itemName As String
    Dim itemVisible As Boolean
    Dim names(1 To 500) As String
    n = 0
    全件数 = 0

    For idx = 1 To pf.PivotItems.Count
        itemName = ""
        itemVisible = False
        On Error Resume Next
        Err.Clear
        itemName = CStr(pf.PivotItems(idx).Name)
        itemVisible = CBool(pf.PivotItems(idx).Visible)
        On Error GoTo エラー処理

        If itemName <> "(blank)" Then
            全件数 = 全件数 + 1
            If itemVisible Then
                n = n + 1
                names(n) = itemName
            End If
        End If
    Next idx

    Debug.Print "[スライサー連動] " & フィールド名 & " 可視件数=" & n & " 全件数(blank除く)=" & 全件数

    Dim 元EnableEvents As Boolean
    元EnableEvents = Application.EnableEvents
    Application.EnableEvents = False

    If n = 0 Or n = 全件数 Then
        Debug.Print "[スライサー連動] " & フィールド名 & " 全選択/全解除 → フィルタ解除"
        On Error Resume Next
        tbl.Range.AutoFilter Field:=tbl.ListColumns(フィールド名).Index
        On Error GoTo エラー処理
    Else
        Dim arr() As String
        ReDim arr(1 To n)
        Dim i As Long
        For i = 1 To n
            arr(i) = names(i)
        Next i
        Debug.Print "[スライサー連動] " & フィールド名 & " arr(1)=" & arr(1)
        tbl.Range.AutoFilter Field:=tbl.ListColumns(フィールド名).Index, _
            Criteria1:=arr, Operator:=xlFilterValues
        Debug.Print "[スライサー連動] " & フィールド名 & " AutoFilter適用OK FilterMode=" & tbl.AutoFilter.FilterMode
    End If

    Application.EnableEvents = 元EnableEvents

    Exit Sub

エラー処理:
    Application.EnableEvents = True
    Debug.Print "[ERROR][スライサー連動]" & フィールド名, Err.Number, Err.Description

End Sub
