VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInput
   Caption         =   "入力フォーム"
   ClientHeight    =   4800
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = False
    ' 既存値を名前付き範囲から読み込む
    On Error Resume Next
    txtCode.Text = Range("コード番").Value
    txtZuban.Text = Range("図番").Value
    txtShiji.Text = Range("指示番号").Value
    txtSeihin.Text = Range("製品名").Value
    txtBuhinNo.Text = Range("部品番号").Value
    txtBuhinmei.Text = Range("部品名").Value
    On Error GoTo 0
End Sub

Private Sub btnOK_Click()
    ' バリデーション：図番または指示番号のどちらかは必須
    If Trim(txtZuban.Text) = "" And Trim(txtShiji.Text) = "" Then
        MsgBox "図番または指示番号のどちらかを入力してください。", vbExclamation
        txtZuban.SetFocus
        Exit Sub
    End If

    ' 名前付き範囲へ書き込み
    On Error Resume Next
    Range("コード番").Value = Trim(txtCode.Text)
    Range("図番").Value = Trim(txtZuban.Text)
    Range("指示番号").Value = Trim(txtShiji.Text)
    Range("製品名").Value = Trim(txtSeihin.Text)
    Range("部品番号").Value = Trim(txtBuhinNo.Text)
    Range("部品名").Value = Trim(txtBuhinmei.Text)
    On Error GoTo 0

    Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Cancelled = True
        Me.Hide
    End If
End Sub
