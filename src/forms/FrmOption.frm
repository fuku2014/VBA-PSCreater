VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmOption 
   Caption         =   "オプション"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9060
   OleObjectBlob   =   "FrmOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnOk_Click()
    Const ERR_MSG1 As String = "有効な値は１～９９の整数です。"
    '--数値チェックStart
    If Not IsNumeric(TxtModuleContentRow.Text) Or TxtModuleContentRow.Text = "" Then
        MultiPage1.Value = 0
        TxtModuleContentRow.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If Not IsNumeric(TxtModuleContentRow2.Text) Or TxtModuleContentRow2.Text = "" Then
        MultiPage1.Value = 0
        TxtModuleContentRow2.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If Not IsNumeric(TxtProcContentRow.Text) Or TxtProcContentRow.Text = "" Then
        MultiPage1.Value = 1
        TxtProcContentRow.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        
        Exit Sub
    End If
    If Not IsNumeric(TxtProcContentRow2.Text) Or TxtProcContentRow2.Text = "" Then
        TxtProcContentRow2.SetFocus
        MultiPage1.Value = 1
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    '--範囲チェックStart
    If TxtModuleContentRow.Text < 1 Or TxtModuleContentRow.Text > 99 Then
        MultiPage1.Value = 0
        TxtModuleContentRow.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If TxtModuleContentRow2.Text < 1 Or TxtModuleContentRow2.Text > 99 Then
        MultiPage1.Value = 0
        TxtModuleContentRow2.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        
        Exit Sub
    End If
    If TxtProcContentRow.Text < 1 Or TxtProcContentRow.Text > 99 Then
        MultiPage1.Value = 1
        TxtProcContentRow.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If TxtProcContentRow2.Text < 1 Or TxtProcContentRow2.Text > 99 Then
        MultiPage1.Value = 1
        TxtProcContentRow2.SetFocus
        MsgBox ERR_MSG1, vbCritical, ThisWorkbook.Name
        Exit Sub
    End If
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniWrite
    Unload Me
End Sub


Private Sub CheckExitModuleContent_Click()
    Frame2.Enabled = Not CheckExitModuleContent.Value
    Frame3.Enabled = Not CheckExitModuleContent.Value
    If Not Frame2.Enabled Then
        TxtModuleContentRow.Text = 1
        TxtModuleContentRow2.Text = 1
    End If
End Sub

Private Sub CheckExitProcContent_Click()
    Frame5.Enabled = Not CheckExitProcContent.Value
    Frame6.Enabled = Not CheckExitProcContent.Value
    If Not Frame6.Enabled Then
        TxtProcContentRow.Text = 1
        TxtProcContentRow2.Text = 1
    End If
End Sub

Private Sub CheckUseNow_Click()
    If CheckUseNow.Value Then
        TxtDate.Text = Date
    End If
End Sub

Private Sub CheckUseOSNm_Click()
    If CheckUseOSNm.Value Then
        TxtAuthor.Text = CreateObject("WScript.NetWork").UserName
    End If
End Sub

Private Sub TxtModuleContentRow_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtModuleContentRow2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtProcContentRow_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtProcContentRow2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtModuleContentRow_Change()
    Call TxtCheck(TxtModuleContentRow)
End Sub
Private Sub TxtModuleContentRow2_Change()
    Call TxtCheck(TxtModuleContentRow2)
End Sub



Private Sub TxtProcContentRow_Change()
    Call TxtCheck(TxtProcContentRow)
End Sub



Private Sub TxtProcContentRow2_Change()
    Call TxtCheck(TxtProcContentRow2)
End Sub

Private Sub TxtCheck(aTxtBox As MSForms.TextBox)
    If Len(aTxtBox.Text) = 0 Then
        Exit Sub
    End If
    If IsNumeric(Right(aTxtBox.Text, 1)) = True Then
        Exit Sub
    End If
    aTxtBox.Text = Left(aTxtBox.Text, Len(aTxtBox.Text) - 1)
End Sub

Private Sub UserForm_Initialize()
    If Not IsExistsIni Then
        Call CreateIniFile
    End If
    Call IniRead
    With cIniKeyList
        TxtModuleContentRow.Text = CStr(.aModuleContentRow)
        TxtModuleContentRow2.Text = CStr(.aModuleContentRow2)
        TxtModuleRemComment.Text = .aModuleRemComment
        CheckExitModuleContent.Value = .aModuleContentNotExist
        TxtProcContentRow.Text = CStr(.aProcContentRow)
        TxtProcContentRow2.Text = CStr(.aProcContentRow2)
        OptProcRow.Value = (.aProcOptWhere = OptModuleType.aOptRow)
        OptProcComment = (.aProcOptWhere = OptModuleType.aOptCom)
        TxtProcRemComment.Text = .aProcRemComment
        CheckExitProcContent.Value = .aProcContentNotExist
        TxtProcContentComment.Text = .aProcContent
        CheckNormal.Value = .aNormalSelect
        CheckSh.Value = .aSheetSelect
        CheckFrm.Value = .aFrmSelect
        CheckCls.Value = .aClsSelect
        CheckUseOSNm.Value = .aAcnSelect
        CheckUseNow.Value = .aNowSelect
        TxtAuthor.Text = .aAutName
        TxtDate.Text = .aCreDate
    End With
    Frame2.Enabled = Not CheckExitModuleContent.Value
    Frame3.Enabled = Not CheckExitModuleContent.Value
    Frame5.Enabled = Not CheckExitProcContent.Value
    Frame6.Enabled = Not CheckExitProcContent.Value
End Sub


