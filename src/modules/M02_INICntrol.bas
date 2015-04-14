Attribute VB_Name = "M02_INICntrol"
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" _
(ByVal AppName As String, _
ByVal KeyName As String, _
ByVal Default As String, _
ByVal ReturnedString As String, _
ByVal MaxSize As Long, _
ByVal FileName As String) As Long

Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" _
(ByVal AppName As String, _
ByVal KeyName As Any, _
ByVal lpString As Any, _
ByVal FileName As String) As Long


Public Const INI_NAME                      As String = "PSDoc.Ini"
Public Const INI_SEC_OPT_MODULE            As String = "ModuleOption"
Public Const INI_KEY_MODULE_CONTENT_ROW    As String = "Module_Content_Row"
Public Const INI_KEY_MODULE_CONTENT_ROW2   As String = "Module_Content_Row2"
Public Const INI_KEY_MODULE_REM_COMMENT    As String = "Module_Rem_Comment"
Public Const INI_KEY_MODULE_CONTENT_EXIST  As String = "Module_Content_Exist"
Public Const INI_SEC_OPT_PROC              As String = "ProcOption"
Public Const INI_KEY_PROC_CONTENT_ROW      As String = "Proc_Content_Row"
Public Const INI_KEY_PROC_CONTENT_ROW2     As String = "Proc_Content_Row2"
Public Const INI_KEY_PROC_OPT_WHERE        As String = "Proc_Opt_Where"
Public Const INI_KEY_PROC_CONTENT          As String = "Proc_Content"
Public Const INI_KEY_PROC_REM_COMMENT      As String = "Proc_Rem_Comment"
Public Const INI_KEY_PROC_CONTENT_EXIST    As String = "Proc_Content_Exist"
Public Const INI_SEC_OPT_EDIT              As String = "EditOption"
Public Const INI_KEY_EDIT_NORMAL_SELECT    As String = "Edit_Normal_Select"
Public Const INI_KEY_EDIT_SHEET_SELECT     As String = "Edit_Sheet_Select"
Public Const INI_KEY_EDIT_FRM_SELECT       As String = "Edit_Frm_Select"
Public Const INI_KEY_EDIT_CLS_SELECT       As String = "Edit_Cls_Select"
Public Const INI_KEY_EDIT_ACN_SELECT       As String = "Edit_Acn_Select"
Public Const INI_KEY_EDIT_NOW_SELECT       As String = "Edit_Now_Select"
Public Const INI_KEY_EDIT_AUT_NAME         As String = "Edit_Aut_Name"
Public Const INI_KEY_EDIT_CRE_DATE         As String = "Edit_Cre_Date"

Public Enum OptModuleType
    aOptRow = 0
    aOptCom = 1
End Enum

Public Type INI_KEY_LIST
    aModuleContentRow        As Integer
    aModuleContentRow2       As Integer
    aModuleRemComment        As String
    aModuleContentNotExist   As Boolean
    aProcContentRow          As Integer
    aProcContentRow2         As Integer
    aProcOptWhere            As Integer
    aProcRemComment          As String
    aProcContentNotExist     As Boolean
    aProcContent             As String
    aNormalSelect            As Boolean
    aSheetSelect             As Boolean
    aFrmSelect               As Boolean
    aClsSelect               As Boolean
    aAcnSelect               As Boolean
    aNowSelect               As Boolean
    aAutName                 As String
    aCreDate                 As String
End Type



Public cIniKeyList As INI_KEY_LIST
'*********************************************************
'名称：INIファイルの値取得
'---------------------------------------------------------
'引き数：aIniKey     :キー
'        aIniSection :セクション
'---------------------------------------------------------
'戻り値：取得値
'*********************************************************
Public Function GetIniValue(aIniKey As String, aIniSection As String) As String
    Dim wIniVal    As String * 1024
    Dim wRet       As Long
    wRet = GetPrivateProfileString(aIniSection, aIniKey, "", wIniVal, Len(wIniVal), ThisWorkbook.Path & "\" & INI_NAME)
    GetIniValue = Left(wIniVal, InStr(wIniVal, vbNullChar) - 1)
End Function

'*********************************************************
'名称：INIファイルの値設定
'---------------------------------------------------------
'引き数：aIniKey     :キー
'        aIniSection :セクション
'        aValue      :設定値
'---------------------------------------------------------
'戻り値：取得値
'*********************************************************
Public Sub SetIniValue(aIniKey As String, aIniSection As String, aValue As String)
    Dim wRet     As Long
    wRet = WritePrivateProfileString(aIniSection, aIniKey, aValue, ThisWorkbook.Path & "\" & INI_NAME)
End Sub

'*********************************************************
'名称：INIファイル存在判定
'---------------------------------------------------------
'引き数：None
'---------------------------------------------------------
'戻り値：存在成否
'*********************************************************
Public Function IsExistsIni() As Boolean
    IsExistsIni = Dir(ThisWorkbook.Path & "\" & INI_NAME) <> ""
End Function
'*********************************************************
'名称：INIファイル作成
'---------------------------------------------------------
'引き数：None
'---------------------------------------------------------
'戻り値：None
'*********************************************************
Public Sub CreateIniFile()
    Dim wNo As Integer
    wNo = FreeFile
    Open ThisWorkbook.Path & "\" & INI_NAME For Output As #wNo
    Print #wNo, "[Info]"
    Print #wNo, "  This file is used by PSDocToolAddIn"
    Print #wNo, "[" & INI_SEC_OPT_MODULE & "]"
    Print #wNo, vbTab & INI_KEY_MODULE_CONTENT_ROW & "=1"
    Print #wNo, vbTab & INI_KEY_MODULE_CONTENT_ROW2 & "=1"
    Print #wNo, vbTab & INI_KEY_MODULE_REM_COMMENT & "='"
    Print #wNo, vbTab & INI_KEY_MODULE_CONTENT_EXIST & "=False"
    Print #wNo, "[" & INI_SEC_OPT_PROC & "]"
    Print #wNo, vbTab & INI_KEY_PROC_CONTENT_ROW & "=1"
    Print #wNo, vbTab & INI_KEY_PROC_CONTENT_ROW2 & "=1"
    Print #wNo, vbTab & INI_KEY_PROC_OPT_WHERE & "=" & OptModuleType.aOptRow
    Print #wNo, vbTab & INI_KEY_PROC_REM_COMMENT & "=  '"
    Print #wNo, vbTab & INI_KEY_PROC_CONTENT_EXIST & "=False"
    Print #wNo, vbTab & INI_KEY_PROC_CONTENT & "=      '"
    Print #wNo, "[" & INI_SEC_OPT_EDIT & "]"
    Print #wNo, vbTab & INI_KEY_EDIT_NORMAL_SELECT & "=True"
    Print #wNo, vbTab & INI_KEY_EDIT_SHEET_SELECT & "=True"
    Print #wNo, vbTab & INI_KEY_EDIT_FRM_SELECT & "=True"
    Print #wNo, vbTab & INI_KEY_EDIT_CLS_SELECT & "=True"
    Print #wNo, vbTab & INI_KEY_EDIT_ACN_SELECT & "=True"
    Print #wNo, vbTab & INI_KEY_EDIT_NOW_SELECT & "=True"
    Print #wNo, vbTab & INI_KEY_EDIT_AUT_NAME & "="
    Print #wNo, vbTab & INI_KEY_EDIT_CRE_DATE & "="
    Close #wNo
End Sub

Public Sub IniWrite()
    Call SetIniValue(INI_KEY_MODULE_CONTENT_ROW, INI_SEC_OPT_MODULE, FrmOption.TxtModuleContentRow.Text)
    Call SetIniValue(INI_KEY_MODULE_CONTENT_ROW2, INI_SEC_OPT_MODULE, FrmOption.TxtModuleContentRow2.Text)
    Call SetIniValue(INI_KEY_MODULE_REM_COMMENT, INI_SEC_OPT_MODULE, FrmOption.TxtModuleRemComment.Text)
    Call SetIniValue(INI_KEY_MODULE_CONTENT_EXIST, INI_SEC_OPT_MODULE, FrmOption.CheckExitModuleContent.Value)
    Call SetIniValue(INI_KEY_PROC_CONTENT_ROW, INI_SEC_OPT_PROC, FrmOption.TxtProcContentRow.Text)
    Call SetIniValue(INI_KEY_PROC_CONTENT_ROW2, INI_SEC_OPT_PROC, FrmOption.TxtProcContentRow2.Text)
    Call SetIniValue(INI_KEY_PROC_OPT_WHERE, INI_SEC_OPT_PROC, IIf(FrmOption.OptProcRow.Value, OptModuleType.aOptRow, OptModuleType.aOptCom))
    Call SetIniValue(INI_KEY_PROC_REM_COMMENT, INI_SEC_OPT_PROC, FrmOption.TxtProcRemComment.Text)
    Call SetIniValue(INI_KEY_PROC_CONTENT_EXIST, INI_SEC_OPT_PROC, FrmOption.CheckExitProcContent.Value)
    Call SetIniValue(INI_KEY_PROC_CONTENT, INI_SEC_OPT_PROC, FrmOption.TxtProcContentComment.Text)
    
    Call SetIniValue(INI_KEY_EDIT_NORMAL_SELECT, INI_SEC_OPT_EDIT, FrmOption.CheckNormal.Value)
    Call SetIniValue(INI_KEY_EDIT_SHEET_SELECT, INI_SEC_OPT_EDIT, FrmOption.CheckSh.Value)
    Call SetIniValue(INI_KEY_EDIT_FRM_SELECT, INI_SEC_OPT_EDIT, FrmOption.CheckFrm.Value)
    Call SetIniValue(INI_KEY_EDIT_CLS_SELECT, INI_SEC_OPT_EDIT, FrmOption.CheckCls.Value)
    Call SetIniValue(INI_KEY_EDIT_ACN_SELECT, INI_SEC_OPT_EDIT, FrmOption.CheckUseOSNm.Value)
    Call SetIniValue(INI_KEY_EDIT_NOW_SELECT, INI_SEC_OPT_EDIT, FrmOption.CheckUseNow.Value)
    Call SetIniValue(INI_KEY_EDIT_AUT_NAME, INI_SEC_OPT_EDIT, FrmOption.TxtAuthor.Text)
    Call SetIniValue(INI_KEY_EDIT_CRE_DATE, INI_SEC_OPT_EDIT, FrmOption.TxtDate.Text)
End Sub

Public Sub IniRead()
    With cIniKeyList
        .aModuleContentRow = CInt(GetIniValue(INI_KEY_MODULE_CONTENT_ROW, INI_SEC_OPT_MODULE))
        .aModuleContentRow2 = CInt(GetIniValue(INI_KEY_MODULE_CONTENT_ROW2, INI_SEC_OPT_MODULE))
        .aModuleRemComment = GetIniValue(INI_KEY_MODULE_REM_COMMENT, INI_SEC_OPT_MODULE)
        .aModuleContentNotExist = CBool(GetIniValue(INI_KEY_MODULE_CONTENT_EXIST, INI_SEC_OPT_MODULE))
        .aProcContentRow = CInt(GetIniValue(INI_KEY_PROC_CONTENT_ROW, INI_SEC_OPT_PROC))
        .aProcContentRow2 = CInt(GetIniValue(INI_KEY_PROC_CONTENT_ROW2, INI_SEC_OPT_PROC))
        .aProcOptWhere = CInt(GetIniValue(INI_KEY_PROC_OPT_WHERE, INI_SEC_OPT_PROC))
        .aProcRemComment = GetIniValue(INI_KEY_PROC_REM_COMMENT, INI_SEC_OPT_PROC)
        .aProcContentNotExist = CBool(GetIniValue(INI_KEY_PROC_CONTENT_EXIST, INI_SEC_OPT_PROC))
        .aProcContent = GetIniValue(INI_KEY_PROC_CONTENT, INI_SEC_OPT_PROC)
        .aNormalSelect = CBool(GetIniValue(INI_KEY_EDIT_NORMAL_SELECT, INI_SEC_OPT_EDIT))
        .aSheetSelect = CBool(GetIniValue(INI_KEY_EDIT_SHEET_SELECT, INI_SEC_OPT_EDIT))
        .aFrmSelect = CBool(GetIniValue(INI_KEY_EDIT_FRM_SELECT, INI_SEC_OPT_EDIT))
        .aClsSelect = CBool(GetIniValue(INI_KEY_EDIT_CLS_SELECT, INI_SEC_OPT_EDIT))
        .aAcnSelect = CBool(GetIniValue(INI_KEY_EDIT_ACN_SELECT, INI_SEC_OPT_EDIT))
        .aNowSelect = CBool(GetIniValue(INI_KEY_EDIT_NOW_SELECT, INI_SEC_OPT_EDIT))
        .aAutName = GetIniValue(INI_KEY_EDIT_AUT_NAME, INI_SEC_OPT_EDIT)
        .aCreDate = GetIniValue(INI_KEY_EDIT_CRE_DATE, INI_SEC_OPT_EDIT)
    End With
End Sub
