VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPrgsBar 
   Caption         =   "処理中です...しばらくお待ちください"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FrmPrgsBar.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FrmPrgsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'/** ステータスバー表示用メッセージ */
Private cMsg As String

'/** 最大数 */
Private cMax As Long

'/** ステータスバー表示用メッセージ */
Private cMsgMain As String

'/** 最大数 */
Private cMaxMain As Long

'/** 現在の進捗数 */
Private cStatus As Long

'/** 現在の進捗数 */
Private cStatusMain As Long

'/** 以下プログレスバー表示定数 【例：】"処理中です　現在20%:■■□□□□□□□□*/
Private Const WHITEBOX     As String = "□"
Private Const BLACKBOX     As String = "■"
Private Const STATUS_COUNT As Integer = 10

'*********************************************************
'名称：メッセージと最大値の設定
'---------------------------------------------------------
'引き数：aMsg :ステータスバー表示用メッセージ
'        aMax :最大数
'---------------------------------------------------------
'戻り値：None
'*********************************************************
Public Sub SubInit(aMsg As String, aMax As Integer)
    cMsg = aMsg
    cMax = aMax
    cStatus = 0
End Sub
'*********************************************************
'名称：メッセージと最大値の設定
'---------------------------------------------------------
'引き数：aMsg :ステータスバー表示用メッセージ
'        aMax :最大数
'---------------------------------------------------------
'戻り値：None
'*********************************************************
Public Sub MainInit(aMsgMain As String, aMaxMain As Integer)
    cMsgMain = aMsgMain
    cMaxMain = aMaxMain
    cStatusMax = 0
End Sub

'*********************************************************
'名称：進捗を1進める
'---------------------------------------------------------
'引き数：None
'---------------------------------------------------------
'戻り値：None
'*********************************************************
Public Sub Add()
    cStatus = cStatus + 1
    Disp
End Sub

'*********************************************************
'名称：進捗を1進める
'---------------------------------------------------------
'引き数：None
'---------------------------------------------------------
'戻り値：None
'*********************************************************
Public Sub AddMain()
    cStatusMain = cStatusMain + 1
End Sub
'*********************************************************
'名称：ステータスバーの表示
'---------------------------------------------------------
'引き数：None
'---------------------------------------------------------
'戻り値：None
'*********************************************************
Public Sub Disp()
    DoEvents
    LblMainProcessName.Caption = cMsgMain
    LblMainProc.Caption = cStatusMain & " / " & cMaxMain
    LblSubProcessName.Caption = cMsg
    LblSubProc.Caption = " 現在" & Round(cStatus * 100 / cMax) & "%：" & String(cStatus * STATUS_COUNT / cMax, BLACKBOX) & String(STATUS_COUNT - cStatus * STATUS_COUNT / cMax, WHITEBOX)
End Sub


