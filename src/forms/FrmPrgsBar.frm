VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPrgsBar 
   Caption         =   "�������ł�...���΂炭���҂���������"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "FrmPrgsBar.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FrmPrgsBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'/** �X�e�[�^�X�o�[�\���p���b�Z�[�W */
Private cMsg As String

'/** �ő吔 */
Private cMax As Long

'/** �X�e�[�^�X�o�[�\���p���b�Z�[�W */
Private cMsgMain As String

'/** �ő吔 */
Private cMaxMain As Long

'/** ���݂̐i���� */
Private cStatus As Long

'/** ���݂̐i���� */
Private cStatusMain As Long

'/** �ȉ��v���O���X�o�[�\���萔 �y��F�z"�������ł��@����20%:��������������������*/
Private Const WHITEBOX     As String = "��"
Private Const BLACKBOX     As String = "��"
Private Const STATUS_COUNT As Integer = 10

'*********************************************************
'���́F���b�Z�[�W�ƍő�l�̐ݒ�
'---------------------------------------------------------
'�������FaMsg :�X�e�[�^�X�o�[�\���p���b�Z�[�W
'        aMax :�ő吔
'---------------------------------------------------------
'�߂�l�FNone
'*********************************************************
Public Sub SubInit(aMsg As String, aMax As Integer)
    cMsg = aMsg
    cMax = aMax
    cStatus = 0
End Sub
'*********************************************************
'���́F���b�Z�[�W�ƍő�l�̐ݒ�
'---------------------------------------------------------
'�������FaMsg :�X�e�[�^�X�o�[�\���p���b�Z�[�W
'        aMax :�ő吔
'---------------------------------------------------------
'�߂�l�FNone
'*********************************************************
Public Sub MainInit(aMsgMain As String, aMaxMain As Integer)
    cMsgMain = aMsgMain
    cMaxMain = aMaxMain
    cStatusMax = 0
End Sub

'*********************************************************
'���́F�i����1�i�߂�
'---------------------------------------------------------
'�������FNone
'---------------------------------------------------------
'�߂�l�FNone
'*********************************************************
Public Sub Add()
    cStatus = cStatus + 1
    Disp
End Sub

'*********************************************************
'���́F�i����1�i�߂�
'---------------------------------------------------------
'�������FNone
'---------------------------------------------------------
'�߂�l�FNone
'*********************************************************
Public Sub AddMain()
    cStatusMain = cStatusMain + 1
End Sub
'*********************************************************
'���́F�X�e�[�^�X�o�[�̕\��
'---------------------------------------------------------
'�������FNone
'---------------------------------------------------------
'�߂�l�FNone
'*********************************************************
Public Sub Disp()
    DoEvents
    LblMainProcessName.Caption = cMsgMain
    LblMainProc.Caption = cStatusMain & " / " & cMaxMain
    LblSubProcessName.Caption = cMsg
    LblSubProc.Caption = " ����" & Round(cStatus * 100 / cMax) & "%�F" & String(cStatus * STATUS_COUNT / cMax, BLACKBOX) & String(STATUS_COUNT - cStatus * STATUS_COUNT / cMax, WHITEBOX)
End Sub


