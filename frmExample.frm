VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� �������ȸ SDK ����"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   9150
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnCheckCorpNums 
      Caption         =   "�뷮��ȸ"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton btnCheckCorpNum 
      Caption         =   "�ܰ���ȸ"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtCheckCorpNum 
      Height          =   270
      Left            =   2400
      TabIndex        =   12
      Text            =   "4108600477"
      Top             =   3195
      Width           =   1815
   End
   Begin VB.CommandButton btnGetPartnerBalance 
      Caption         =   "��Ʈ������Ʈ Ȯ��"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   1485
      Width           =   1695
   End
   Begin VB.CommandButton btnGetPopbillURL_CHRG 
      Caption         =   "����Ʈ ���� URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   1965
      Width           =   1575
   End
   Begin VB.CommandButton btnGetPopbillURL_LOGIN 
      Caption         =   "�˺� �α��� URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   1485
      Width           =   1575
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "��ȸ�ܰ� Ȯ��"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1965
      Width           =   1575
   End
   Begin VB.CommandButton btnGetBalance 
      Caption         =   "�ܿ�����Ʈ Ȯ��"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1485
      Width           =   1575
   End
   Begin VB.CommandButton btnJoinMember 
      Caption         =   "ȸ������"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1965
      Width           =   1455
   End
   Begin VB.CommandButton btnCheckIsMember 
      Caption         =   "���Կ��� Ȯ��"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1485
      Width           =   1455
   End
   Begin VB.TextBox txtUserID 
      Height          =   270
      Left            =   5760
      TabIndex        =   3
      Text            =   "testkorea"
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtCorpNum 
      Height          =   270
      Left            =   2280
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "ȸ������"
      Height          =   1335
      Left            =   480
      TabIndex        =   15
      Top             =   1245
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "����Ʈ ����"
      Height          =   1335
      Left            =   2400
      TabIndex        =   16
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "�˺� URL ����"
      Height          =   1335
      Left            =   4440
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "��Ʈ�� ����"
      Height          =   1335
      Left            =   6480
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Frame Frame5 
      Caption         =   "�������ȸ"
      Height          =   855
      Left            =   360
      TabIndex        =   19
      Top             =   2880
      Width           =   8295
      Begin VB.Label Label4 
         Caption         =   "��ȸ�� ����ڹ�ȣ : "
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�˺��⺻ API"
      Height          =   1815
      Left            =   360
      TabIndex        =   21
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label Label3 
      Caption         =   "��ȸ�� ����ڹ�ȣ : "
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "�˺�ȸ�� ���̵� : "
      Height          =   225
      Left            =   4200
      TabIndex        =   2
      Top             =   375
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�˺�ȸ�� ����ڹ�ȣ : "
      Height          =   230
      Left            =   360
      TabIndex        =   0
      Top             =   380
      Width           =   1935
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��ũ���̵�
Private Const linkID = "TESTER"
'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private ClosedownService As New PBCDService
'���Կ��� Ȯ��
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (CStr(Response.code) + " | " + Response.message)
End Sub
'ȸ������
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID            '��ũ ���̵�
    joinData.CorpNum = "1231212312"     '����ڹ�ȣ, "-" ���� 10�ڸ�.
    joinData.CEOName = "��ǥ�ڼ���"
    joinData.CorpName = "ȸ����ȣ"
    joinData.Addr = "�ּ�"
    joinData.ZipCode = "500-100"
    joinData.BizType = "����"
    joinData.BizClass = "����"
    joinData.ID = "userid"                      '6�� �̻� 20�� �̸�.
    joinData.PWD = "pwd_must_be_long_enough"    '6�� �̻� 20�� �̸�.
    joinData.contactName = "����ڼ���"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = ClosedownService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox (CStr(Response.code) + " | " + Response.message)
End Sub
'�ܿ�����Ʈ Ȯ��
Private Sub btnGetBalance_Click()
     Dim balance As Double
    
    balance = ClosedownService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub
'��ȸ�ܰ� Ȯ��
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = ClosedownService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��ȸ�ܰ� : " + CStr(unitCost)
End Sub
'�˺��α��� URL
Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = ClosedownService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub
'����Ʈ���� URL
Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = ClosedownService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub
'��Ʈ�� �ܿ�����Ʈ
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub
'�������ȸ �ܰ�
Private Sub btnCheckCorpNum_Click()
    Dim CorpState As PBCorpState
    Dim tmp As String
    
    Set CorpState = ClosedownService.CheckCorpNum(txtCorpNum.Text, txtCheckCorpNum.Text)
    
    If CorpState Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "]" + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    tmp = tmp + "corpNum : " + CorpState.CorpNum + vbCrLf
    tmp = tmp + "state : " + CorpState.state + vbCrLf
    tmp = tmp + "type : " + CorpState.ctype + vbCrLf
    tmp = tmp + "stateDate(���������) : " + CorpState.stateDate + vbCrLf
    tmp = tmp + "checkDate(����û Ȯ������) : " + CorpState.checkDate
    
    MsgBox tmp, , "�������ȸ - �ܰ�"
End Sub
Private Sub txtCheckCorpNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnCheckCorpNum_Click
    End If
End Sub
'�������ȸ �뷮
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    Dim CorpNumList As New Collection
    
    '��ȸ�� ����ڹ�ȣ �迭, �ִ� 1000��
    CorpNumList.Add "1234567890"
    CorpNumList.Add "4108600477"
        
    Set resultList = ClosedownService.CheckCorpNums(txtCorpNum.Text, CorpNumList)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    Dim state As PBCorpState
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    For Each state In resultList
        tmp = tmp + "corpNum : " + state.CorpNum + vbCrLf
        tmp = tmp + "state : " + state.state + vbCrLf
        tmp = tmp + "type : " + state.ctype + vbCrLf
        tmp = tmp + "stateDate(���������) : " + state.stateDate + vbCrLf
        tmp = tmp + "checkDate(����û Ȯ������) : " + state.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp, , "�������ȸ - �뷮"
End Sub

Private Sub Form_Load()
    '��� �ʱ�ȭ
    ClosedownService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(�׽�Ʈ��), False(�����)
    ClosedownService.IsTest = True
End Sub

