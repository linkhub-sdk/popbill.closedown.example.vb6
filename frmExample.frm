VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� �������ȸ SDK ����"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   10995
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton btnGetPopbillURL_CHRG 
      Caption         =   "����Ʈ ���� URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1965
      Width           =   1575
   End
   Begin VB.CommandButton btnGetPopbillURL_LOGIN 
      Caption         =   "�˺� �α��� URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1485
      Width           =   1575
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "��ȸ�ܰ� Ȯ��"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton btnGetBalance 
      Caption         =   "�ܿ�����Ʈ Ȯ��"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
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
      Height          =   2175
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "ȸ������"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "���Կ��� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID �ߺ� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����Ʈ ����"
      Height          =   2175
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "�������� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton btnGetPartnerBalance 
         Caption         =   "��Ʈ������Ʈ Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�˺� URL ����"
      Height          =   2175
      Left            =   4440
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "����� ����"
      Height          =   2175
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "����� ���� ����"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "����� ��� ��ȸ"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "����� �߰�"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "�������ȸ"
      Height          =   975
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   8295
      Begin VB.CommandButton btnCheckCorpNums 
         Caption         =   "�뷮��ȸ"
         Height          =   495
         Left            =   5520
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "�ܰ���ȸ"
         Height          =   495
         Left            =   4080
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCheckCorpNum 
         Height          =   270
         Left            =   2040
         TabIndex        =   15
         Text            =   "6798700433"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "��ȸ�� ����ڹ�ȣ : "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�˺��⺻ API"
      Height          =   2535
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   10095
      Begin VB.Frame Frame7 
         Caption         =   "ȸ������ ����"
         Height          =   2175
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
      End
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
'=========================================================================
'
' �˺� �������ȸ API VB 6.0 SDK Example
'
' - VB6 SDK ����ȯ�� ������� �ȳ� :
' - ������Ʈ ���� : 2016-10-12
' - ���� ������� ����ó : 1600-8536 / 070-4504-2991 (���� / �����Ѵ븮)
' - ���� ������� �̸��� : dev@linkhub.co.kr
'
' <�׽�Ʈ �������� �غ����>
' 1) 25, 28�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
' 2) �˺� ���߿� ����Ʈ(test.popbill.com)�� ����ȸ������ �����մϴ�.
'=========================================================================

Option Explicit

'=========================================================================
' - ��������(��ũ���̵�, ���Ű)�� ��Ʈ���� ����ȸ���� �ĺ��ϴ�
'   ������ ���Ǵ� ������ ������� �ʵ��� �����Ͻñ� �ٶ��ϴ�.
' - ����� ��ȯ���Ŀ��� ��������(��ũ���̵�, ���Ű)�� ������� �ʽ��ϴ�.
'=========================================================================

'��ũ���̵�
Private Const linkID = "TESTER"

'���Ű. ���⿡ �����Ͻñ� �ٶ��ϴ�.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'�������ȸ ���� ��ü ����
Private ClosedownService As New PBCDService

'=========================================================================
' �˺� ȸ�����̵� �ߺ����θ� Ȯ���մϴ�.
'=========================================================================

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �ش� ������� ��Ʈ�� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - LinkID�� ���������� �����Ǿ� �ִ� ��ũ���̵� ���Դϴ�.
'=========================================================================

Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �������ȸ API ���� ���������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = ClosedownService.GetChargeInfo(txtCorpNum.Text, txtUserID.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (�ܰ�[������-�����, ������-��ȸ�ܰ�]) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
'=========================================================================

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = ClosedownService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname(��ǥ�ڼ���) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName(��ȣ��) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr(�ּ�) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType(����) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass(����) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����ȸ�� ������ ��û�մϴ�.
'=========================================================================

Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '��ũ ���̵�
    joinData.linkID = linkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1231212312"
    
    '��ǥ�ڼ���, �ִ� 30��
    joinData.CEOName = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 70��
    joinData.CorpName = "ȸ����ȣ"
    
    '�ּ�, �ִ� 300��
    joinData.Addr = "�ּ�"
    
    '����, �ִ� 40��
    joinData.BizType = "����"
    
    '����, �ִ� 40��
    joinData.BizClass = "����"
    
    '���̵�, 6���̻� 20�� �̸�
    joinData.ID = "userid"
    
    '��й�ȣ, 6���̻� 20�� �̸�
    joinData.PWD = "pwd_must_be_long_enough"
    
    '����ڸ�, �ִ� 30��
    joinData.contactName = "����ڼ���"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.ContactHP = "010-1234-5678"
    
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.ContactFAX = "02-999-9998"
    
    '����� ����, �ִ� 70��
    joinData.ContactEmail = "test@test.com"
    
    Set Response = ClosedownService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ��Ʈ�ʰ����� ��� ��Ʈ�� �ܿ�����Ʈ(GetPartnerBalance API)
'   �� ���� Ȯ���Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "�ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ���� ����� ����� Ȯ���մϴ�.
'=========================================================================

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = ClosedownService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = "id | email | hp | personName | searchAllAllowYN | tel | fax | mgrYN | regDT " + vbCrLf
    
    Dim info As PBContactInfo
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.email + " | " + info.hp + " | " + info.personName + " | " + CStr(info.searchAllAllowYN) _
                + info.tel + " | " + info.fax + " | " + CStr(info.mgrYN) + " | " + info.regDT + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ����ڸ� �űԷ� ����մϴ�.
'=========================================================================

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 20�� �̸�
    joinData.ID = "testkorea_20161011"
    
    '��й�ȣ, 6�� �̻� 20�� �̸�
    joinData.PWD = "test@test.com"
    
    '����ڸ�, �ִ� 30��
    joinData.personName = "����ڸ�"
    
    '����� ����ó
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '����� �����ּ�
    joinData.email = "test@test.com"
    
    '����� �ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    'ȸ����ȸ ���ѿ���, true-ȸ����ȸ / false-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
        
    Set Response = ClosedownService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' �������ȸ �ܰ��� Ȯ���մϴ�.
'=========================================================================

Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = ClosedownService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��ȸ�ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺�(www.popbill.com)�� �α��ε� �˺� URL�� ��ȯ�մϴ�.
' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = ClosedownService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ���� URL�� ��ȯ�մϴ�.
' - URL ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
'=========================================================================

Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = ClosedownService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - ���ݹ���� ���������� ��� ����ȸ�� �ܿ�����Ʈ(GetBalance API)��
'   �̿��Ͻñ� �ٶ��ϴ�.
'=========================================================================

Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' 1���� ����ڿ� ���� ��������θ� ��ȸ�մϴ�.
'=========================================================================

Private Sub btnCheckCorpNum_Click()
    Dim CorpState As PBCorpState
    Dim tmp As String
    
    Set CorpState = ClosedownService.CheckCorpNum(txtCorpNum.Text, txtCheckCorpNum.Text)
    
    If CorpState Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    tmp = tmp + "corpNum (����ڹ�ȣ) : " + CorpState.CorpNum + vbCrLf
    tmp = tmp + "state (���������) : " + CorpState.state + vbCrLf
    tmp = tmp + "type (�������) : " + CorpState.ctype + vbCrLf
    tmp = tmp + "stateDate(���������) : " + CorpState.stateDate + vbCrLf
    tmp = tmp + "checkDate(����û Ȯ������) : " + CorpState.checkDate
    
    MsgBox tmp, , "�������ȸ - �ܰ�"
End Sub

'=========================================================================
' ����ȸ���� ����� ������ �����մϴ�.
'=========================================================================

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����ڸ�
    joinData.personName = "����ڸ�_����"
    
    '����ó
    joinData.tel = "070-1234-1234"
    
    '�޴�����ȣ
    joinData.hp = "010-1234-1234"
    
    '�̸��� �ּ�
    joinData.email = "test@test.com"
    
    '�ѽ���ȣ
    joinData.fax = "070-1234-1234"
    
    '��ü��ȸ����, True-ȸ����ȸ, False-������ȸ
    joinData.searchAllAllowYN = True
    
    '������ ���ѿ���
    joinData.mgrYN = False
                
    Set Response = ClosedownService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� �����մϴ�
'=========================================================================

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�
    CorpInfo.CEOName = "��ǥ��"
    
    '��ȣ
    CorpInfo.CorpName = "��ȣ"
    
    '�ּ�
    CorpInfo.Addr = "����Ư����"
    
    '����
    CorpInfo.BizType = "����"
    
    '����
    CorpInfo.BizClass = "����"
    
    Set Response = ClosedownService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

Private Sub txtCheckCorpNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnCheckCorpNum_Click
    End If
End Sub

'=========================================================================
' �ټ��� ����ڿ� ���� ��������θ� ��ȸ�մϴ�.
'=========================================================================

Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    Dim CorpNumList As New Collection
    Dim tmp As String
    Dim state As PBCorpState
    
    '��ȸ�� ����ڹ�ȣ �迭, �ִ� 1000��
    CorpNumList.Add "1234567890"
    CorpNumList.Add "6798700433"
    CorpNumList.Add "1111111111"
        
    Set resultList = ClosedownService.CheckCorpNums(txtCorpNum.Text, CorpNumList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    For Each state In resultList
        tmp = tmp + "corpNum (����ڹ�ȣ) : " + state.CorpNum + vbCrLf
        tmp = tmp + "state (���������) : " + state.state + vbCrLf
        tmp = tmp + "type (�������) : " + state.ctype + vbCrLf
        tmp = tmp + "stateDate(���������) : " + state.stateDate + vbCrLf
        tmp = tmp + "checkDate(����û Ȯ������) : " + state.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp, , "�������ȸ - �뷮"
End Sub

Private Sub Form_Load()
    '��� �ʱ�ȭ
    ClosedownService.Initialize linkID, SecretKey
    
    '����ȯ�� ������ True(���߿�), False(�����)
    ClosedownService.IsTest = True
End Sub

