VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "�˺� �������ȸ SDK ����"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   15180
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.TextBox txtURL 
      Height          =   270
      Left            =   11400
      TabIndex        =   36
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "��ȸ�ܰ� Ȯ��"
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
   Begin VB.TextBox txtUserCorpNum 
      Height          =   270
      Left            =   2280
      TabIndex        =   1
      Text            =   "1234567890"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "ȸ������"
      Height          =   2295
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "ȸ������"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "���Կ��� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID �ߺ� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����Ʈ ����"
      Height          =   2295
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "�������� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�˺� URL ����"
      Height          =   2295
      Left            =   9000
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
      Begin VB.CommandButton btnGetAccessURL 
         Caption         =   "�˺� �α��� URL"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "����� ����"
      Height          =   2295
      Left            =   10920
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetContactInfo 
         Caption         =   "����� ���� Ȯ��"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "����� ���� ����"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "����� ��� ��ȸ"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "����� �߰�"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "�������ȸ"
      Height          =   975
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   8295
      Begin VB.CommandButton btnCheckCorpNums 
         Caption         =   "�뷮��ȸ"
         Height          =   495
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "�ܰ���ȸ"
         Height          =   495
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCheckCorpNum 
         Height          =   270
         Left            =   2040
         TabIndex        =   12
         Text            =   "6798700433"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "��ȸ�� ����ڹ�ȣ : "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "�˺��⺻ API"
      Height          =   2655
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   14535
      Begin VB.Frame Frame9 
         Caption         =   "��Ʈ�ʰ��� ����Ʈ"
         Height          =   2295
         Left            =   6360
         TabIndex        =   26
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "����Ʈ ���� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "��Ʈ������Ʈ Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "�������� ����Ʈ"
         Height          =   2295
         Left            =   4080
         TabIndex        =   25
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "����Ʈ �������� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "����Ʈ ��볻�� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "����Ʈ ���� URL"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "�ܿ�����Ʈ Ȯ��"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "ȸ������ ����"
         Height          =   2295
         Left            =   12600
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "ȸ������ ����"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "ȸ������ ��ȸ"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label3 
      Caption         =   "URL : "
      Height          =   225
      Left            =   10680
      TabIndex        =   35
      Top             =   360
      Width           =   735
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
' �˺� �������ȸ API VB SDK Example
'
' - ������Ʈ ���� : 2022-04-06
' - ���� ������� ����ó : 1600-9854
' - ���� ������� �̸��� : code@linkhubcorp.com
' - VB SDK ����ȯ�� ������� �ȳ� : https://docs.popbill.com/closedown/tutorial/vb
'
' <�׽�Ʈ �������� �غ����>
' 1) 19, 22�� ���ο� ����� ��ũ���̵�(LinkID)�� ���Ű(SecretKey)��
'    ��ũ��� ���Խ� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
'
'=========================================================================

Option Explicit

'��ũ���̵�
Private Const LinkID = "TESTER"

'���Ű
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'�������ȸ Ŭ���� ����
Private ClosedownService As New PBCDService


'=========================================================================
' �˺� �������ȸ API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = ClosedownService.GetChargeInfo(txtUserCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (��ȸ�ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ڹ�ȣ�� ��ȸ�Ͽ� ����ȸ�� ���Կ��θ� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckIsMember(txtUserCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#CheckID
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
' ����ڸ� ����ȸ������ ����ó���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '���̵�, 6���̻� 50�� ����
    joinData.id = "userid"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '��Ʈ�ʸ�ũ ���̵�
    joinData.LinkID = LinkID
    
    '����ڹ�ȣ, '-'����, 10�ڸ�
    joinData.CorpNum = "1234567890"
    
    '��ǥ�ڼ���, �ִ� 100��
    joinData.CEOName = "��ǥ�ڼ���"
    
    '��ȣ��, �ִ� 200��
    joinData.CorpName = "ȸ����ȣ"
    
    '����� �ּ�, �ִ� 300��
    joinData.Addr = "�ּ�"
    
    '����, �ִ� 100��
    joinData.BizType = "����"
    
    '����, �ִ� 100��
    joinData.BizClass = "����"

    '����� ����, �ִ� 100��
    joinData.ContactName = "����ڼ���"
    
    '����� �̸���, �ִ� 100��
    joinData.ContactEmail = "test@test.com"
    
    '����� ����ó, �ִ� 20��
    joinData.ContactTEL = "02-999-9999"
    
    Set Response = ClosedownService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����� ��ȸ�� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = ClosedownService.GetUnitCost(txtUserCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��ȸ�ܰ� : " + CStr(unitCost)
End Sub

'=========================================================================
' �˺� �������ȸ API ���� ���������� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = ClosedownService.GetChargeInfo(txtUserCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (��ȸ�ܰ�) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (��������) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (��������) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim URL As String
    
    URL = ClosedownService.GetAccessURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� �����(�˺� �α��� ����)�� �߰��մϴ�.
' - https://docs.popbill.com/closedown/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�, 6�� �̻� 50�� ����
    joinData.id = "testkorea"
    
    '��й�ȣ, 8�� �̻� 20�� ����(����, ����, Ư������ ����)
    joinData.Password = "asdf$%^123"
    
    '����ڸ�, �ִ� 100��
    joinData.personName = "����ڸ�"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �����ּ�, �ִ� 100��
    joinData.email = "test@test.com"
    
    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
    
    Set Response = ClosedownService.RegistContact(txtUserCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ Ȯ���մϴ�.
' https://docs.popbill.com/closedown/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = "testkorea"
    
    Set info = ClosedownService.GetContactInfo(txtUserCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.tel + " | " _
            + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ����� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = ClosedownService.ListContact(txtUserCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(���̵�) | personName(����) | email(�̸���) | tel(����ó) | " _
         + "regDT(����Ͻ�) | searchRole(����� ����) | mgrYN(������ ����) | state(����) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ�� ����ڹ�ȣ�� ��ϵ� �����(�˺� �α��� ����) ������ �����մϴ�.
' - https://docs.popbill.com/closedown/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '����� ���̵�
    joinData.id = txtUserID.Text
    
    '����� ����, �ִ� 100��
    joinData.personName = "����ڸ�_����"
    
    '����� ����ó, �ִ� 20��
    joinData.tel = "070-1234-1234"
    
    '����� �޴�����ȣ, �ִ� 20��
    joinData.hp = "010-1234-1234"
        
    '����� �ѽ���ȣ, �ִ� 20��
    joinData.fax = "070-1234-1234"
    
    '����� �̸���, �ִ� 100��
    joinData.email = "test@test.com"

    '����� ����, 1-���� / 2-�б� / 3-ȸ��
    joinData.searchRole = 3
                
    Set Response = ClosedownService.UpdateContact(txtUserCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� ȸ�������� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = ClosedownService.GetCorpInfo(txtUserCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (��ǥ�ڼ���) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName (��ȣ��) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr (�ּ�) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType (����) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass (����) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' ����ȸ���� ȸ�� ������ �����մϴ�.
' - https://docs.popbill.com/closedown/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '��ǥ�ڸ�, �ִ� 100��
    CorpInfo.CEOName = "��ǥ��"
    
    '��ȣ, �ִ� 200��
    CorpInfo.CorpName = "��ȣ"
    
    '�ּ�, �ִ� 300��
    CorpInfo.Addr = "����Ư����"
    
    '����, �ִ� 100��
    CorpInfo.BizType = "����"
    
    '����, �ִ� 100��
    CorpInfo.BizClass = "����"
    
    Set Response = ClosedownService.UpdateCorpInfo(txtUserCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("�����ڵ� : " + CStr(Response.code) + vbCrLf + "����޽��� : " + Response.message)
End Sub

'=========================================================================
' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetBalance
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "����ȸ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim URL As String
    
    URL = ClosedownService.GetChargeURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ �������� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = ClosedownService.GetPaymentURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ȸ�� ����Ʈ ��볻�� Ȯ���� ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = ClosedownService.GetUseHistoryURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ��Ʈ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetPartnerBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "��Ʈ�� �ܿ�����Ʈ : " + CStr(balance)
End Sub

'=========================================================================
' ��Ʈ�� ����Ʈ ������ ���� �������� �˾� URL�� ��ȯ�մϴ�.
' - URL ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
' - https://docs.popbill.com/closedown/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
    
    URL = ClosedownService.GetPartnerURL(txtUserCorpNum.Text, "CHRG")
       
    If URL = "" Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' ����ڹ�ȣ 1�ǿ� ���� ����������� Ȯ���մϴ�.
' - https://docs.popbill.com/closedown/vb/api#CheckCorpNum
'=========================================================================
Private Sub btnCheckCorpNum_Click()
    Dim CorpState As PBCorpState
    Dim tmp As String
    
    Set CorpState = ClosedownService.CheckCorpNum(txtUserCorpNum.Text, txtCheckCorpNum.Text)
    
    If CorpState Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* taxType (����� ��������) : null-�˼�����, 10-�Ϲݰ�����, 20-�鼼������, 30-���̰�����, 31-���̰�����(���ݰ�꼭 �߱޻����), 40-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    tmp = tmp + "corpNum (����ڹ�ȣ) : " + CorpState.CorpNum + vbCrLf
    tmp = tmp + "state (���������) : " + CorpState.state + vbCrLf
    tmp = tmp + "taxType (����� ��������) : " + CorpState.taxType + vbCrLf
    tmp = tmp + "typeDate (�������� ��ȯ����) : " + CorpState.typeDate + vbCrLf
    tmp = tmp + "stateDate (���������) : " + CorpState.stateDate + vbCrLf
    tmp = tmp + "checkDate (����û Ȯ������) : " + CorpState.checkDate
    
    MsgBox tmp, , "�������ȸ - �ܰ�"
End Sub

'=========================================================================
' �ټ����� ����ڹ�ȣ�� ���� ����������� Ȯ���մϴ�. (�ִ� 1,000��)
' - https://docs.popbill.com/closedown/vb/api#CheckCorpNums
'=========================================================================
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    Dim CorpNumList As New Collection
    Dim tmp As String
    Dim state As PBCorpState
    
    '��ȸ�� ����ڹ�ȣ �迭 (�ִ� 1000��)
    CorpNumList.Add "1234567890"
    CorpNumList.Add "401-03-94930"
    CorpNumList.Add "120-86-75212"
        
    Set resultList = ClosedownService.CheckCorpNums(txtUserCorpNum.Text, CorpNumList)
     
    If resultList Is Nothing Then
        MsgBox ("�����ڵ� : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "����޽��� : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�" + vbCrLf
    tmp = tmp + "* taxType (����� ��������) : null-�˼�����, 10-�Ϲݰ�����, 20-�鼼������, 30-���̰�����, 31-���̰�����(���ݰ�꼭 �߱޻����), 40-�񿵸�����, �������" + vbCrLf + vbCrLf
    
    For Each state In resultList
        tmp = tmp + "corpNum (����ڹ�ȣ) : " + state.CorpNum + vbCrLf
        tmp = tmp + "state (���������) : " + state.state + vbCrLf
        tmp = tmp + "taxType (����� ��������) : " + state.taxType + vbCrLf
        tmp = tmp + "typeDate (�������� ��ȯ����) : " + state.typeDate + vbCrLf
        tmp = tmp + "stateDate (���������) : " + state.stateDate + vbCrLf
        tmp = tmp + "checkDate (����û Ȯ������) : " + state.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp, , "�������ȸ - �뷮"
End Sub

Private Sub Form_Load()

    '�������ȸ ��� �ʱ�ȭ
    ClosedownService.Initialize LinkID, SecretKey
    
    '����ȯ�漳����, True-���߿� False-�����
    ClosedownService.IsTest = True
    
    '������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
    ClosedownService.IPRestrictOnOff = True
    
    '���ýý��� �ð� ��뿩�� True-���, False-�̻��, �⺻��(False)
    ClosedownService.UseLocalTimeYN = False
End Sub




