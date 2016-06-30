VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 휴폐업조회 SDK 예제"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   10995
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnGetPopbillURL_CHRG 
      Caption         =   "포인트 충전 URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   1965
      Width           =   1575
   End
   Begin VB.CommandButton btnGetPopbillURL_LOGIN 
      Caption         =   "팝빌 로그인 URL"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1485
      Width           =   1575
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "조회단가 확인"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1965
      Width           =   1695
   End
   Begin VB.CommandButton btnGetBalance 
      Caption         =   "잔여포인트 확인"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1485
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
      Caption         =   "회원정보"
      Height          =   2175
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "회원가입"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "가입여부 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID 중복 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "포인트 관련"
      Height          =   2175
      Left            =   2400
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "과금정보 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnGetPartnerBalance 
         Caption         =   "파트너포인트 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "팝빌 URL 관련"
      Height          =   2175
      Left            =   4440
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "담당자 관련"
      Height          =   2175
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "담당자 정보 수정"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "담당자 목록 조회"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "담당자 추가"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "휴폐업조회"
      Height          =   975
      Left            =   360
      TabIndex        =   12
      Top             =   3600
      Width           =   8295
      Begin VB.CommandButton btnCheckCorpNums 
         Caption         =   "대량조회"
         Height          =   495
         Left            =   5520
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "단건조회"
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
         Text            =   "4108600477"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "조회할 사업자번호 : "
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "팝빌기본 API"
      Height          =   2535
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   10095
      Begin VB.Frame Frame7 
         Caption         =   "회사정보 관련"
         Height          =   2175
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label2 
      Caption         =   "팝빌회원 아이디 : "
      Height          =   225
      Left            =   4200
      TabIndex        =   2
      Top             =   375
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "팝빌회원 사업자번호 : "
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

'링크아이디
Private Const linkID = "TESTER"
'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

Private ClosedownService As New PBCDService

Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

'가입여부 확인
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckIsMember(txtCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    
    Set ChargeInfo = ClosedownService.GetChargeInfo(txtCorpNum.Text, txtUserID.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "unitCost (단가[정액제-월요금, 정량제-조회단가]) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    
    Set CorpInfo = ClosedownService.GetCorpInfo(txtCorpNum.Text, txtUserID.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    
    tmp = tmp + "ceoname : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'회원가입
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    joinData.linkID = linkID            '링크 아이디
    joinData.CorpNum = "1231212312"     '사업자번호, "-" 제외 10자리.
    joinData.CEOName = "대표자성명"
    joinData.CorpName = "회원상호"
    joinData.Addr = "주소"
    joinData.ZipCode = "500-100"
    joinData.BizType = "업태"
    joinData.BizClass = "업종"
    joinData.ID = "userid"                      '6자 이상 20자 미만.
    joinData.PWD = "pwd_must_be_long_enough"    '6자 이상 20자 미만.
    joinData.contactName = "담당자성명"
    joinData.ContactTEL = "02-999-9999"
    joinData.ContactHP = "010-1234-5678"
    joinData.ContactFAX = "02-999-9998"
    joinData.ContactEmail = "test@test.com"
    
    Set Response = ClosedownService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub
'잔여포인트 확인
Private Sub btnGetBalance_Click()
     Dim balance As Double
    
    balance = ClosedownService.GetBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "잔여포인트 : " + CStr(balance)
End Sub

Private Sub btnListContact_Click()
    Dim resultList As Collection
        
    Set resultList = ClosedownService.ListContact(txtCorpNum.Text, txtUserID.Text)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
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

Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.ID = "testkorea_20151007"      '담당자 아이디
    joinData.PWD = "test@test.com"          '비밀번호
    joinData.personName = "담당자명"        '담당자명
    joinData.tel = "070-1234-1234"          '연락처
    joinData.hp = "010-1234-1234"           '휴대폰번호
    joinData.email = "test@test.com"        '이메일 주소
    joinData.fax = "070-1234-1234"          '팩스번호
    joinData.searchAllAllowYN = True        '전체조회여부, Ture-회사조회, False-개인조회
    joinData.mgrYN = False                  '관리자 권한여부
        
    Set Response = ClosedownService.RegistContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

'조회단가 확인
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = ClosedownService.GetUnitCost(txtCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "조회단가 : " + CStr(unitCost)
End Sub
'팝빌로그인 URL
Private Sub btnGetPopbillURL_LOGIN_Click()
    Dim url As String
    
    url = ClosedownService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "LOGIN")
    
    If url = "" Then
         MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub
'포인트충전 URL
Private Sub btnGetPopbillURL_CHRG_Click()
    Dim url As String
    
    url = ClosedownService.GetPopbillURL(txtCorpNum.Text, txtUserID.Text, "CHRG")
    
    If url = "" Then
         MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    MsgBox "URL : " + vbCrLf + url
End Sub
'파트너 잔여포인트
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetPartnerBalance(txtCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub
'휴폐업조회 단건
Private Sub btnCheckCorpNum_Click()
    Dim CorpState As PBCorpState
    Dim tmp As String
    
    Set CorpState = ClosedownService.CheckCorpNum(txtCorpNum.Text, txtCheckCorpNum.Text)
    
    If CorpState Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "]" + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* type (사업 유형) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    tmp = tmp + "corpNum : " + CorpState.CorpNum + vbCrLf
    tmp = tmp + "state : " + CorpState.state + vbCrLf
    tmp = tmp + "type : " + CorpState.ctype + vbCrLf
    tmp = tmp + "stateDate(휴폐업일자) : " + CorpState.stateDate + vbCrLf
    tmp = tmp + "checkDate(국세청 확인일자) : " + CorpState.checkDate
    
    MsgBox tmp, , "휴폐업조회 - 단건"
End Sub

Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    joinData.personName = "담당자명_수정"  '담당자명
    joinData.tel = "070-1234-1234"         '연락처
    joinData.hp = "010-1234-1234"          '휴대폰번호
    joinData.email = "test@test.com"       '이메일 주소
    joinData.fax = "070-1234-1234"         '팩스번호
    joinData.searchAllAllowYN = True       '전체조회여부, Ture-회사조회, False-개인조
    joinData.mgrYN = False                 '관리자 권한여부
                
    Set Response = ClosedownService.UpdateContact(txtCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    CorpInfo.CEOName = "대표자"         '대표자명
    CorpInfo.CorpName = "상호_수정"          '상호명
    CorpInfo.Addr = "서울특별시"        '주소
    CorpInfo.BizType = "업태"           '업태
    CorpInfo.BizClass = "업종"          '업종
    
    Set Response = ClosedownService.UpdateCorpInfo(txtCorpNum.Text, CorpInfo, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("[" + CStr(Response.code) + "] " + Response.message)
End Sub

Private Sub txtCheckCorpNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call btnCheckCorpNum_Click
    End If
End Sub
'휴폐업조회 대량
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    Dim CorpNumList As New Collection
    
    '조회할 사업자번호 배열, 최대 1000건
    CorpNumList.Add "1234567890"
    CorpNumList.Add "4108600477"
        
    Set resultList = ClosedownService.CheckCorpNums(txtCorpNum.Text, CorpNumList)
     
    If resultList Is Nothing Then
        MsgBox ("[" + CStr(ClosedownService.LastErrCode) + "] " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    Dim tmp As String
    Dim state As PBCorpState
    
    tmp = tmp + "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* type (사업 유형) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    For Each state In resultList
        tmp = tmp + "corpNum : " + state.CorpNum + vbCrLf
        tmp = tmp + "state : " + state.state + vbCrLf
        tmp = tmp + "type : " + state.ctype + vbCrLf
        tmp = tmp + "stateDate(휴폐업일자) : " + state.stateDate + vbCrLf
        tmp = tmp + "checkDate(국세청 확인일자) : " + state.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp, , "휴폐업조회 - 대량"
End Sub

Private Sub Form_Load()
    '모듈 초기화
    ClosedownService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(테스트용), False(상업용)
    ClosedownService.IsTest = True
End Sub

