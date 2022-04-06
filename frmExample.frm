VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 휴폐업조회 SDK 예제"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   15180
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtURL 
      Height          =   270
      Left            =   11400
      TabIndex        =   36
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton btnUnitCost 
      Caption         =   "조회단가 확인"
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
      Caption         =   "회원정보"
      Height          =   2295
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "회원가입"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "가입여부 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID 중복 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "포인트 관련"
      Height          =   2295
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "과금정보 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "팝빌 URL 관련"
      Height          =   2295
      Left            =   9000
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
      Begin VB.CommandButton btnGetAccessURL 
         Caption         =   "팝빌 로그인 URL"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "담당자 관련"
      Height          =   2295
      Left            =   10920
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetContactInfo 
         Caption         =   "담당자 정보 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "담당자 정보 수정"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "담당자 목록 조회"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "담당자 추가"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "휴폐업조회"
      Height          =   975
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Width           =   8295
      Begin VB.CommandButton btnCheckCorpNums 
         Caption         =   "대량조회"
         Height          =   495
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "단건조회"
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
         Caption         =   "조회할 사업자번호 : "
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "팝빌기본 API"
      Height          =   2655
      Left            =   360
      TabIndex        =   11
      Top             =   960
      Width           =   14535
      Begin VB.Frame Frame9 
         Caption         =   "파트너과금 포인트"
         Height          =   2295
         Left            =   6360
         TabIndex        =   26
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "연동과금 포인트"
         Height          =   2295
         Left            =   4080
         TabIndex        =   25
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPaymentURL 
            Caption         =   "포인트 결제내역 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton btnGetUseHistoryURL 
            Caption         =   "포인트 사용내역 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "회사정보 관련"
         Height          =   2295
         Left            =   12600
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
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
'=========================================================================
'
' 팝빌 휴폐업조회 API VB SDK Example
'
' - 업데이트 일자 : 2022-04-06
' - 연동 기술지원 연락처 : 1600-9854
' - 연동 기술지원 이메일 : code@linkhubcorp.com
' - VB SDK 연동환경 설정방법 안내 : https://docs.popbill.com/closedown/tutorial/vb
'
' <테스트 연동개발 준비사항>
' 1) 19, 22번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
'
'=========================================================================

Option Explicit

'링크아이디
Private Const LinkID = "TESTER"

'비밀키
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'휴폐업조회 클래스 선언
Private ClosedownService As New PBCDService


'=========================================================================
' 팝빌 휴폐업조회 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = ClosedownService.GetChargeInfo(txtUserCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (조회단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 사업자번호를 조회하여 연동회원 가입여부를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckIsMember(txtUserCorpNum.Text, LinkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사용하고자 하는 아이디의 중복여부를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#CheckID
'=========================================================================
Private Sub btnCheckID_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckID(txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 사용자를 연동회원으로 가입처리합니다.
' - https://docs.popbill.com/closedown/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 이하
    joinData.id = "userid"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '파트너링크 아이디
    joinData.LinkID = LinkID
    
    '사업자번호, '-'제외, 10자리
    joinData.CorpNum = "1234567890"
    
    '대표자성명, 최대 100자
    joinData.CEOName = "대표자성명"
    
    '상호명, 최대 200자
    joinData.CorpName = "회원상호"
    
    '사업장 주소, 최대 300자
    joinData.Addr = "주소"
    
    '업태, 최대 100자
    joinData.BizType = "업태"
    
    '종목, 최대 100자
    joinData.BizClass = "종목"

    '담당자 성명, 최대 100자
    joinData.ContactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    Set Response = ClosedownService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 휴폐업 조회시 과금되는 포인트 단가를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#GetUnitCost
'=========================================================================
Private Sub btnUnitCost_Click()
    Dim unitCost As Double
    
    unitCost = ClosedownService.GetUnitCost(txtUserCorpNum.Text)
    
    If unitCost < 0 Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "조회단가 : " + CStr(unitCost)
End Sub

'=========================================================================
' 팝빌 휴폐업조회 API 서비스 과금정보를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#GetChargeInfo
'=========================================================================
Private Sub btnGetChargeInfo_Click()
    Dim ChargeInfo As PBChargeInfo
    Dim tmp As String
    
    Set ChargeInfo = ClosedownService.GetChargeInfo(txtUserCorpNum.Text)
     
    If ChargeInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "unitCost (조회단가) : " + ChargeInfo.unitCost + vbCrLf
    tmp = tmp + "chargeMethod (과금유형) : " + ChargeInfo.chargeMethod + vbCrLf
    tmp = tmp + "rateSystem (과금제도) : " + ChargeInfo.rateSystem + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 팝빌 사이트에 로그인 상태로 접근할 수 있는 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/closedown/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim URL As String
    
    URL = ClosedownService.GetAccessURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 사업자번호에 담당자(팝빌 로그인 계정)를 추가합니다.
' - https://docs.popbill.com/closedown/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 이하
    joinData.id = "testkorea"
    
    '비밀번호, 8자 이상 20자 이하(영문, 숫자, 특수문자 조합)
    joinData.Password = "asdf$%^123"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
    
    Set Response = ClosedownService.RegistContact(txtUserCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 확인합니다.
' https://docs.popbill.com/closedown/vb/api#GetContactInfo
'=========================================================================
Private Sub btnGetContactInfo_Click()
    Dim tmp As String
    Dim info As PBContactInfo
    Dim ContactID As String
    
    ContactID = "testkorea"
    
    Set info = ClosedownService.GetContactInfo(txtUserCorpNum.Text, ContactID, txtUserID.Text)
    
    If info Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
   
    tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " + info.tel + " | " _
            + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
        
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 목록을 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#ListContact
'=========================================================================
Private Sub btnListContact_Click()
    Dim resultList As Collection
    Dim tmp As String
    Dim info As PBContactInfo
    
    Set resultList = ClosedownService.ListContact(txtUserCorpNum.Text)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | tel(연락처) | " _
         + "regDT(등록일시) | searchRole(담당자 권한) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.id + " | " + info.personName + " | " + info.email + " | " _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchRole) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원 사업자번호에 등록된 담당자(팝빌 로그인 계정) 정보를 수정합니다.
' - https://docs.popbill.com/closedown/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.id = txtUserID.Text
    
    '담당자 성명, 최대 100자
    joinData.personName = "담당자명_수정"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
        
    '담당자 팩스번호, 최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 이메일, 최대 100자
    joinData.email = "test@test.com"

    '담당자 권한, 1-개인 / 2-읽기 / 3-회사
    joinData.searchRole = 3
                
    Set Response = ClosedownService.UpdateContact(txtUserCorpNum.Text, joinData, txtUserID.Text)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 회사정보를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#GetCorpInfo
'=========================================================================
Private Sub btnGetCorpInfo_Click()
    Dim CorpInfo As PBCorpInfo
    Dim tmp As String
    
    Set CorpInfo = ClosedownService.GetCorpInfo(txtUserCorpNum.Text)
     
    If CorpInfo Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "ceoname (대표자성명) : " + CorpInfo.CEOName + vbCrLf
    tmp = tmp + "corpName (상호명) : " + CorpInfo.CorpName + vbCrLf
    tmp = tmp + "addr (주소) : " + CorpInfo.Addr + vbCrLf
    tmp = tmp + "bizType (업태) : " + CorpInfo.BizType + vbCrLf
    tmp = tmp + "bizClass (종목) : " + CorpInfo.BizClass + vbCrLf
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 회사 정보를 수정합니다.
' - https://docs.popbill.com/closedown/vb/api#UpdateCorpInfo
'=========================================================================
Private Sub btnUpdateCorpInfo_Click()
    Dim CorpInfo As New PBCorpInfo
    Dim Response As PBResponse
    
    '대표자명, 최대 100자
    CorpInfo.CEOName = "대표자"
    
    '상호, 최대 200자
    CorpInfo.CorpName = "상호"
    
    '주소, 최대 300자
    CorpInfo.Addr = "서울특별시"
    
    '업태, 최대 100자
    CorpInfo.BizType = "업태"
    
    '종목, 최대 100자
    CorpInfo.BizClass = "종목"
    
    Set Response = ClosedownService.UpdateCorpInfo(txtUserCorpNum.Text, CorpInfo)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 잔여포인트를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#GetBalance
'=========================================================================

Private Sub btnGetBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "연동회원 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 연동회원 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/closedown/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim URL As String
    
    URL = ClosedownService.GetChargeURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 포인트 결제내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/closedown/vb/api#GetPaymentURL
'=========================================================================
Private Sub btnGetPaymentURL_Click()
    Dim URL As String
           
    URL = ClosedownService.GetPaymentURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 연동회원 포인트 사용내역 확인을 위한 페이지의 팝업 URL을 반환합니다.
' - 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/closedown/vb/api#GetUseHistoryURL
'=========================================================================
Private Sub btnGetUseHistoryURL_Click()
    Dim URL As String
           
    URL = ClosedownService.GetUseHistoryURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 파트너의 잔여포인트를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#GetPartnerBalance
'=========================================================================
Private Sub btnGetPartnerBalance_Click()
    Dim balance As Double
    
    balance = ClosedownService.GetPartnerBalance(txtUserCorpNum.Text)
    
    If balance < 0 Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "파트너 잔여포인트 : " + CStr(balance)
End Sub

'=========================================================================
' 파트너 포인트 충전을 위한 페이지의 팝업 URL을 반환합니다.
' - URL 반환되는 URL은 보안 정책상 30초 동안 유효하며, 시간을 초과한 후에는 해당 URL을 통한 페이지 접근이 불가합니다.
' - https://docs.popbill.com/closedown/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim URL As String
    
    URL = ClosedownService.GetPartnerURL(txtUserCorpNum.Text, "CHRG")
       
    If URL = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + URL
    txtURL.Text = URL
End Sub

'=========================================================================
' 사업자번호 1건에 대한 휴폐업정보를 확인합니다.
' - https://docs.popbill.com/closedown/vb/api#CheckCorpNum
'=========================================================================
Private Sub btnCheckCorpNum_Click()
    Dim CorpState As PBCorpState
    Dim tmp As String
    
    Set CorpState = ClosedownService.CheckCorpNum(txtUserCorpNum.Text, txtCheckCorpNum.Text)
    
    If CorpState Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* taxType (사업자 과세유형) : null-알수없음, 10-일반과세자, 20-면세과세자, 30-간이과세자, 31-간이과세자(세금계산서 발급사업자), 40-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    tmp = tmp + "corpNum (사업자번호) : " + CorpState.CorpNum + vbCrLf
    tmp = tmp + "state (휴폐업상태) : " + CorpState.state + vbCrLf
    tmp = tmp + "taxType (사업자 과세유형) : " + CorpState.taxType + vbCrLf
    tmp = tmp + "typeDate (과세유형 전환일자) : " + CorpState.typeDate + vbCrLf
    tmp = tmp + "stateDate (휴폐업일자) : " + CorpState.stateDate + vbCrLf
    tmp = tmp + "checkDate (국세청 확인일자) : " + CorpState.checkDate
    
    MsgBox tmp, , "휴폐업조회 - 단건"
End Sub

'=========================================================================
' 다수건의 사업자번호에 대한 휴폐업정보를 확인합니다. (최대 1,000건)
' - https://docs.popbill.com/closedown/vb/api#CheckCorpNums
'=========================================================================
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    Dim CorpNumList As New Collection
    Dim tmp As String
    Dim state As PBCorpState
    
    '조회할 사업자번호 배열 (최대 1000건)
    CorpNumList.Add "1234567890"
    CorpNumList.Add "401-03-94930"
    CorpNumList.Add "120-86-75212"
        
    Set resultList = ClosedownService.CheckCorpNums(txtUserCorpNum.Text, CorpNumList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* taxType (사업자 과세유형) : null-알수없음, 10-일반과세자, 20-면세과세자, 30-간이과세자, 31-간이과세자(세금계산서 발급사업자), 40-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    For Each state In resultList
        tmp = tmp + "corpNum (사업자번호) : " + state.CorpNum + vbCrLf
        tmp = tmp + "state (휴폐업상태) : " + state.state + vbCrLf
        tmp = tmp + "taxType (사업자 과세유형) : " + state.taxType + vbCrLf
        tmp = tmp + "typeDate (과세유형 전환일자) : " + state.typeDate + vbCrLf
        tmp = tmp + "stateDate (휴폐업일자) : " + state.stateDate + vbCrLf
        tmp = tmp + "checkDate (국세청 확인일자) : " + state.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp, , "휴폐업조회 - 대량"
End Sub

Private Sub Form_Load()

    '휴폐업조회 모듈 초기화
    ClosedownService.Initialize LinkID, SecretKey
    
    '연동환경설정값, True-개발용 False-상업용
    ClosedownService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True-사용, False-미사용, 기본값(True)
    ClosedownService.IPRestrictOnOff = True
    
    '로컬시스템 시간 사용여부 True-사용, False-미사용, 기본값(False)
    ClosedownService.UseLocalTimeYN = False
End Sub




