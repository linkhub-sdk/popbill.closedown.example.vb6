VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "팝빌 휴폐업조회 SDK 예제"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   15180
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton btnGetAccessURL 
      Caption         =   "팝빌 로그인 URL"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   1485
      Width           =   1575
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
      Height          =   1815
      Left            =   600
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
      Begin VB.CommandButton btnJoinMember 
         Caption         =   "회원가입"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckIsMember 
         Caption         =   "가입여부 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton btnCheckID 
         Caption         =   "ID 중복 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "포인트 관련"
      Height          =   1815
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnGetChargeInfo 
         Caption         =   "과금정보 확인"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "팝빌 URL 관련"
      Height          =   1815
      Left            =   8880
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "담당자 관련"
      Height          =   1815
      Left            =   10800
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
      Begin VB.CommandButton btnUpdateContact 
         Caption         =   "담당자 정보 수정"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton btnListContact 
         Caption         =   "담당자 목록 조회"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton btnRegistContact 
         Caption         =   "담당자 추가"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "휴폐업조회"
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   8295
      Begin VB.CommandButton btnCheckCorpNums 
         Caption         =   "대량조회"
         Height          =   495
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton btnCheckCorpNum 
         Caption         =   "단건조회"
         Height          =   495
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCheckCorpNum 
         Height          =   270
         Left            =   2040
         TabIndex        =   13
         Text            =   "6798700433"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "조회할 사업자번호 : "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "팝빌기본 API"
      Height          =   2295
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   14535
      Begin VB.Frame Frame9 
         Caption         =   "파트너과금 포인트"
         Height          =   1815
         Left            =   6240
         TabIndex        =   27
         Top             =   240
         Width           =   2175
         Begin VB.CommandButton btnGetPartnerURL_CHRG 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1935
         End
         Begin VB.CommandButton btnGetPartnerBalance 
            Caption         =   "파트너포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "연동과금 포인트"
         Height          =   1815
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Width           =   2055
         Begin VB.CommandButton btnGetChargeURL 
            Caption         =   "포인트 충전 URL"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton btnGetBalance 
            Caption         =   "잔여포인트 확인"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "회사정보 관련"
         Height          =   1815
         Left            =   12480
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         Begin VB.CommandButton btnUpdateCorpInfo 
            Caption         =   "회사정보 수정"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton btnGetCorpInfo 
            Caption         =   "회사정보 조회"
            Height          =   375
            Left            =   120
            TabIndex        =   23
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
'=========================================================================
'
' 팝빌 휴폐업조회 API VB 6.0 SDK Example
'
' - VB6 SDK 연동환경 설정방법 안내 : https://docs.popbill.com/closedown/tutorial/vb
' - 업데이트 일자 : 2020-02-03
' - 연동 기술지원 연락처 : 1600-9854 / 070-4504-2991
' - 연동 기술지원 이메일 : code@linkhub.co.kr
'
' <테스트 연동개발 준비사항>
' 1) 25, 28번 라인에 선언된 링크아이디(LinkID)와 비밀키(SecretKey)를
'    링크허브 가입시 메일로 발급받은 인증정보를 참조하여 변경합니다.
' 2) 팝빌 개발용 사이트(test.popbill.com)에 연동회원으로 가입합니다.
'=========================================================================

Option Explicit

'=========================================================================
' - 인증정보(링크아이디, 비밀키)는 파트너의 연동회원을 식별하는
'   인증에 사용되는 정보로 유출되지 않도록 주의하시기 바랍니다.
' - 상업용 전환이후에도 인증정보(링크아이디, 비밀키)는 변경되지 않습니다.
'=========================================================================

'링크아이디
Private Const linkID = "TESTER"

'비밀키. 유출에 주의하시기 바랍니다.
Private Const SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="

'휴폐업조회 서비스 객체 생성
Private ClosedownService As New PBCDService

'=========================================================================
' 해당 사업자의 파트너 연동회원 가입여부를 확인합니다.
' - LinkID는 인증정보로 설정되어 있는 링크아이디 값입니다.
' - https://docs.popbill.com/closedown/vb/api#CheckIsMember
'=========================================================================
Private Sub btnCheckIsMember_Click()
    Dim Response As PBResponse
    
    Set Response = ClosedownService.CheckIsMember(txtUserCorpNum.Text, linkID)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 팝빌 회원아이디 중복여부를 확인합니다.
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
' 팝빌 연동회원 가입을 요청합니다.
' - https://docs.popbill.com/closedown/vb/api#JoinMember
'=========================================================================
Private Sub btnJoinMember_Click()
    Dim joinData As New PBJoinForm
    Dim Response As PBResponse
    
    '아이디, 6자이상 50자 미만
    joinData.ID = "userid"
    
    '비밀번호, 6자이상 20자 미만
    joinData.PWD = "pwd_must_be_long_enough"
    
    '파트너링크 아이디
    joinData.linkID = linkID
    
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
    joinData.contactName = "담당자성명"
    
    '담당자 이메일, 최대 100자
    joinData.ContactEmail = "test@test.com"
    
    '담당자 연락처, 최대 20자
    joinData.ContactTEL = "02-999-9999"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.ContactHP = "010-1234-5678"
    
    '담당자 팩스번호, 최대 20자
    joinData.ContactFAX = "02-999-9998"
    
    Set Response = ClosedownService.JoinMember(joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 휴폐업조회 단가를 확인합니다.
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
' 휴폐업조회 API 서비스 과금정보를 확인합니다.
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
' 팝빌에 로그인된 팝업 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
' - https://docs.popbill.com/closedown/vb/api#GetAccessURL
'=========================================================================
Private Sub btnGetAccessURL_Click()
    Dim url As String
    
    url = ClosedownService.GetAccessURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 연동회원의 담당자를 신규로 등록합니다.
' - https://docs.popbill.com/closedown/vb/api#RegistContact
'=========================================================================
Private Sub btnRegistContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디, 6자 이상 50자 미만
    joinData.ID = "testkorea"
    
    '비밀번호, 6자 이상 20자 미만
    joinData.PWD = "test@test.com"
    
    '담당자명, 최대 100자
    joinData.personName = "담당자명"
    
    '담당자 연락처, 최대 20자
    joinData.tel = "070-1234-1234"
    
    '담당자 휴대폰번호, 최대 20자
    joinData.hp = "010-1234-1234"
    
    '담당자 팩스번,최대 20자
    joinData.fax = "070-1234-1234"
    
    '담당자 메일주소, 최대 100자
    joinData.email = "test@test.com"
    
    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
    
    Set Response = ClosedownService.RegistContact(txtUserCorpNum.Text, joinData)
    
    If Response Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox ("응답코드 : " + CStr(Response.code) + vbCrLf + "응답메시지 : " + Response.message)
End Sub

'=========================================================================
' 연동회원의 담당자 목록을 확인합니다.
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
    
    tmp = "id(아이디) | personName(성명) | email(이메일) | hp(휴대폰번호) |  fax(팩스번호) | tel(연락처) | " _
         + "regDT(등록일시) | searchAllAllowYN(회사조회 권한여부) | mgrYN(관리자 여부) | state(상태) " + vbCrLf
    
    For Each info In resultList
        tmp = tmp + info.ID + " | " + info.personName + " | " + info.email + " | " + info.hp + " | " + info.fax _
        + info.tel + " | " + info.regDT + " | " + CStr(info.searchAllAllowYN) + " | " + CStr(info.mgrYN) + " | " + CStr(info.state) + vbCrLf
    Next
    
    MsgBox tmp
End Sub

'=========================================================================
' 연동회원의 담당자 정보를 수정합니다.
' - https://docs.popbill.com/closedown/vb/api#UpdateContact
'=========================================================================
Private Sub btnUpdateContact_Click()
    Dim joinData As New PBContactInfo
    Dim Response As PBResponse
    
    '담당자 아이디
    joinData.ID = txtUserID.Text
    
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

    '회사조회 권한여부, True-회사조회 / False-개인조회
    joinData.searchAllAllowYN = True
    
    '관리자 여부, True-관리자 / False-사용자
    joinData.mgrYN = False
                
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
' - 과금방식이 파트너과금인 경우 파트너 잔여포인트(GetPartnerBalance API)
'   를 통해 확인하시기 바랍니다.
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
' 연동회원 포인트 충전 URL을 반환합니다.
' - 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
' - https://docs.popbill.com/closedown/vb/api#GetChargeURL
'=========================================================================
Private Sub btnGetChargeURL_Click()
    Dim url As String
    
    url = ClosedownService.GetChargeURL(txtUserCorpNum.Text, txtUserID.Text)
    
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
End Sub

'=========================================================================
' 파트너 잔여포인트를 확인합니다.
' - 과금방식이 연동과금인 경우 연동회원 잔여포인트(GetBalance API)
'   를 통해 확인하시기 바랍니다.
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
' 파트너 포인트 충전 URL을 반환합니다.
' - URL 보안정책에 따라 반환된 URL은 30초의 유효시간을 갖습니다.
' - https://docs.popbill.com/closedown/vb/api#GetPartnerURL
'=========================================================================
Private Sub btnGetPartnerURL_CHRG_Click()
    Dim url As String
    
    url = ClosedownService.GetPartnerURL(txtUserCorpNum.Text, "CHRG")
       
    If url = "" Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    MsgBox "URL : " + vbCrLf + url
    
End Sub

'=========================================================================
' 1건의 사업자에 대한 휴폐업 여부를 조회합니다.
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
    tmp = tmp + "* type (사업자 과세유형) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    tmp = tmp + "corpNum (사업자번호) : " + CorpState.CorpNum + vbCrLf
    tmp = tmp + "state (휴폐업상태) : " + CorpState.state + vbCrLf
    tmp = tmp + "type (사업자 과세유형) : " + CorpState.ctype + vbCrLf
    tmp = tmp + "typeDate (과세유형 전환일자) : " + CorpState.typeDate + vbCrLf
    tmp = tmp + "stateDate (휴폐업일자) : " + CorpState.stateDate + vbCrLf
    tmp = tmp + "checkDate (국세청 확인일자) : " + CorpState.checkDate
    
    MsgBox tmp, , "휴폐업조회 - 단건"
End Sub

'=========================================================================
' 사업자에 대한 휴폐업 여부를 조회합니다. (최대 1000건)
' - https://docs.popbill.com/closedown/vb/api#CheckCorpNums
'=========================================================================
Private Sub btnCheckCorpNums_Click()
    Dim resultList As Collection
    Dim CorpNumList As New Collection
    Dim tmp As String
    Dim state As PBCorpState
    
    '조회할 사업자번호 배열 (최대 1000건)
    CorpNumList.Add "1234567890"
    CorpNumList.Add "1231212312"
    CorpNumList.Add "679-87-00433"
        
    Set resultList = ClosedownService.CheckCorpNums(txtUserCorpNum.Text, CorpNumList)
     
    If resultList Is Nothing Then
        MsgBox ("응답코드 : " + CStr(ClosedownService.LastErrCode) + vbCrLf + "응답메시지 : " + ClosedownService.LastErrMessage)
        Exit Sub
    End If
    
    tmp = tmp + "* state (휴폐업상태) : null-알수없음, 0-등록되지 않은 사업자번호, 1-사업중, 2-폐업, 3-휴업" + vbCrLf
    tmp = tmp + "* type (과세유형 전환일자) : null-알수없음, 1-일반과세자, 2-면세과세자, 3-간이과세자, 4-비영리법인, 국가기관" + vbCrLf + vbCrLf
    
    For Each state In resultList
        tmp = tmp + "corpNum (사업자번호) : " + state.CorpNum + vbCrLf
        tmp = tmp + "state (휴폐업상태) : " + state.state + vbCrLf
        tmp = tmp + "type (사업자 과세유형) : " + state.ctype + vbCrLf
        tmp = tmp + "typeDate (과세유형 전환일자) : " + state.typeDate + vbCrLf
        tmp = tmp + "stateDate (휴폐업일자) : " + state.stateDate + vbCrLf
        tmp = tmp + "checkDate (국세청 확인일자) : " + state.checkDate + vbCrLf + vbCrLf
    Next
    
    MsgBox tmp, , "휴폐업조회 - 대량"
End Sub
Private Sub Form_Load()

    '모듈 초기화
    ClosedownService.Initialize linkID, SecretKey
    
    '연동환경 설정값 True(개발용), False(상업용)
    ClosedownService.IsTest = True
    
    '인증토큰 IP제한기능 사용여부, True(권장)
    ClosedownService.IPRestrictOnOff = True
    
End Sub

