Attribute VB_Name = "ReplyAlert"
'==========================================================================
'  메일 할 일 알리미 (Outlook VBA)
'
'  목록 4개 (모두 즐겨찾기에 자동 추가되는 검색 폴더)
'    [내 메일 회신] 빨강 : 내가 보낸 메일에 회신이 온 것       -> 도착 즉시 Windows 알림
'    [내 앞으로 온 메일] 보라 : 본문이 "Hi Jennifer", "김나현 차장님" 처럼 내 이름을 부르며 시작하는 것
'    [내 답장 필요] 파랑 : 나한테 왔는데 아직 회신/확인 안 한 것 -> 정해진 시각(9시, 14시)에 리마인드
'    [답장 대기]    주황 : 내가 보냈는데 N일째 답이 없는 것     -> 하루 한 번 리마인드
'
'  [완료 처리] MarkDone : 선택한 메일을 모든 목록에서 빼고 다시 들어오지 않게 함
'  메일에 회신/전달하면 [내 답장 필요] 에서 자동으로 빠짐
'  상대가 답장하면 [답장 대기] 에서 자동으로 빠짐
'
'  설치 방법은 README.md 참고.
'  ThisOutlookSession 에 아래 코드를 붙여넣어야 자동으로 동작합니다. (ThisOutlookSession.txt)
'
'    Private WithEvents SentItems As Outlook.Items
'
'    Private Sub Application_Startup()
'        Set SentItems = Application.Session.GetDefaultFolder(olFolderSentMail).Items
'        ReplyAlert.OnStartup
'    End Sub
'
'    Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'        ReplyAlert.OnNewMail EntryIDCollection
'    End Sub
'
'    Private Sub SentItems_ItemAdd(ByVal Item As Object)
'        ReplyAlert.OnItemSent Item
'    End Sub
'
'  수동 실행 매크로 (Alt+F8)
'    MarkDone           : 선택한 메일 완료 처리 (모든 목록에서 빠짐)  <- 리본 버튼으로 추가 추천
'    ShowReplyList      : [내 메일 회신] 목록 열기
'    ShowAttentionList  : [내 앞으로 온 메일] 목록 열기
'    ShowMyTodoList     : [내 답장 필요] 목록 열기
'    ShowNoReplyList    : [답장 대기] 목록 열기
'    CheckMyTodo        : [내 답장 필요] 지금 다시 검사
'    CheckNoReply       : [답장 대기] 지금 다시 검사
'    ScanRecentInbox    : 최근 받은 메일에서 내 메일 회신 / 내 앞으로 온 메일 다시 찾기 (기본 14일)
'    TestSelectedMail   : 선택한 메일의 판단 근거 진단
'    TestToast          : Windows 알림 테스트 (메일을 선택하고 실행하면 클릭 시 그 메일이 열림)
'==========================================================================
Option Explicit

'------------------------------ 설정 ------------------------------
' [내 메일 회신] - 내가 보낸 메일에 회신이 오면 알림
Public Const CATEGORY_NAME As String = "내 메일 회신"   ' 범주 (빨간색)
Public Const LIST_FOLDER_NAME As String = "내 메일 회신" ' 목록 폴더 이름
Private Const MATCH_WHOLE_THREAD As Boolean = False ' True = 스레드 안에 내 메일이 있으면 다른 사람 메일에 단 답장도 포함
Private Const CLEAR_REPLY_MARK_WHEN_I_REPLY As Boolean = True ' 그 회신에 내가 다시 답하면 목록에서 빠짐
Private Const SCAN_DAYS As Long = 14                ' ScanRecentInbox 검사 기간(일)
Private Const STARTUP_SCAN_DAYS As Long = 1         ' Outlook 시작 시 검사 기간(일), 0 = 안 함

' [내 앞으로 온 메일] - 본문이 내 이름을 부르며 시작하는 메일
Private Const ATTENTION_ENABLED As Boolean = True
Public Const ATTENTION_CATEGORY As String = "내 앞으로 온 메일"  ' 범주 (보라색)
Public Const ATTENTION_LIST_NAME As String = "내 앞으로 온 메일" ' 목록 폴더 이름
' 내 이름 (세미콜론으로 구분, 대소문자 무시). 인사말 뒤나 줄 맨 앞에 이 이름이 오면 표시
'   예) "Hi Jennifer", "Dear Jennifer,", "Jennifer, please..", "김나현 차장님", "안녕하세요 김 나현 차장님"
Private Const ATTENTION_NAMES As String = "Jennifer;김나현;김 나현"
Private Const ATTENTION_LINES As Long = 5           ' 본문 앞에서 몇 줄까지 볼지 (외부메일 경고 문구 등 고려)
Private Const ATTENTION_TOAST As Boolean = False    ' True = 도착 시 Windows 알림도 띄움
Private Const CLEAR_ATTENTION_WHEN_I_REPLY As Boolean = True ' 내가 회신하면 목록에서 빠짐
Private Const ATTENTION_GREETINGS As String = _
    "hi;hello;hey;dear;good morning;good afternoon;good evening;morning;greetings;attn;attention;to;" & _
    "안녕하세요;안녕하십니까;수신;받는 사람;존경하는"

' [내 답장 필요] - 나한테 왔는데 아직 회신/확인 안 한 메일
Private Const TODO_ENABLED As Boolean = True
Public Const TODO_CATEGORY As String = "내 답장 필요"   ' 범주 (파란색)
Public Const TODO_LIST_NAME As String = "내 답장 필요"  ' 목록 폴더 이름
Private Const TODO_MIN_HOURS As Long = 4            ' 받은 지 몇 시간 지나도 회신 안 했으면 표시
Private Const TODO_LOOKBACK_DAYS As Long = 7        ' 최근 며칠 동안 받은 메일까지 (지나면 목록에서 빠짐)
Private Const TODO_ONLY_TO_ME As Boolean = True     ' True = 받는 사람(To)에 내가 있는 메일만 (참조/CC 제외, 단 내 이름을 부른 메일은 포함)
Private Const TODO_MAX_RECIPIENTS As Long = 10      ' 받는 사람이 이보다 많은 단체 메일은 제외
Private Const TODO_INCLUDE_SUBFOLDERS As Boolean = True ' 받은 편지함의 하위 폴더도 검사
Private Const TODO_REFRESH_MINUTES As Long = 60     ' 목록 자동 갱신 간격(분, 새 메일이 올 때 확인)

' [답장 대기] - 내가 보냈는데 답이 없는 메일
Private Const FOLLOWUP_ENABLED As Boolean = True
Public Const FOLLOWUP_CATEGORY As String = "답장 대기"   ' 범주 (주황색)
Public Const FOLLOWUP_LIST_NAME As String = "답장 대기"  ' 목록 폴더 이름
Private Const FOLLOWUP_DAYS As Long = 3             ' 보낸 지 며칠 지나도 답장 없으면 표시
Private Const FOLLOWUP_LOOKBACK_DAYS As Long = 30   ' 최근 며칠 동안 보낸 메일까지 (지나면 목록에서 빠짐)
Private Const FOLLOWUP_SKIP_MY_REPLIES As Boolean = False ' True = 내가 보낸 "답장(RE:)" 은 추적 안 함

' [알림 / 리마인드]
Private Const REMIND_HOURS As String = "9,14"       ' 리마인드 시각 (쉼표로 구분, "" = 안 함)
Private Const SHOW_TOAST As Boolean = True          ' Windows 알림 표시
Private Const SHOW_POPUP As Boolean = False         ' 회신 도착 시 메시지 박스 (확인 누를 때까지 Outlook 멈춤)
Private Const CLICK_TO_OPEN As Boolean = True       ' 알림 클릭 시 메일/목록 열기
Private Const CREATE_LISTS As Boolean = True        ' 목록 폴더 자동 생성 + 즐겨찾기 추가
Private Const SET_FLAG As Boolean = False           ' 회신 도착 시 플래그(오늘) 추가
Private Const SET_HIGH_IMPORTANCE As Boolean = False ' 회신 도착 시 중요도 "높음"

Private Const USE_CONVERSATION_FALLBACK As Boolean = True ' 대화(스레드) 기반 보조 판단
Private Const SKIP_AUTO_REPLIES As Boolean = True   ' 부재중 자동회신 등은 제외

' 내 주소가 여러 개(별칭)면 세미콜론으로 추가: "a@company.com;b@company.com"
Private Const EXTRA_MY_ADDRESSES As String = ""
'------------------------------------------------------------------

Private Const PR_IN_REPLY_TO_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
Private Const PR_INTERNET_REFERENCES As String = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
Private Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
Private Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
Private Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
Private Const PR_MESSAGE_TO_ME As String = "http://schemas.microsoft.com/mapi/proptag/0x0057000B"
Private Const PR_FLAG_STATUS As String = "http://schemas.microsoft.com/mapi/proptag/0x10900003"

Private Const VERB_REPLY As Long = 102
Private Const VERB_REPLY_ALL As Long = 103
Private Const VERB_FORWARD As Long = 104
Private Const FLAG_COMPLETE As Long = 1

Private Const TOAST_APP_ID As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
Private Const URI_SCHEME As String = "replyalert"   ' replyalert:open/<EntryID>, :list, :attn, :todo, :followup
Private Const DONE_PROP As String = "ReplyAlertNoFollowup" ' 완료 처리 표시 (메일에 숨은 속성으로 저장)
Private Const SETTINGS_APP As String = "ReplyAlert"

Private Const FU_WAITING As Long = 0      ' 답장 대기
Private Const FU_ANSWERED As Long = 1     ' 답장 받음
Private Const FU_SUPERSEDED As Long = 2   ' 같은 스레드에 내가 더 나중에 보낸 메일이 있음
Private Const FU_IGNORED As Long = 3      ' 완료 처리됨

Private m_MyAddresses As String   ' ";a@x.com;b@y.com;" 형태 캐시
Private m_HandlerReady As Boolean ' 클릭 처리기 설치 여부
Private m_Busy As Boolean         ' 검사 중복 실행 방지
Private m_LastTodoScan As Date
Private m_TodoSkip As Object      ' 이미 "답장 불필요" 로 판단한 메일 EntryID (이번 실행 동안 캐시)


'==========================================================================
'  이벤트 진입점 (ThisOutlookSession 에서 호출)
'==========================================================================
Public Sub OnNewMail(ByVal EntryIDCollection As String)
    Dim ids() As String, i As Long, itm As Object
    On Error Resume Next
    ids = Split(EntryIDCollection, ",")
    For i = LBound(ids) To UBound(ids)
        Set itm = Nothing
        Set itm = Application.Session.GetItemFromID(Trim$(ids(i)))
        If Not itm Is Nothing Then
            ProcessMail itm, True
        End If
    Next i
    TryPendingFavorites
    RunScheduledChecks         ' 새 메일이 올 때마다 리마인드 시각/갱신 주기 확인
End Sub

Public Sub OnStartup()
    Dim n As Long
    On Error Resume Next
    If CLICK_TO_OPEN Then InstallClickHandler
    If CREATE_LISTS Then
        EnsureSearchFolder LIST_FOLDER_NAME, olFolderInbox, CATEGORY_NAME, olCategoryColorRed
        If ATTENTION_ENABLED Then EnsureSearchFolder ATTENTION_LIST_NAME, olFolderInbox, ATTENTION_CATEGORY, olCategoryColorPurple
        If TODO_ENABLED Then EnsureSearchFolder TODO_LIST_NAME, olFolderInbox, TODO_CATEGORY, olCategoryColorBlue
        If FOLLOWUP_ENABLED Then EnsureSearchFolder FOLLOWUP_LIST_NAME, olFolderSentMail, FOLLOWUP_CATEGORY, olCategoryColorOrange
    End If
    If STARTUP_SCAN_DAYS > 0 Then
        n = ScanInbox(STARTUP_SCAN_DAYS)
        If n > 0 And SHOW_TOAST Then
            ShowToast "내 메일에 회신 " & n & "건", _
                      "Outlook 이 꺼져 있는 동안 도착한 회신입니다.", _
                      "클릭하면 회신 목록을 엽니다.", _
                      URI_SCHEME & ":list"
        End If
    End If
    RunScheduledChecks
End Sub

' 보낸 편지함에 메일이 들어올 때 (= 내가 메일을 보냈을 때)
Public Sub OnItemSent(ByVal sentItem As Object)
    Dim items As Collection, it As Object, selfId As String, sent As Date, cats As String
    On Error Resume Next
    If sentItem.Class <> olMail Then Exit Sub
    sent = SafeSentOn(sentItem)
    If sent = 0 Then sent = Now
    selfId = SafeEntryId(sentItem)

    Set items = ConversationItems(sentItem)
    For Each it In items
        If IsOtherMail(it, selfId) Then
            If ItemTime(it) < sent Then
                cats = SafeCategories(it)
                ' 내가 답했으니 [내 답장 필요] 에서 제거
                If CatListHas(cats, TODO_CATEGORY) Then
                    RemoveCategoryName it, TODO_CATEGORY, True
                    SkipTodo SafeEntryId(it)
                End If
                ' 회신 받은 메일에 내가 다시 답했으면 [내 메일 회신] 에서 제거
                If CLEAR_REPLY_MARK_WHEN_I_REPLY And CatListHas(cats, CATEGORY_NAME) Then
                    RemoveCategoryName it, CATEGORY_NAME, True
                End If
                ' 내 앞으로 온 메일에 답했으면 [내 앞으로 온 메일] 에서 제거
                If CLEAR_ATTENTION_WHEN_I_REPLY And CatListHas(cats, ATTENTION_CATEGORY) Then
                    RemoveCategoryName it, ATTENTION_CATEGORY, True
                End If
                ' 같은 스레드의 예전 내 메일은 새로 보낸 메일로 대체 -> [답장 대기] 에서 제거
                If CatListHas(cats, FOLLOWUP_CATEGORY) Then
                    If IsFromMe(it) Then RemoveCategoryName it, FOLLOWUP_CATEGORY, True
                End If
            End If
        End If
    Next it
End Sub


'==========================================================================
'  수동 실행 매크로
'==========================================================================
' 선택한 메일 완료 처리: 모든 목록에서 빼고 다시 들어오지 않게 함
Public Sub MarkDone()
    Dim sel As Object, itm As Object, up As Object, n As Long
    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    For Each itm In sel
        If itm.Class = olMail Then
            Set up = Nothing
            Set up = itm.UserProperties.Find(DONE_PROP)
            If up Is Nothing Then Set up = itm.UserProperties.Add(DONE_PROP, olYesNo, False)
            up.Value = True
            RemoveCategoryName itm, CATEGORY_NAME, False
            RemoveCategoryName itm, ATTENTION_CATEGORY, False
            RemoveCategoryName itm, TODO_CATEGORY, False
            RemoveCategoryName itm, FOLLOWUP_CATEGORY, False
            itm.Save
            SkipTodo SafeEntryId(itm)
            n = n + 1
        End If
    Next itm
    If n = 0 Then MsgBox "완료 처리할 메일을 선택한 뒤 실행하세요.", vbExclamation, "메일 알리미"
End Sub

' 예전 이름 호환
Public Sub MarkNoReplyNeeded()
    MarkDone
End Sub

Public Sub ClearReplyMark()
    MarkDone
End Sub

Public Sub ShowReplyList()
    OpenSearchFolder LIST_FOLDER_NAME, olFolderInbox, CATEGORY_NAME, olCategoryColorRed
End Sub

Public Sub ShowAttentionList()
    OpenSearchFolder ATTENTION_LIST_NAME, olFolderInbox, ATTENTION_CATEGORY, olCategoryColorPurple
End Sub

Public Sub ShowMyTodoList()
    OpenSearchFolder TODO_LIST_NAME, olFolderInbox, TODO_CATEGORY, olCategoryColorBlue
End Sub

Public Sub ShowNoReplyList()
    OpenSearchFolder FOLLOWUP_LIST_NAME, olFolderSentMail, FOLLOWUP_CATEGORY, olCategoryColorOrange
End Sub

Public Sub CheckMyTodo()
    Dim total As Long, unread As Long
    If Not TODO_ENABLED Then
        MsgBox "[내 답장 필요] 기능이 꺼져 있습니다. (TODO_ENABLED)", vbExclamation, "메일 알리미"
        Exit Sub
    End If
    total = ScanTodo(unread)
    MsgBox "받은 지 " & TODO_MIN_HOURS & "시간이 지났는데 아직 회신하지 않은 메일: " & total & "건" & vbCrLf & _
           "(그중 안 읽은 메일 " & unread & "건, 최근 " & TODO_LOOKBACK_DAYS & "일 기준)", _
           vbInformation, "내 답장 필요"
    If CREATE_LISTS Then ShowMyTodoList
End Sub

Public Sub CheckNoReply()
    Dim waiting As Long, nNew As Long
    If Not FOLLOWUP_ENABLED Then
        MsgBox "[답장 대기] 기능이 꺼져 있습니다. (FOLLOWUP_ENABLED)", vbExclamation, "메일 알리미"
        Exit Sub
    End If
    SaveSetting SETTINGS_APP, "Followup", "LastRun", Format$(Date, "yyyy-mm-dd")
    nNew = ScanFollowups(waiting)
    MsgBox "보낸 지 " & FOLLOWUP_DAYS & "일이 지났는데 답장이 없는 메일: " & waiting & "건" & vbCrLf & _
           "(최근 " & FOLLOWUP_LOOKBACK_DAYS & "일 동안 보낸 메일 기준, 새로 추가 " & nNew & "건)", _
           vbInformation, "답장 대기"
    If CREATE_LISTS Then ShowNoReplyList
End Sub

Public Sub ScanRecentInbox()
    Dim n As Long, attn As Long
    n = ScanInbox(SCAN_DAYS, attn)
    MsgBox "최근 " & SCAN_DAYS & "일 받은 메일에서 새로 표시했습니다." & vbCrLf & vbCrLf & _
           "  내 메일에 대한 회신: " & n & "건" & vbCrLf & _
           "  내 앞으로 온 메일: " & attn & "건" & vbCrLf & vbCrLf & _
           "(이미 표시된 메일은 제외)", vbInformation, "메일 알리미"
    If CREATE_LISTS Then ShowReplyList
End Sub

' 메일을 하나 선택하고 실행하면, 알림을 클릭했을 때 그 메일이 열리는지까지 확인 가능
Public Sub TestToast()
    Dim itm As Object, openUri As String, subj As String
    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    If Not itm Is Nothing Then
        openUri = URI_SCHEME & ":open/" & itm.EntryID
        subj = itm.Subject
    End If
    On Error GoTo 0
    If Len(openUri) = 0 Then
        ShowToast "내 메일에 회신이 왔습니다", "홍길동 (테스트)", "RE: 알림 테스트 메일", URI_SCHEME & ":list"
    Else
        ShowToast "내 메일에 회신이 왔습니다", "(테스트) 클릭하면 선택한 메일이 열립니다", subj, openUri, _
                  "메일 열기|" & openUri & ";회신 목록 보기|" & URI_SCHEME & ":list"
    End If
End Sub

Public Sub TestSelectedMail()
    Dim itm As Object, msg As String, mine As Object, inReplyTo As String
    Dim st As Long, stText As String, reason As String

    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBox "메일을 하나 선택한 뒤 실행하세요.", vbExclamation, "메일 알리미"
        Exit Sub
    End If
    If itm.Class <> olMail Then
        MsgBox "선택한 항목이 일반 메일이 아닙니다.", vbExclamation, "메일 알리미"
        Exit Sub
    End If

    msg = "제목: " & itm.Subject & vbCrLf & _
          "보낸 사람 주소: " & GetSenderSmtp(itm) & vbCrLf & _
          "내 주소 목록: " & MyAddresses() & vbCrLf & _
          "완료 처리됨: " & YesNo(IsDone(itm)) & vbCrLf & vbCrLf

    If IsFromMe(itm) Then
        st = FollowupState(itm)
        Select Case st
            Case FU_WAITING: stText = "답장 없음 (보낸 지 " & FOLLOWUP_DAYS & "일 지나면 [답장 대기] 에 표시)"
            Case FU_ANSWERED: stText = "답장 받음"
            Case FU_SUPERSEDED: stText = "같은 스레드에 내가 나중에 보낸 메일이 있음 (그 메일 기준으로 판단)"
            Case FU_IGNORED: stText = "완료 처리됨"
        End Select
        msg = msg & "[내가 보낸 메일]" & vbCrLf & _
              "보낸 시각: " & itm.SentOn & vbCrLf & _
              "==> 답장 대기 판단: " & stText
    Else
        inReplyTo = GetPropStr(itm.PropertyAccessor, PR_IN_REPLY_TO_ID)
        Set mine = FindRepliedMine(itm)
        msg = msg & "[받은 메일]" & vbCrLf & _
              "In-Reply-To: " & IIf(Len(inReplyTo) = 0, "(없음)", inReplyTo) & vbCrLf
        If mine Is Nothing Then
            msg = msg & "==> 내 메일에 대한 회신: 아님" & vbCrLf
        Else
            msg = msg & "==> 내 메일에 대한 회신: 맞음 (" & mine.Subject & " / " & mine.SentOn & ")" & vbCrLf
        End If
        stText = AttentionMatch(itm)
        If Len(stText) = 0 Then
            msg = msg & "==> 내 앞으로 온 메일: 아님 (본문 앞 " & ATTENTION_LINES & "줄에 내 이름을 부르는 인사가 없음)" & vbCrLf
        Else
            msg = msg & "==> 내 앞으로 온 메일: 맞음 [" & Left$(stText, 60) & "]" & vbCrLf
        End If
        reason = TodoReason(itm)
        If Len(reason) = 0 Then
            msg = msg & "==> 내 답장 필요: 예 (받은 지 " & TODO_MIN_HOURS & "시간 지나면 목록에 표시)"
        Else
            msg = msg & "==> 내 답장 필요: 아니오 - " & reason
        End If
    End If
    MsgBox msg, vbInformation, "메일 알리미 - 진단"
End Sub


'==========================================================================
'  정기 검사 / 리마인드
'==========================================================================
Private Sub RunScheduledChecks()
    Dim today As String, slotKey As String, doRemind As Boolean, doFollow As Boolean
    Dim todoN As Long, unreadN As Long, fuN As Long, line1 As String, line2 As String, launch As String
    On Error GoTo Done
    If m_Busy Then Exit Sub
    m_Busy = True
    today = Format$(Date, "yyyy-mm-dd")

    ' [답장 대기] 하루 한 번
    If FOLLOWUP_ENABLED Then
        If GetSetting(SETTINGS_APP, "Followup", "LastRun", "") <> today Then
            SaveSetting SETTINGS_APP, "Followup", "LastRun", today
            ScanFollowups fuN
            doFollow = True
        End If
    End If

    ' [내 답장 필요] 리마인드 시각이 되었거나 갱신 주기가 지났으면 검사
    slotKey = CurrentRemindSlot()
    If Len(slotKey) > 0 Then doRemind = (GetSetting(SETTINGS_APP, "Todo", "LastRemind", "") <> slotKey)
    If TODO_ENABLED Then
        If doRemind Or m_LastTodoScan = 0 Or DateDiff("n", m_LastTodoScan, Now) >= TODO_REFRESH_MINUTES Then
            todoN = ScanTodo(unreadN)
        End If
    End If
    If doRemind Then SaveSetting SETTINGS_APP, "Todo", "LastRemind", slotKey

    ' 알림 한 번에 모아서
    If SHOW_TOAST Then
        If doRemind And TODO_ENABLED And todoN > 0 Then
            line1 = "내 답장 필요 " & todoN & "건" & IIf(unreadN > 0, " (안 읽음 " & unreadN & "건)", "")
            launch = URI_SCHEME & ":todo"
        End If
        If doFollow And fuN > 0 Then
            line2 = "답장 대기 " & fuN & "건 (보낸 지 " & FOLLOWUP_DAYS & "일 넘게 답 없음)"
            If Len(launch) = 0 Then launch = URI_SCHEME & ":followup"
        End If
        If Len(line1) = 0 Then
            line1 = line2
            line2 = ""
        End If
        If Len(line1) > 0 Then
            ShowToast "메일 리마인드", line1, line2, launch, _
                      "내 답장 필요|" & URI_SCHEME & ":todo;답장 대기|" & URI_SCHEME & ":followup"
        End If
    End If
Done:
    If Err.Number <> 0 Then Debug.Print "ReplyAlert.RunScheduledChecks 오류: " & Err.Description
    m_Busy = False
End Sub

' 지금 시각 기준으로 가장 최근 리마인드 시각 ("2026-11-26 14"), 없으면 ""
Private Function CurrentRemindSlot() As String
    Dim parts() As String, i As Long, h As Long, best As Long
    best = -1
    If Len(Trim$(REMIND_HOURS)) = 0 Then Exit Function
    parts = Split(REMIND_HOURS, ",")
    For i = LBound(parts) To UBound(parts)
        If IsNumeric(Trim$(parts(i))) Then
            h = CLng(Trim$(parts(i)))
            If h <= Hour(Now) And h > best Then best = h
        End If
    Next i
    If best >= 0 Then CurrentRemindSlot = Format$(Date, "yyyy-mm-dd") & " " & best
End Function


'==========================================================================
'  [내 메일 회신] 핵심 로직
'==========================================================================
' 새로 회신 표시한 개수 반환, attnCount = 새로 [내 앞으로 온 메일] 표시한 개수
Private Function ScanInbox(ByVal days As Long, Optional ByRef attnCount As Long) As Long
    Dim inbox As Object, its As Object, itm As Object, attnMarked As Boolean
    Dim cutoff As Date, rcv As Date, n As Long, total As Long

    cutoff = Now - days
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    Set its = inbox.Items
    its.Sort "[ReceivedTime]", True     ' 최신 메일부터

    For Each itm In its
        rcv = GetReceived(itm)
        If rcv <> 0 And rcv < cutoff Then Exit For
        If ProcessMail(itm, False, attnMarked) Then n = n + 1
        If attnMarked Then attnCount = attnCount + 1
        total = total + 1
        If total Mod 25 = 0 Then DoEvents
    Next itm
    ScanInbox = n
End Function

' 새로 [내 메일 회신] 표시했으면 True. attnMarked = 이번에 [내 앞으로 온 메일] 표시했는지
Private Function ProcessMail(ByVal itm As Object, ByVal notify As Boolean, Optional ByRef attnMarked As Boolean) As Boolean
    Dim mine As Object
    On Error GoTo Fail
    attnMarked = False
    If itm.Class <> olMail Then Exit Function
    If IsDone(itm) Then Exit Function
    If SKIP_AUTO_REPLIES Then
        If IsAutoReply(itm) Then Exit Function
    End If
    If IsFromMe(itm) Then Exit Function

    ' [내 앞으로 온 메일] 본문이 내 이름을 부르며 시작하는지
    If ATTENTION_ENABLED Then
        If Not CatListHas(itm.Categories, ATTENTION_CATEGORY) Then
            If Not IsBulkMail(itm) Then
                If Len(AttentionMatch(itm)) > 0 Then
                    EnsureCategory ATTENTION_CATEGORY, olCategoryColorPurple
                    AddCategoryName itm, ATTENTION_CATEGORY
                    attnMarked = True
                End If
            End If
        End If
    End If

    ' [내 메일 회신] 내가 보낸 메일에 대한 회신인지
    If Not CatListHas(itm.Categories, CATEGORY_NAME) Then
        Set mine = FindRepliedMine(itm)
        If Not mine Is Nothing Then
            MarkMail itm
            If FOLLOWUP_ENABLED Then RemoveCategoryName mine, FOLLOWUP_CATEGORY, True   ' 답장 왔으니 대기 목록에서 제거
            If notify Then NotifyReply itm
            ProcessMail = True
        End If
    End If

    ' 회신 알림을 이미 띄웠으면 중복 알림 안 함
    If notify And attnMarked And Not ProcessMail And ATTENTION_TOAST Then NotifyAttention itm
    Exit Function
Fail:
    Debug.Print "ReplyAlert.ProcessMail 오류: " & Err.Description
End Function

' 받은 메일이 답장한 "내 메일"을 찾아서 반환 (없으면 Nothing)
Private Function FindRepliedMine(ByVal mail As Object) As Object
    Dim pa As Object, inReplyTo As String, refs As String
    Dim ids As Collection, i As Long, idSet As String, found As Object
    Dim items As Collection, it As Object, selfId As String, before As Date, bestTime As Date, t As Date

    Set pa = mail.PropertyAccessor
    inReplyTo = GetPropStr(pa, PR_IN_REPLY_TO_ID)
    If MATCH_WHOLE_THREAD Then refs = GetPropStr(pa, PR_INTERNET_REFERENCES)

    ' 1) 답장 대상 메일 ID 가 보낸 편지함에 있는지
    Set ids = SplitIds(inReplyTo & " " & refs)
    For i = 1 To ids.Count
        Set found = FindSentByMessageId(ids(i))
        If Not found Is Nothing Then
            Set FindRepliedMine = found
            Exit Function
        End If
        idSet = idSet & "|" & LCase$(ids(i))
    Next i
    idSet = idSet & "|"

    If Not USE_CONVERSATION_FALLBACK Then Exit Function

    Set items = ConversationItems(mail)
    selfId = SafeEntryId(mail)
    before = GetReceived(mail)

    For Each it In items
        If IsOtherMail(it, selfId) Then
            If ids.Count > 0 Then
                ' 2) 보낸 편지함에는 없지만(다른 폴더로 옮김 등) 대화 안에 그 메일이 있고 내가 보낸 것인지
                If InStr(1, idSet, "|" & ItemMsgId(it) & "|", vbBinaryCompare) > 0 Then
                    If IsFromMe(it) Then
                        Set FindRepliedMine = it
                        Exit Function
                    End If
                End If
            Else
                ' 3) 헤더가 없으면: 같은 대화에서 이 메일 직전에 내가 보낸 메일
                t = SafeSentOn(it)
                If t <> 0 And t < before And t > bestTime Then
                    If IsFromMe(it) Then
                        Set FindRepliedMine = it
                        bestTime = t
                    End If
                End If
            End If
        End If
    Next it
End Function

Private Function FindSentByMessageId(ByVal msgId As String) As Object
    Dim f As Object, res As Object, flt As String
    On Error GoTo Done
    Set f = Application.Session.GetDefaultFolder(olFolderSentMail)
    flt = "@SQL=""" & PR_INTERNET_MESSAGE_ID & """ = '" & Replace(msgId, "'", "''") & "'"
    Set res = f.Items.Restrict(flt)
    If res.Count > 0 Then Set FindSentByMessageId = res.GetFirst
Done:
End Function


'==========================================================================
'  [내 앞으로 온 메일] 핵심 로직
'==========================================================================
' 본문 앞부분에서 내 이름을 부르는 줄을 찾아 반환 (없으면 "")
Private Function AttentionMatch(ByVal mail As Object) As String
    Dim body As String, lines() As String, i As Long, n As Long, ln As String
    On Error GoTo Done
    If Len(Trim$(ATTENTION_NAMES)) = 0 Then Exit Function
    body = Left$(mail.Body, 3000)
    body = Replace(body, ChrW(160), " ")       ' 줄바꿈 없는 공백
    body = Replace(body, ChrW(8203), "")       ' 폭 없는 공백
    body = Replace(body, ChrW(65279), "")
    body = Replace(body, vbTab, " ")
    lines = Split(body, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(Replace(lines(i), vbCr, ""))
        If Len(ln) > 0 Then
            n = n + 1
            If LineAddressesMe(ln) Then
                AttentionMatch = ln
                Exit Function
            End If
            If n >= ATTENTION_LINES Then Exit For
        End If
    Next i
Done:
End Function

' 한 줄이 나를 부르는 인사인지
'   - 인사말(Hi, Dear, 안녕하세요 ...)을 떼고 나서 줄이 내 이름으로 시작하면 O   "Hi Jennifer," "김나현 차장님,"
'   - 인사말로 시작하는 짧은 줄 안에 내 이름이 있으면 O                          "Hi Tom and Jennifer,"
Private Function LineAddressesMe(ByVal ln As String) As Boolean
    Dim s As String, hadGreeting As Boolean, changed As Boolean
    Dim greetings() As String, names() As String, i As Long, g As String, nm As String

    hadGreeting = (Left$(LTrim$(ln), 1) = "@")      ' @Jennifer 처럼 멘션
    s = StripLead(LCase$(ln))
    greetings = Split(ATTENTION_GREETINGS, ";")
    Do
        changed = False
        For i = LBound(greetings) To UBound(greetings)
            g = Trim$(greetings(i))
            If Len(g) > 0 And Left$(s, Len(g)) = g Then
                If NameBoundary(g, Mid$(s, Len(g) + 1, 1)) Then
                    s = StripLead(Mid$(s, Len(g) + 1))
                    hadGreeting = True
                    changed = True
                    Exit For
                End If
            End If
        Next i
    Loop While changed And Len(s) > 0

    names = Split(ATTENTION_NAMES, ";")
    For i = LBound(names) To UBound(names)
        nm = LCase$(Trim$(names(i)))
        If Len(nm) > 0 Then
            If Left$(s, Len(nm)) = nm Then
                If NameBoundary(nm, Mid$(s, Len(nm) + 1, 1)) Then
                    ' 인사말 없이 이름으로 시작하면, 부르는 말인지 확인
                    ' O "Jennifer, please.." "김나현 차장님," "김나현 차장님 안녕하세요"
                    ' X "Jennifer will send.." "김나현 차장님이 보내주신.."
                    If hadGreeting Or AddressTail(Mid$(s, Len(nm) + 1)) Then
                        LineAddressesMe = True
                        Exit Function
                    End If
                End If
            End If
            If hadGreeting And Len(ln) <= 80 Then
                If ContainsName(s, nm) Then
                    LineAddressesMe = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' 이름 뒤가 부르는 말의 끝인지 (호칭 "차장님", "님께" 뒤에 쉼표/줄끝/안녕하세요)
Private Function AddressTail(ByVal r As String) As Boolean
    Dim p As Long
    r = LTrim$(r)
    p = InStr(1, r, "님", vbBinaryCompare)
    If p > 0 And p <= 6 Then
        r = LTrim$(Mid$(r, p + 1))
        If Left$(r, 1) = "께" And Mid$(r, 2, 1) <> "서" Then r = LTrim$(Mid$(r, 2))
    End If
    If Len(r) = 0 Then AddressTail = True: Exit Function
    If InStr(1, ",.!:;~-)" & ChrW(65292), Left$(r, 1), vbBinaryCompare) > 0 Then AddressTail = True: Exit Function
    If Left$(r, 2) = "안녕" Then AddressTail = True
End Function

' 앞쪽 공백/문장부호/@ 제거
Private Function StripLead(ByVal s As String) As String
    Do While Len(s) > 0
        If InStr(1, " ,.:;-!@~*>" & ChrW(183) & ChrW(65292), Left$(s, 1), vbBinaryCompare) > 0 Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    StripLead = s
End Function

' 영문 단어 뒤에 영문/숫자가 바로 붙어 있으면 다른 단어 (예: "to" vs "today"). 한글은 호칭이 바로 붙으므로 허용
Private Function NameBoundary(ByVal word As String, ByVal nextChar As String) As Boolean
    If Len(nextChar) = 0 Then NameBoundary = True: Exit Function
    If Not (LCase$(Right$(word, 1)) Like "[a-z0-9.]") Then NameBoundary = True: Exit Function
    NameBoundary = Not (LCase$(nextChar) Like "[a-z0-9]")
End Function

Private Function ContainsName(ByVal s As String, ByVal nm As String) As Boolean
    Dim p As Long, prevChar As String
    p = InStr(1, s, nm, vbBinaryCompare)
    Do While p > 0
        If p > 1 Then prevChar = Mid$(s, p - 1, 1) Else prevChar = ""
        If NameBoundary(nm, Mid$(s, p + Len(nm), 1)) And Not (prevChar Like "[a-z0-9]") Then
            ContainsName = True
            Exit Function
        End If
        p = InStr(p + 1, s, nm, vbBinaryCompare)
    Loop
End Function

Private Function IsAttention(ByVal itm As Object) As Boolean
    If CatListHas(SafeCategories(itm), ATTENTION_CATEGORY) Then
        IsAttention = True
    ElseIf ATTENTION_ENABLED Then
        IsAttention = (Len(AttentionMatch(itm)) > 0)
    End If
End Function

Private Sub NotifyAttention(ByVal mail As Object)
    Dim openUri As String
    On Error Resume Next
    openUri = URI_SCHEME & ":open/" & mail.EntryID
    ShowToast "내 앞으로 온 메일", mail.SenderName, mail.Subject, openUri, _
              "메일 열기|" & openUri & ";목록 보기|" & URI_SCHEME & ":attn"
End Sub


'==========================================================================
'  [내 답장 필요] 핵심 로직
'==========================================================================
' 받은 편지함(하위 폴더 포함)을 검사해서 [내 답장 필요] 범주를 붙이거나 뗌. 현재 총수 반환
Private Function ScanTodo(ByRef unreadCount As Long) As Long
    Dim folders As New Collection, fld As Object, total As Long
    On Error Resume Next
    unreadCount = 0
    If m_TodoSkip Is Nothing Then Set m_TodoSkip = CreateObject("Scripting.Dictionary")
    EnsureCategory TODO_CATEGORY, olCategoryColorBlue

    CollectMailFolders Application.Session.GetDefaultFolder(olFolderInbox), folders, 0
    For Each fld In folders
        total = total + ScanTodoFolder(fld, unreadCount)
    Next fld
    m_LastTodoScan = Now
    ScanTodo = total
End Function

Private Sub CollectMailFolders(ByVal fld As Object, ByVal c As Collection, ByVal depth As Long)
    Dim child As Object
    On Error Resume Next
    If fld.DefaultItemType <> olMailItem Then Exit Sub
    c.Add fld
    If Not TODO_INCLUDE_SUBFOLDERS Or depth > 5 Then Exit Sub
    For Each child In fld.Folders
        CollectMailFolders child, c, depth + 1
    Next child
End Sub

Private Function ScanTodoFolder(ByVal fld As Object, ByRef unreadCount As Long) As Long
    Dim tbl As Object, row As Object, itm As Object
    Dim newCut As Date, oldCut As Date, purgeCut As Date, rcv As Date
    Dim eid As String, hasCat As Boolean, need As Boolean, lv As Long, mc As String
    Dim total As Long, cnt As Long

    On Error GoTo Done
    newCut = Now - TODO_MIN_HOURS / 24#
    oldCut = Now - TODO_LOOKBACK_DAYS
    purgeCut = oldCut - 30

    Set tbl = fld.GetTable
    tbl.Columns.Add "ReceivedTime"
    tbl.Columns.Add "UnRead"
    tbl.Columns.Add "Categories"
    tbl.Columns.Add PR_LAST_VERB_EXECUTED
    tbl.Columns.Add PR_FLAG_STATUS
    tbl.Sort "[ReceivedTime]", True

    Do Until tbl.EndOfTable
        Set row = tbl.GetNextRow
        Set itm = Nothing
        cnt = cnt + 1
        If cnt Mod 50 = 0 Then DoEvents

        rcv = RowDate(row, "ReceivedTime")
        If rcv <> 0 And rcv < purgeCut Then Exit Do
        eid = RowStr(row, "EntryID")
        hasCat = CatListHas(RowStr(row, "Categories"), TODO_CATEGORY)

        If rcv = 0 Or rcv > newCut Then GoTo NextRow             ' 아직 기다리는 시간
        If rcv < oldCut Then                                    ' 추적 기간 지남 -> 목록에서 제거
            If hasCat Then RemoveCategoryById eid, TODO_CATEGORY
            GoTo NextRow
        End If

        ' 빠른 확인 (메일을 열지 않고)
        need = True
        mc = LCase$(RowStr(row, "MessageClass"))
        lv = RowLong(row, PR_LAST_VERB_EXECUTED)
        If Left$(mc, 8) <> "ipm.note" Or Left$(mc, 14) = "ipm.note.rules" Then need = False
        If lv = VERB_REPLY Or lv = VERB_REPLY_ALL Or lv = VERB_FORWARD Then need = False
        If RowLong(row, PR_FLAG_STATUS) = FLAG_COMPLETE Then need = False
        If need And Not hasCat Then
            If m_TodoSkip.Exists(eid) Then need = False          ' 예전에 이미 "불필요" 로 판단
        End If

        ' 자세한 확인 (메일 열어서)
        If need Then
            Set itm = GetItemSafe(eid)
            If itm Is Nothing Then
                need = False
            Else
                need = (Len(TodoReason(itm)) = 0)
            End If
        End If

        If need Then
            total = total + 1
            If RowBool(row, "UnRead") Then unreadCount = unreadCount + 1
            If Not hasCat Then AddCategoryName itm, TODO_CATEGORY
        Else
            SkipTodo eid
            If hasCat Then RemoveCategoryById eid, TODO_CATEGORY
        End If
NextRow:
    Loop
Done:
    If Err.Number <> 0 Then Debug.Print "ReplyAlert.ScanTodoFolder 오류: " & Err.Description
    ScanTodoFolder = total
End Function

' 받은 메일이 [내 답장 필요] 인지. 필요하면 "" , 아니면 이유를 반환
Private Function TodoReason(ByVal itm As Object) As String
    Dim lv As Long, n As Long, attn As Boolean
    On Error GoTo Fail
    If itm.Class <> olMail Then TodoReason = "일반 메일 아님": Exit Function
    If Left$(LCase$(itm.MessageClass), 14) = "ipm.note.rules" Then TodoReason = "자동 메일": Exit Function
    If IsDone(itm) Then TodoReason = "완료 처리됨": Exit Function
    If IsFromMe(itm) Then TodoReason = "내가 보낸 메일": Exit Function

    lv = GetPropLong(itm.PropertyAccessor, PR_LAST_VERB_EXECUTED)
    If lv = VERB_REPLY Or lv = VERB_REPLY_ALL Then TodoReason = "이미 회신함": Exit Function
    If lv = VERB_FORWARD Then TodoReason = "이미 전달함": Exit Function
    If GetPropLong(itm.PropertyAccessor, PR_FLAG_STATUS) = FLAG_COMPLETE Then TodoReason = "플래그 완료": Exit Function

    If IsAutoReply(itm) Then TodoReason = "자동회신": Exit Function
    If IsBulkMail(itm) Then TodoReason = "자동 발송/단체 메일": Exit Function

    ' 내 이름을 부른 메일은 참조(CC)나 단체 메일이어도 포함
    attn = IsAttention(itm)
    If TODO_ONLY_TO_ME And Not attn Then
        If Not IsToMe(itm) Then TodoReason = "받는 사람(To)에 내가 없음 (참조/단체 주소)": Exit Function
    End If
    n = itm.Recipients.Count
    If n > TODO_MAX_RECIPIENTS And Not attn Then TodoReason = "받는 사람이 많은 메일 (" & n & "명)": Exit Function
    If IRepliedInThread(itm) Then TodoReason = "같은 스레드에 내가 나중에 보낸 메일 있음": Exit Function
    TodoReason = ""
    Exit Function
Fail:
    TodoReason = "확인 중 오류"
End Function

Private Function IsToMe(ByVal itm As Object) As Boolean
    Dim r As Object, v As Variant
    On Error Resume Next
    v = Empty
    v = itm.PropertyAccessor.GetProperty(PR_MESSAGE_TO_ME)
    If VarType(v) = vbBoolean Then
        If v Then IsToMe = True: Exit Function
    End If
    ' 속성이 없거나 False 면 받는 사람 목록을 직접 확인 (별칭 주소 등)
    For Each r In itm.Recipients
        If r.Type = olTo Then
            If InStr(1, MyAddresses(), ";" & RecipientSmtp(r) & ";", vbTextCompare) > 0 Then
                IsToMe = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function RecipientSmtp(ByVal r As Object) As String
    Dim s As String, eu As Object
    On Error Resume Next
    s = GetPropStr(r.PropertyAccessor, PR_SMTP_ADDRESS)
    If Len(s) = 0 Then
        Set eu = r.AddressEntry.GetExchangeUser
        If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
    End If
    If Len(s) = 0 Then s = r.Address
    RecipientSmtp = LCase$(Trim$(s))
    If Len(RecipientSmtp) = 0 Then RecipientSmtp = "(none)"
End Function

Private Function IsBulkMail(ByVal itm As Object) As Boolean
    Dim hdr As String, snd As String
    On Error Resume Next
    hdr = LCase$(GetPropStr(itm.PropertyAccessor, PR_TRANSPORT_MESSAGE_HEADERS))
    If InStr(1, hdr, "list-unsubscribe:", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, hdr, "precedence: bulk", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, hdr, "precedence: list", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, hdr, "precedence: junk", vbBinaryCompare) > 0 Then IsBulkMail = True
    snd = GetSenderSmtp(itm)
    If InStr(1, snd, "noreply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "no-reply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "donotreply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "do-not-reply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "mailer-daemon", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "postmaster", vbBinaryCompare) > 0 Then IsBulkMail = True
End Function

' 이 메일을 받은 뒤 같은 스레드에 내가 보낸 메일이 있는지 (휴대폰/다른 메일에서 답한 경우)
Private Function IRepliedInThread(ByVal itm As Object) As Boolean
    Dim items As Collection, it As Object, selfId As String, rcv As Date
    rcv = GetReceived(itm)
    If rcv = 0 Then Exit Function
    selfId = SafeEntryId(itm)
    Set items = ConversationItems(itm)
    For Each it In items
        If IsOtherMail(it, selfId) Then
            If SafeSentOn(it) > rcv Then
                If IsFromMe(it) Then
                    IRepliedInThread = True
                    Exit Function
                End If
            End If
        End If
    Next it
End Function

Private Sub SkipTodo(ByVal eid As String)
    On Error Resume Next
    If Len(eid) = 0 Then Exit Sub
    If m_TodoSkip Is Nothing Then Set m_TodoSkip = CreateObject("Scripting.Dictionary")
    m_TodoSkip(eid) = True
End Sub


'==========================================================================
'  [답장 대기] 핵심 로직
'==========================================================================
' 보낸 편지함을 검사해서 [답장 대기] 범주를 붙이거나 뗌. 새로 붙인 개수 반환, waitingTotal = 현재 대기 총수
Private Function ScanFollowups(ByRef waitingTotal As Long) As Long
    Dim replied As Object, tbl As Object, row As Object, itm As Object
    Dim newCut As Date, oldCut As Date, purgeCut As Date, sentOn As Date
    Dim eid As String, msgId As String, hasCat As Boolean, isReply As Boolean
    Dim st As Long, nNew As Long, cnt As Long

    On Error GoTo Done
    waitingTotal = 0
    newCut = Now - FOLLOWUP_DAYS
    oldCut = Now - FOLLOWUP_LOOKBACK_DAYS
    purgeCut = oldCut - 90

    EnsureCategory FOLLOWUP_CATEGORY, olCategoryColorOrange
    Set replied = BuildRepliedIdSet(FOLLOWUP_LOOKBACK_DAYS + 1)

    Set tbl = Application.Session.GetDefaultFolder(olFolderSentMail).GetTable
    tbl.Columns.Add "SentOn"
    tbl.Columns.Add "Categories"
    tbl.Columns.Add PR_INTERNET_MESSAGE_ID
    tbl.Columns.Add PR_IN_REPLY_TO_ID
    tbl.Sort "[SentOn]", True           ' 최근 보낸 메일부터

    Do Until tbl.EndOfTable
        Set row = tbl.GetNextRow
        Set itm = Nothing
        cnt = cnt + 1
        If cnt Mod 50 = 0 Then DoEvents

        sentOn = RowDate(row, "SentOn")
        If sentOn <> 0 And sentOn < purgeCut Then Exit Do
        If Left$(LCase$(RowStr(row, "MessageClass")), 8) <> "ipm.note" Then GoTo NextRow
        If sentOn = 0 Or sentOn >= newCut Then GoTo NextRow      ' 아직 기다리는 기간

        eid = RowStr(row, "EntryID")
        hasCat = CatListHas(RowStr(row, "Categories"), FOLLOWUP_CATEGORY)

        If sentOn < oldCut Then                                 ' 추적 기간이 지난 메일은 목록에서 제거
            If hasCat Then RemoveCategoryById eid, FOLLOWUP_CATEGORY
            GoTo NextRow
        End If

        isReply = (Len(RowStr(row, PR_IN_REPLY_TO_ID)) > 0)
        msgId = NormId(RowStr(row, PR_INTERNET_MESSAGE_ID))

        If Len(msgId) > 0 And replied.Exists(msgId) Then
            st = FU_ANSWERED                                    ' 받은 편지함에 직접 답장이 있음 (빠른 확인)
        ElseIf FOLLOWUP_SKIP_MY_REPLIES And isReply Then
            st = FU_IGNORED
        Else
            Set itm = GetItemSafe(eid)
            st = FollowupState(itm)                             ' 대화(스레드) 전체를 보고 확인
        End If

        If st = FU_WAITING Then
            waitingTotal = waitingTotal + 1
            If Not hasCat Then
                If itm Is Nothing Then Set itm = GetItemSafe(eid)
                If Not itm Is Nothing Then
                    AddCategoryName itm, FOLLOWUP_CATEGORY
                    nNew = nNew + 1
                End If
            End If
        ElseIf hasCat Then
            RemoveCategoryById eid, FOLLOWUP_CATEGORY
        End If
NextRow:
    Loop
Done:
    If Err.Number <> 0 Then Debug.Print "ReplyAlert.ScanFollowups 오류: " & Err.Description
    ScanFollowups = nNew
End Function

' 받은 편지함에서 "누군가 답장한 메일 ID" 목록 (In-Reply-To 값) 을 한 번에 수집
Private Function BuildRepliedIdSet(ByVal days As Long) As Object
    Dim d As Object, tbl As Object, row As Object, cutoff As Date, t As Date, mc As String, id As String
    Set d = CreateObject("Scripting.Dictionary")
    Set BuildRepliedIdSet = d
    On Error GoTo Done
    cutoff = Now - days
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.Add "ReceivedTime"
    tbl.Columns.Add PR_IN_REPLY_TO_ID
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        Set row = tbl.GetNextRow
        t = RowDate(row, "ReceivedTime")
        If t <> 0 And t < cutoff Then Exit Do
        mc = LCase$(RowStr(row, "MessageClass"))
        If Left$(mc, 8) = "ipm.note" And Left$(mc, 14) <> "ipm.note.rules" Then
            id = NormId(RowStr(row, PR_IN_REPLY_TO_ID))
            If Len(id) > 0 Then d(id) = True
        End If
    Loop
Done:
End Function

' 내가 보낸 메일 하나의 답장 대기 상태
Private Function FollowupState(ByVal sentItem As Object) As Long
    Dim items As Collection, it As Object, selfId As String, sent As Date, t As Date
    Dim superseded As Boolean

    FollowupState = FU_ANSWERED
    If sentItem Is Nothing Then Exit Function
    If IsDone(sentItem) Then
        FollowupState = FU_IGNORED
        Exit Function
    End If
    sent = SafeSentOn(sentItem)
    If sent = 0 Then Exit Function
    selfId = SafeEntryId(sentItem)

    Set items = ConversationItems(sentItem)
    For Each it In items
        If IsOtherMail(it, selfId) Then
            t = SafeSentOn(it)
            If t > sent Then
                If IsFromMe(it) Then
                    superseded = True
                ElseIf Not IsAutoReply(it) Then
                    Exit Function                   ' 내 메일 이후 다른 사람 메일이 있음 = 답장 받음
                End If
            End If
        End If
    Next it

    If superseded Then
        FollowupState = FU_SUPERSEDED
    Else
        FollowupState = FU_WAITING
    End If
End Function

' MarkDone 으로 완료 처리한 메일인지
Private Function IsDone(ByVal itm As Object) As Boolean
    Dim up As Object
    On Error Resume Next
    Set up = itm.UserProperties.Find(DONE_PROP)
    If Not up Is Nothing Then
        IsDone = (up.Value = True)
    End If
End Function


'==========================================================================
'  대화(스레드)
'==========================================================================
Private Function ConversationItems(ByVal mail As Object) As Collection
    Dim conv As Object, roots As Object, r As Object, c As New Collection
    On Error GoTo Done
    Set conv = mail.GetConversation
    If conv Is Nothing Then GoTo Done
    Set roots = conv.GetRootItems
    For Each r In roots
        CollectConversation conv, r, c, 0
    Next r
Done:
    Set ConversationItems = c
End Function

Private Sub CollectConversation(ByVal conv As Object, ByVal itm As Object, ByVal c As Collection, ByVal depth As Long)
    Dim kids As Object, k As Object
    On Error GoTo Done
    If depth > 40 Or c.Count > 300 Then Exit Sub
    c.Add itm
    Set kids = conv.GetChildren(itm)
    For Each k In kids
        CollectConversation conv, k, c, depth + 1
    Next k
Done:
End Sub

' 대화 안의 다른 일반 메일인지 (자기 자신 제외)
Private Function IsOtherMail(ByVal it As Object, ByVal selfId As String) As Boolean
    On Error GoTo Done
    If it.Class <> olMail Then Exit Function
    If it.EntryID = selfId Then Exit Function
    IsOtherMail = True
Done:
End Function


'==========================================================================
'  보낸 사람 / 내 주소
'==========================================================================
Private Function IsFromMe(ByVal mail As Object) As Boolean
    Dim a As String
    a = GetSenderSmtp(mail)
    If Len(a) = 0 Then Exit Function
    IsFromMe = (InStr(1, MyAddresses(), ";" & a & ";", vbTextCompare) > 0)
End Function

Private Function GetSenderSmtp(ByVal mail As Object) As String
    Dim s As String, ae As Object, eu As Object
    On Error Resume Next
    If mail.SenderEmailType = "EX" Then
        Set ae = mail.Sender
        If Not ae Is Nothing Then
            Set eu = ae.GetExchangeUser
            If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
        End If
        If Len(s) = 0 Then s = GetPropStr(mail.PropertyAccessor, PR_SENDER_SMTP_ADDRESS)
    Else
        s = mail.SenderEmailAddress
    End If
    GetSenderSmtp = LCase$(Trim$(s))
End Function

Private Function MyAddresses() As String
    Dim s As String, eu As Object, acc As Object
    If Len(m_MyAddresses) > 0 Then
        MyAddresses = m_MyAddresses
        Exit Function
    End If
    On Error Resume Next
    Set eu = Application.Session.CurrentUser.AddressEntry.GetExchangeUser
    If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
    s = s & ";" & Application.Session.CurrentUser.Address
    For Each acc In Application.Session.Accounts
        s = s & ";" & acc.SmtpAddress
    Next acc
    s = s & ";" & EXTRA_MY_ADDRESSES
    m_MyAddresses = ";" & LCase$(Replace(s, " ", "")) & ";"
    MyAddresses = m_MyAddresses
End Function


'==========================================================================
'  표시 / 알림
'==========================================================================
Private Sub MarkMail(ByVal mail As Object)
    EnsureCategory CATEGORY_NAME, olCategoryColorRed
    If Not CatListHas(mail.Categories, CATEGORY_NAME) Then
        If Len(mail.Categories) = 0 Then
            mail.Categories = CATEGORY_NAME
        Else
            mail.Categories = mail.Categories & ", " & CATEGORY_NAME
        End If
    End If
    If SET_FLAG Then mail.MarkAsTask olMarkToday
    If SET_HIGH_IMPORTANCE Then mail.Importance = olImportanceHigh
    mail.Save
End Sub

Private Sub NotifyReply(ByVal mail As Object)
    Dim sender As String, subj As String, openUri As String
    On Error Resume Next
    sender = mail.SenderName
    subj = mail.Subject
    openUri = URI_SCHEME & ":open/" & mail.EntryID
    If SHOW_TOAST Then
        ShowToast "내 메일에 회신이 왔습니다", sender, subj, openUri, _
                  "메일 열기|" & openUri & ";회신 목록 보기|" & URI_SCHEME & ":list"
    End If
    If SHOW_POPUP Then
        MsgBox "내 메일에 회신이 왔습니다." & vbCrLf & vbCrLf & _
               "보낸 사람: " & sender & vbCrLf & "제목: " & subj, _
               vbInformation + vbSystemModal, "메일 알리미"
    End If
End Sub

' Windows 10/11 알림 (PowerShell 사용, 별도 설치 불필요)
'   launchUri : 알림 본문을 클릭했을 때 열 주소
'   buttons   : "버튼이름|주소;버튼이름|주소" (없으면 버튼 없음)
Private Sub ShowToast(ByVal title As String, ByVal line1 As String, ByVal line2 As String, _
                      Optional ByVal launchUri As String = "", Optional ByVal buttons As String = "")
    Dim xml As String, ps As String, head As String, actions As String
    Dim btn() As String, i As Long, kv() As String
    On Error Resume Next

    If CLICK_TO_OPEN And Len(launchUri) > 0 Then
        If Not m_HandlerReady Then InstallClickHandler
    End If

    If CLICK_TO_OPEN And m_HandlerReady And Len(launchUri) > 0 Then
        head = "<toast activationType=""protocol"" launch=""" & XmlEsc(launchUri) & """ duration=""long"">"
        If Len(buttons) > 0 Then
            btn = Split(buttons, ";")
            For i = LBound(btn) To UBound(btn)
                kv = Split(btn(i), "|")
                If UBound(kv) >= 1 Then
                    actions = actions & "<action content=""" & XmlEsc(kv(0)) & """ activationType=""protocol"" arguments=""" & XmlEsc(kv(1)) & """/>"
                End If
            Next i
            If Len(actions) > 0 Then actions = "<actions>" & actions & "</actions>"
        End If
    Else
        head = "<toast duration=""long"">"
    End If

    xml = head & "<visual><binding template=""ToastGeneric"">" & _
          "<text>" & XmlEsc(title) & "</text>" & _
          "<text>" & XmlEsc(line1) & "</text>" & _
          "<text>" & XmlEsc(line2) & "</text>" & _
          "</binding></visual>" & actions & _
          "<audio src=""ms-winsoundevent:Notification.Mail""/></toast>"
    ps = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;" & _
         "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null;" & _
         "$x = New-Object Windows.Data.Xml.Dom.XmlDocument;" & _
         "$x.LoadXml('" & xml & "');" & _
         "$t = New-Object Windows.UI.Notifications.ToastNotification $x;" & _
         "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('" & TOAST_APP_ID & "').Show($t)"
    Shell "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " & B64Utf16(ps), vbHide
End Sub


'==========================================================================
'  검색 폴더 (목록)
'==========================================================================
Private Function FindSearchFolder(ByVal folderName As String) As Object
    Dim f As Object, nm As String
    On Error Resume Next
    For Each f In Application.Session.GetDefaultFolder(olFolderInbox).Store.GetSearchFolders
        nm = ""
        nm = f.Name
        If nm = folderName Then
            Set FindSearchFolder = f
            Exit Function
        End If
    Next f
End Function

' scopeFolderType 폴더(하위 폴더 포함)에서 category 범주가 붙은 메일만 모아 보는 검색 폴더
Private Sub EnsureSearchFolder(ByVal folderName As String, ByVal scopeFolderType As Long, _
                               ByVal category As String, ByVal color As Long)
    Dim scopeFolder As Object, sch As Object, sf As Object, flt As String
    On Error GoTo Done
    If Not FindSearchFolder(folderName) Is Nothing Then Exit Sub
    EnsureCategory category, color
    Set scopeFolder = Application.Session.GetDefaultFolder(scopeFolderType)
    flt = """urn:schemas-microsoft-com:office:office#Keywords"" = '" & Replace(category, "'", "''") & "'"
    Set sch = Application.AdvancedSearch("'" & scopeFolder.FolderPath & "'", flt, True, folderName)
    Set sf = sch.Save(folderName)
    ' 처음 만들 때만 즐겨찾기에 추가 (지워도 다시 추가하지 않음).
    ' 시작 직후라 Outlook 창이 아직 없으면 나중에 다시 시도
    SaveSetting SETTINGS_APP, "Favorites", folderName, "pending"
    TryPendingFavorites
Done:
End Sub

Private Sub TryPendingFavorites()
    Dim names As Variant, i As Long
    On Error Resume Next
    names = Array(LIST_FOLDER_NAME, ATTENTION_LIST_NAME, TODO_LIST_NAME, FOLLOWUP_LIST_NAME)
    For i = LBound(names) To UBound(names)
        If GetSetting(SETTINGS_APP, "Favorites", CStr(names(i)), "") = "pending" Then
            If AddToFavorites(FindSearchFolder(CStr(names(i)))) Then
                SaveSetting SETTINGS_APP, "Favorites", CStr(names(i)), "done"
            End If
        End If
    Next i
End Sub

Private Sub OpenSearchFolder(ByVal folderName As String, ByVal scopeFolderType As Long, _
                             ByVal category As String, ByVal color As Long)
    Dim sf As Object
    On Error Resume Next
    Set sf = FindSearchFolder(folderName)
    If sf Is Nothing Then
        EnsureSearchFolder folderName, scopeFolderType, category, color
        Set sf = FindSearchFolder(folderName)
    End If
    If sf Is Nothing Then
        MsgBox "[" & folderName & "] 목록 폴더를 찾지 못했습니다.", vbExclamation, "메일 알리미"
        Exit Sub
    End If
    If Application.ActiveExplorer Is Nothing Then
        sf.Display
    Else
        Set Application.ActiveExplorer.CurrentFolder = sf
    End If
    TryPendingFavorites
End Sub

' 즐겨찾기에 있거나 추가했으면 True
Private Function AddToFavorites(ByVal fld As Object) As Boolean
    Dim grp As Object, nf As Object
    On Error GoTo Done
    If fld Is Nothing Then Exit Function
    If Application.ActiveExplorer Is Nothing Then Exit Function
    Set grp = Application.ActiveExplorer.NavigationPane.Modules _
                .GetNavigationModule(olModuleMail).NavigationGroups _
                .GetDefaultNavigationGroup(olFavoriteFoldersGroup)
    For Each nf In grp.NavigationFolders
        If nf.Folder.EntryID = fld.EntryID Then
            AddToFavorites = True
            Exit Function
        End If
    Next nf
    grp.NavigationFolders.Add fld
    AddToFavorites = True
Done:
End Function


'==========================================================================
'  알림 클릭 처리기
'  - replyalert: 주소를 현재 사용자 레지스트리(HKCU)에 등록 (관리자 권한 불필요)
'  - 클릭하면 %APPDATA%\ReplyAlert\open.vbs 가 실행되어 Outlook 에서 메일/목록을 엶
'==========================================================================
Private Sub InstallClickHandler()
    Dim fso As Object, ts As Object, sh As Object
    Dim folderPath As String, scriptPath As String, regKey As String
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Environ$("APPDATA") & "\ReplyAlert"
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    scriptPath = folderPath & "\open.vbs"

    Set ts = fso.CreateTextFile(scriptPath, True, True)   ' 유니코드로 저장 (한글 폴더 이름 보존)
    ts.Write HandlerScript()
    ts.Close

    regKey = "HKCU\Software\Classes\" & URI_SCHEME & "\"
    Set sh = CreateObject("WScript.Shell")
    sh.RegWrite regKey, "URL:ReplyAlert Protocol", "REG_SZ"
    sh.RegWrite regKey & "URL Protocol", "", "REG_SZ"
    sh.RegWrite regKey & "shell\open\command\", _
                "wscript.exe //B //Nologo """ & scriptPath & """ ""%1""", "REG_SZ"
    m_HandlerReady = True
Done:
End Sub

' open.vbs 내용 (~ 는 큰따옴표로 바뀜)
Private Function HandlerScript() As String
    Dim s As String
    s = "' ReplyAlert 알림 클릭 처리기 - Outlook 매크로 ReplyAlert 가 자동 생성" & vbCrLf & _
        "On Error Resume Next" & vbCrLf & _
        "Dim uri, arg, ol, ns, itm, f, ex, sh, re, found, target" & vbCrLf & _
        "If WScript.Arguments.Count = 0 Then WScript.Quit" & vbCrLf & _
        "uri = WScript.Arguments(0)" & vbCrLf & _
        "If LCase(Left(uri, <<LEN>>)) = ~<<SCHEME>>:~ Then uri = Mid(uri, <<LEN>> + 1)" & vbCrLf & _
        "Do While Len(uri) > 0 And Right(uri, 1) = ~/~" & vbCrLf & _
        "  uri = Left(uri, Len(uri) - 1)" & vbCrLf & _
        "Loop" & vbCrLf & _
        "Set sh = CreateObject(~WScript.Shell~)" & vbCrLf & _
        "Set ol = GetObject(, ~Outlook.Application~)" & vbCrLf & _
        "If Not IsObject(ol) Then Set ol = CreateObject(~Outlook.Application~)" & vbCrLf & _
        "If Not IsObject(ol) Then WScript.Quit" & vbCrLf & _
        "Set ns = ol.GetNamespace(~MAPI~)" & vbCrLf & _
        "found = False" & vbCrLf
    s = s & _
        "If LCase(Left(uri, 5)) = ~open/~ Then" & vbCrLf & _
        "  arg = Mid(uri, 6)" & vbCrLf & _
        "  Set re = New RegExp" & vbCrLf & _
        "  re.Pattern = ~^[0-9A-Fa-f]+$~" & vbCrLf & _
        "  If re.Test(arg) Then" & vbCrLf & _
        "    Set itm = ns.GetItemFromID(arg)" & vbCrLf & _
        "    If IsObject(itm) Then" & vbCrLf & _
        "      itm.Display" & vbCrLf & _
        "      sh.AppActivate itm.GetInspector.Caption" & vbCrLf & _
        "      found = True" & vbCrLf & _
        "    End If" & vbCrLf & _
        "  End If" & vbCrLf & _
        "End If" & vbCrLf
    s = s & _
        "' 메일을 못 찾았거나(다른 폴더로 옮김 등) 목록 버튼이면 목록 폴더 열기" & vbCrLf & _
        "If Not found Then" & vbCrLf & _
        "  target = ~<<LIST>>~" & vbCrLf & _
        "  If LCase(uri) = ~attn~ Then target = ~<<ATTN>>~" & vbCrLf & _
        "  If LCase(uri) = ~todo~ Then target = ~<<TODO>>~" & vbCrLf & _
        "  If LCase(uri) = ~followup~ Then target = ~<<FOLLOWUP>>~" & vbCrLf & _
        "  For Each f In ns.GetDefaultFolder(6).Store.GetSearchFolders" & vbCrLf & _
        "    If f.Name = target Then" & vbCrLf & _
        "      Set ex = ol.ActiveExplorer" & vbCrLf & _
        "      If ex Is Nothing Then" & vbCrLf & _
        "        f.Display" & vbCrLf & _
        "      Else" & vbCrLf & _
        "        Set ex.CurrentFolder = f" & vbCrLf & _
        "        ex.Activate" & vbCrLf & _
        "      End If" & vbCrLf & _
        "      sh.AppActivate ol.ActiveExplorer.Caption" & vbCrLf & _
        "      Exit For" & vbCrLf & _
        "    End If" & vbCrLf & _
        "  Next" & vbCrLf & _
        "End If" & vbCrLf
    s = Replace(s, "<<SCHEME>>", URI_SCHEME)
    s = Replace(s, "<<LEN>>", CStr(Len(URI_SCHEME) + 1))
    s = Replace(s, "<<LIST>>", Replace(LIST_FOLDER_NAME, "~", ""))
    s = Replace(s, "<<ATTN>>", Replace(ATTENTION_LIST_NAME, "~", ""))
    s = Replace(s, "<<TODO>>", Replace(TODO_LIST_NAME, "~", ""))
    s = Replace(s, "<<FOLLOWUP>>", Replace(FOLLOWUP_LIST_NAME, "~", ""))
    HandlerScript = Replace(s, "~", Chr$(34))
End Function


'==========================================================================
'  범주
'==========================================================================
Private Sub EnsureCategory(ByVal catName As String, ByVal color As Long)
    Dim cat As Object
    On Error Resume Next
    Set cat = Application.Session.Categories.Item(catName)
    If cat Is Nothing Then
        Err.Clear
        Application.Session.Categories.Add catName, color
    End If
End Sub

Private Function CatListHas(ByVal cats As String, ByVal catName As String) As Boolean
    Dim parts() As String, i As Long
    If Len(cats) = 0 Then Exit Function
    parts = Split(Replace(cats, ";", ","), ",")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), catName, vbTextCompare) = 0 Then
            CatListHas = True
            Exit Function
        End If
    Next i
End Function

Private Sub AddCategoryName(ByVal itm As Object, ByVal catName As String)
    On Error Resume Next
    If itm Is Nothing Then Exit Sub
    If CatListHas(itm.Categories, catName) Then Exit Sub
    If Len(itm.Categories) = 0 Then
        itm.Categories = catName
    Else
        itm.Categories = itm.Categories & ", " & catName
    End If
    itm.Save
End Sub

Private Sub RemoveCategoryName(ByVal itm As Object, ByVal catName As String, ByVal doSave As Boolean)
    Dim parts() As String, i As Long, result As String, p As String
    On Error Resume Next
    If itm Is Nothing Then Exit Sub
    If Not CatListHas(itm.Categories, catName) Then Exit Sub
    parts = Split(itm.Categories, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) > 0 And StrComp(p, catName, vbTextCompare) <> 0 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & p
        End If
    Next i
    itm.Categories = result
    If doSave Then itm.Save
End Sub

Private Sub RemoveCategoryById(ByVal entryId As String, ByVal catName As String)
    RemoveCategoryName GetItemSafe(entryId), catName, True
End Sub

Private Function SafeCategories(ByVal itm As Object) As String
    On Error Resume Next
    SafeCategories = itm.Categories
End Function


'==========================================================================
'  유틸리티
'==========================================================================
Private Function IsAutoReply(ByVal mail As Object) As Boolean
    Dim mc As String, hdr As String, subj As String
    On Error Resume Next
    mc = LCase$(mail.MessageClass)
    subj = LCase$(Trim$(mail.Subject))
    hdr = LCase$(GetPropStr(mail.PropertyAccessor, PR_TRANSPORT_MESSAGE_HEADERS))
    If Left$(mc, 14) = "ipm.note.rules" Then IsAutoReply = True
    If InStr(1, hdr, "auto-submitted: auto-", vbBinaryCompare) > 0 Then IsAutoReply = True
    If Left$(subj, 15) = "automatic reply" Then IsAutoReply = True
    If Left$(subj, 13) = "out of office" Then IsAutoReply = True
    If Left$(subj, 5) = "자동 회신" Then IsAutoReply = True
    If Left$(subj, 4) = "자동회신" Then IsAutoReply = True
    If Left$(subj, 3) = "부재중" Then IsAutoReply = True
End Function

Private Function GetPropStr(ByVal pa As Object, ByVal tag As String) As String
    On Error Resume Next
    GetPropStr = CStr(pa.GetProperty(tag))
End Function

Private Function GetPropLong(ByVal pa As Object, ByVal tag As String) As Long
    On Error Resume Next
    GetPropLong = CLng(pa.GetProperty(tag))
End Function

Private Function GetItemSafe(ByVal entryId As String) As Object
    On Error Resume Next
    If Len(entryId) = 0 Then Exit Function
    Set GetItemSafe = Application.Session.GetItemFromID(entryId)
End Function

Private Function GetReceived(ByVal itm As Object) As Date
    On Error Resume Next
    GetReceived = itm.ReceivedTime
End Function

Private Function SafeSentOn(ByVal itm As Object) As Date
    On Error Resume Next
    SafeSentOn = itm.SentOn
End Function

' 받은 메일이면 받은 시각, 아니면 보낸 시각
Private Function ItemTime(ByVal itm As Object) As Date
    ItemTime = GetReceived(itm)
    If ItemTime = 0 Then ItemTime = SafeSentOn(itm)
End Function

Private Function SafeEntryId(ByVal itm As Object) As String
    On Error Resume Next
    SafeEntryId = itm.EntryID
End Function

Private Function RowStr(ByVal row As Object, ByVal col As String) As String
    Dim v As Variant
    On Error Resume Next
    v = row(col)
    If IsArray(v) Then
        RowStr = Join(v, ", ")
    ElseIf Not IsNull(v) And Not IsEmpty(v) Then
        RowStr = CStr(v)
    End If
End Function

Private Function RowDate(ByVal row As Object, ByVal col As String) As Date
    Dim v As Variant
    On Error Resume Next
    v = row(col)
    If IsDate(v) Then RowDate = CDate(v)
End Function

Private Function RowLong(ByVal row As Object, ByVal col As String) As Long
    Dim v As Variant
    On Error Resume Next
    v = row(col)
    If IsNumeric(v) Then RowLong = CLng(v)
End Function

Private Function RowBool(ByVal row As Object, ByVal col As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = row(col)
    If VarType(v) = vbBoolean Then RowBool = v
End Function

' "<a@b> <c@d>" -> Collection("<a@b>", "<c@d>")
Private Function SplitIds(ByVal s As String) As Collection
    Dim c As New Collection, parts() As String, i As Long, p As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ",", " ")
    parts = Split(s, " ")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) > 2 Then
            If Left$(p, 1) <> "<" Then p = "<" & p
            If Right$(p, 1) <> ">" Then p = p & ">"
            c.Add p
        End If
    Next i
    Set SplitIds = c
End Function

Private Function ItemMsgId(ByVal itm As Object) As String
    On Error Resume Next
    ItemMsgId = NormId(GetPropStr(itm.PropertyAccessor, PR_INTERNET_MESSAGE_ID))
    If Len(ItemMsgId) = 0 Then ItemMsgId = "(none)"
End Function

' 메일 ID 를 비교용으로 정리 (첫 번째 ID, 소문자, <> 포함)
Private Function NormId(ByVal s As String) As String
    Dim c As Collection
    Set c = SplitIds(s)
    If c.Count > 0 Then NormId = LCase$(c(1))
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    XmlEsc = s
End Function

' PowerShell -EncodedCommand 용 (UTF-16LE -> Base64)
Private Function B64Utf16(ByVal s As String) As String
    Dim b() As Byte, dom As Object, el As Object
    b = s
    Set dom = CreateObject("MSXML2.DOMDocument")
    Set el = dom.createElement("b64")
    el.DataType = "bin.base64"
    el.nodeTypedValue = b
    B64Utf16 = Replace(Replace(el.Text, vbLf, ""), vbCr, "")
End Function

Private Function YesNo(ByVal v As Boolean) As String
    If v Then YesNo = "예" Else YesNo = "아니오"
End Function
