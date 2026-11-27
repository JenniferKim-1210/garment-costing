Attribute VB_Name = "ReplyAlert"
'==========================================================================
'  Mail To-Do Reminder for Outlook (ReplyAlert)          - README.md (Korean)
'
'  Lists (search folders, added to Favorites automatically)
'    [red]    replies to mail I sent           -> Windows notification on arrival
'    [purple] mail whose body greets me by name ("Hi Jennifer", Korean name + title)
'    [blue]   mail to me that I have not replied to -> reminder at REMIND_HOURS
'    [orange] mail I sent with no answer after FOLLOWUP_DAYS -> daily reminder
'  MarkDone : remove selected mail from every list for good (adds a "done" category)
'
'  This file contains only ASCII characters, so it works on any Windows language.
'  Korean text is created at run time (see InitStrings).
'  Names that address me: edit %APPDATA%\ReplyAlert\names.txt with Notepad (one per line).
'
'  ThisOutlookSession must contain (see ThisOutlookSession.txt):
'    Private WithEvents SentItems As Outlook.Items
'    Private Sub Application_Startup()
'        Set SentItems = Application.Session.GetDefaultFolder(olFolderSentMail).Items
'        ReplyAlert.OnStartup
'    End Sub
'    Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'        ReplyAlert.OnNewMail EntryIDCollection
'    End Sub
'    Private Sub SentItems_ItemAdd(ByVal Item As Object)
'        ReplyAlert.OnItemSent Item
'    End Sub
'
'  Macros (Alt+F8): MarkDone, ShowReplyList, ShowAttentionList, ShowMyTodoList,
'    ShowNoReplyList, CheckMyTodo, CheckNoReply, ScanRecentInbox, TestSelectedMail,
'    TestToast, EditMyNames
'==========================================================================
Option Explicit

'============================== SETTINGS ==============================
' [Reply to my mail] (red)
Private Const MATCH_WHOLE_THREAD As Boolean = False ' True = also replies to others' mail in a thread I wrote in
Private Const CLEAR_REPLY_MARK_WHEN_I_REPLY As Boolean = True ' remove the mark after I reply to it
Private Const REPLY_TRACK_DAYS As Long = 60         ' alert replies to mail I sent in the last N days
Private Const SCAN_DAYS As Long = 14                ' ScanRecentInbox range (days)
Private Const STARTUP_SCAN_DAYS As Long = 1         ' check mail that arrived while Outlook was closed (days), 0 = off

' [Addressed to me] (purple) - body starts with "Hi Jennifer", "Dear Jennifer", Korean name + title ...
Private Const ATTENTION_ENABLED As Boolean = True
Private Const ATTENTION_EXTRA_NAMES As String = ""  ' extra English names, e.g. "Jenny;Ms. Kim" (Korean names: EditMyNames)
Private Const ATTENTION_LINES As Long = 5           ' how many lines from the top of the body to check
Private Const ATTENTION_TOAST As Boolean = False    ' True = Windows notification on arrival
Private Const CLEAR_ATTENTION_WHEN_I_REPLY As Boolean = True ' remove the mark after I reply

' [Needs my reply] (blue) - mail to me that I have not replied to / read
Private Const TODO_ENABLED As Boolean = True
Private Const TODO_MIN_HOURS As Long = 4            ' show after N hours without reply
Private Const TODO_LOOKBACK_DAYS As Long = 7        ' mail received in the last N days
Private Const TODO_ONLY_TO_ME As Boolean = True     ' True = only when I am in To (not CC), unless the mail greets me by name
Private Const TODO_MAX_RECIPIENTS As Long = 10      ' skip group mail with more recipients (unless it greets me by name)
Private Const TODO_INCLUDE_SUBFOLDERS As Boolean = False ' True = also Inbox subfolders (slower with many folders)
Private Const TODO_REFRESH_MINUTES As Long = 30     ' refresh the list every N minutes (checked when mail arrives)

' [Waiting for answer] (orange) - mail I sent with no answer
Private Const FOLLOWUP_ENABLED As Boolean = True
Private Const FOLLOWUP_DAYS As Long = 3             ' show after N days without answer
Private Const FOLLOWUP_LOOKBACK_DAYS As Long = 30   ' mail sent in the last N days
Private Const FOLLOWUP_SKIP_MY_REPLIES As Boolean = False ' True = track only new mail / forwards, not my replies (RE:)

' [Notifications]
Private Const REMIND_HOURS As String = "9,14"       ' reminder hours, comma separated ("" = off)
Private Const SHOW_TOAST As Boolean = True          ' Windows notifications
Private Const SHOW_POPUP As Boolean = False         ' message box when a reply arrives (Outlook waits until OK)
Private Const CLICK_TO_OPEN As Boolean = True       ' click a notification to open the mail / list
Private Const CREATE_LISTS As Boolean = True        ' create the list folders and add them to Favorites
Private Const SET_FLAG As Boolean = False           ' flag replies (today)
Private Const SET_HIGH_IMPORTANCE As Boolean = False ' set replies to high importance
Private Const SKIP_AUTO_REPLIES As Boolean = True   ' ignore out-of-office replies

' [Performance] keep Outlook responsive
Private Const STARTUP_DELAY_SECONDS As Long = 90    ' do nothing for N seconds after Outlook starts
Private Const SLICE_SECONDS As Single = 1.5         ' work at most N seconds at a time, continue on the next mail

' My other addresses (aliases), e.g. "a@company.com;b@company.com"
Private Const EXTRA_MY_ADDRESSES As String = ""
'======================================================================

' Korean names (set in InitStrings at run time so this file stays ASCII)
Private CATEGORY_NAME As String, LIST_FOLDER_NAME As String
Private ATTENTION_CATEGORY As String, ATTENTION_LIST_NAME As String
Private TODO_CATEGORY As String, TODO_LIST_NAME As String
Private FOLLOWUP_CATEGORY As String, FOLLOWUP_LIST_NAME As String
Private DONE_CATEGORY As String
Private ATTENTION_NAMES As String, ATTENTION_GREETINGS As String
Private APP_TITLE As String
Private m_Inited As Boolean

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If


Private Const PR_IN_REPLY_TO_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
Private Const PR_INTERNET_REFERENCES As String = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
Private Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
Private Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
Private Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"
Private Const PR_MESSAGE_TO_ME As String = "http://schemas.microsoft.com/mapi/proptag/0x0057000B"
Private Const PR_FLAG_STATUS As String = "http://schemas.microsoft.com/mapi/proptag/0x10900003"
Private Const PR_DISPLAY_TO As String = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F"
Private Const PR_DISPLAY_CC As String = "http://schemas.microsoft.com/mapi/proptag/0x0E03001F"
Private Const PR_CONVERSATION_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x30130102"

Private Const VERB_REPLY As Long = 102
Private Const VERB_REPLY_ALL As Long = 103
Private Const VERB_FORWARD As Long = 104
Private Const FLAG_COMPLETE As Long = 1

Private Const TOAST_APP_ID As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
Private Const URI_SCHEME As String = "replyalert"   ' replyalert:open/<EntryID>, :list, :attn, :todo, :followup
Private Const SETTINGS_APP As String = "ReplyAlert"
Private Const TABLE_CHUNK As Long = 200

Private Const FU_WAITING As Long = 0      ' 답장 대기
Private Const FU_ANSWERED As Long = 1     ' 답장 받음
Private Const FU_SUPERSEDED As Long = 2   ' 같은 스레드에 내가 더 나중에 보낸 메일이 있음
Private Const FU_IGNORED As Long = 3      ' 완료 처리됨 / 추적 안 함

' 상태 (Outlook 을 켜 둔 동안만 유지)
Private m_MyAddresses As String   ' ";a@x.com;b@y.com;" 형태
Private m_HandlerReady As Boolean ' 클릭 처리기 설치 여부
Private m_Busy As Boolean         ' 검사 중복 실행 방지
Private m_StartTime As Date
Private m_SliceStart As Single

Private m_MapsTime As Date
Private m_SentIdsRefreshAt As Date
Private m_SentIds As Object        ' 내가 보낸 메일 ID(소문자) -> EntryID
Private m_MyLastTime As Object     ' 대화 ID -> 내가 마지막으로 보낸 시각
Private m_MyLastId As Object       ' 대화 ID -> 내가 마지막으로 보낸 메일 EntryID
Private m_OtherLastTime As Object  ' 대화 ID -> 다른 사람 메일을 마지막으로 받은 시각
Private m_RepliedIds As Object     ' 누군가 답장한 메일 ID(소문자) (받은 메일의 In-Reply-To)
Private m_SmtpCache As Object      ' Exchange 주소 -> SMTP 주소
Private m_TodoChecked As Object    ' EntryID -> "" (답장 필요) 또는 제외 이유
Private m_Seen As Object           ' 이미 처리한 받은 메일 EntryID

Private m_PendingNew As Collection ' 아직 처리 안 한 새 메일 EntryID
Private m_PendingSummary As Boolean
Private m_SumReplies As Long
Private m_CatchUpPending As Boolean
Private m_CatchUpReplies As Long
Private m_TodoDirty As Boolean
Private m_TodoIncomplete As Boolean
Private m_LastTodoScan As Date
Private m_FuDirty As Boolean
Private m_FuToastPending As Boolean
Private m_FuToastCount As Long


'==========================================================================
'  이벤트 진입점 (ThisOutlookSession 에서 호출)
'==========================================================================
Public Sub OnStartup()
    EnsureInit
    On Error Resume Next
    m_StartTime = Now
    m_CatchUpPending = (STARTUP_SCAN_DAYS > 0)
    If CLICK_TO_OPEN Then InstallClickHandler
    If CREATE_LISTS Then
        EnsureSearchFolder LIST_FOLDER_NAME, olFolderInbox, CATEGORY_NAME, olCategoryColorRed
        If ATTENTION_ENABLED Then EnsureSearchFolder ATTENTION_LIST_NAME, olFolderInbox, ATTENTION_CATEGORY, olCategoryColorPurple
        If TODO_ENABLED Then EnsureSearchFolder TODO_LIST_NAME, olFolderInbox, TODO_CATEGORY, olCategoryColorBlue
        If FOLLOWUP_ENABLED Then EnsureSearchFolder FOLLOWUP_LIST_NAME, olFolderSentMail, FOLLOWUP_CATEGORY, olCategoryColorOrange
    End If
    ' 무거운 검사는 여기서 하지 않음 (시작 후 STARTUP_DELAY_SECONDS 뒤, 새 메일이 올 때 조금씩)
End Sub

Public Sub OnNewMail(ByVal EntryIDCollection As String)
    Dim ids() As String, i As Long
    EnsureInit
    On Error Resume Next
    If m_PendingNew Is Nothing Then Set m_PendingNew = New Collection
    ids = Split(EntryIDCollection, ",")
    For i = LBound(ids) To UBound(ids)
        If Len(Trim$(ids(i))) > 0 Then m_PendingNew.Add Trim$(ids(i))
    Next i
    TryPendingFavorites
    RunScheduledChecks
End Sub

' 보낸 편지함에 메일이 들어올 때 (= 내가 메일을 보냈을 때)
Public Sub OnItemSent(ByVal sentItem As Object)
    Dim msgId As String, eid As String, t As Date
    EnsureInit
    On Error Resume Next
    If sentItem.Class <> olMail Then Exit Sub
    If Not m_SentIds Is Nothing Then
        eid = SafeEntryId(sentItem)
        t = SafeSentOn(sentItem)
        If t = 0 Then t = Now
        msgId = NormId(GetPropStr(sentItem.PropertyAccessor, PR_INTERNET_MESSAGE_ID))
        If Len(msgId) > 0 Then
            m_SentIds(msgId) = eid
        Else
            m_SentIdsRefreshAt = Now + TimeSerial(0, 5, 0)  ' 메일 ID 가 아직 없으면 5분 뒤 다시 읽음
        End If
        NoteMine SafeConvId(sentItem), t, eid
    End If
    m_TodoDirty = True     ' 내가 답했으니 [내 답장 필요] 등 정리
    m_FuDirty = True
    RunScheduledChecks
End Sub


'==========================================================================
'  수동 실행 매크로
'==========================================================================
' 선택한 메일 완료 처리: 모든 목록에서 빼고 다시 들어오지 않게 함
Public Sub MarkDone()
    Dim sel As Object, itm As Object, n As Long
    EnsureInit
    On Error Resume Next
    EnsureCategory DONE_CATEGORY, olCategoryColorNone
    EnsureDicts
    Set sel = Application.ActiveExplorer.Selection
    For Each itm In sel
        If itm.Class = olMail Then
            RemoveCategoryName itm, CATEGORY_NAME, False
            RemoveCategoryName itm, ATTENTION_CATEGORY, False
            RemoveCategoryName itm, TODO_CATEGORY, False
            RemoveCategoryName itm, FOLLOWUP_CATEGORY, False
            If Not CatListHas(itm.Categories, DONE_CATEGORY) Then
                If Len(itm.Categories) = 0 Then
                    itm.Categories = DONE_CATEGORY
                Else
                    itm.Categories = itm.Categories & ", " & DONE_CATEGORY
                End If
            End If
            itm.Save
            m_TodoChecked(SafeEntryId(itm)) = "완료 처리됨"
            n = n + 1
        End If
    Next itm
    If n = 0 Then MsgBoxU "완료 처리할 메일을 선택한 뒤 실행하세요.", vbExclamation, APP_TITLE
End Sub

' 예전 이름 호환
Public Sub MarkNoReplyNeeded()
    MarkDone
End Sub

Public Sub ClearReplyMark()
    MarkDone
End Sub

Public Sub ShowReplyList()
    EnsureInit
    OpenSearchFolder LIST_FOLDER_NAME, olFolderInbox, CATEGORY_NAME, olCategoryColorRed
End Sub

Public Sub ShowAttentionList()
    EnsureInit
    OpenSearchFolder ATTENTION_LIST_NAME, olFolderInbox, ATTENTION_CATEGORY, olCategoryColorPurple
End Sub

Public Sub ShowMyTodoList()
    EnsureInit
    OpenSearchFolder TODO_LIST_NAME, olFolderInbox, TODO_CATEGORY, olCategoryColorBlue
End Sub

Public Sub ShowNoReplyList()
    EnsureInit
    OpenSearchFolder FOLLOWUP_LIST_NAME, olFolderSentMail, FOLLOWUP_CATEGORY, olCategoryColorOrange
End Sub

Public Sub CheckMyTodo()
    Dim total As Long, unread As Long
    EnsureInit
    If Not TODO_ENABLED Then
        MsgBoxU "[내 답장 필요] 기능이 꺼져 있습니다. (TODO_ENABLED)", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If m_Busy Then
        MsgBoxU "지금 다른 검사를 하는 중입니다. 잠시 후 다시 실행하세요.", vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    EnsureMaps True
    total = ScanTodo(unread, False)
    m_Busy = False
    MsgBoxU "받은 지 " & TODO_MIN_HOURS & "시간이 지났는데 아직 회신하지 않은 메일: " & total & "건" & vbCrLf & _
           "(그중 안 읽은 메일 " & unread & "건, 최근 " & TODO_LOOKBACK_DAYS & "일 기준)", _
           vbInformation, "내 답장 필요"
    If CREATE_LISTS Then ShowMyTodoList
End Sub

Public Sub CheckNoReply()
    Dim waiting As Long, nNew As Long
    EnsureInit
    If Not FOLLOWUP_ENABLED Then
        MsgBoxU "[답장 대기] 기능이 꺼져 있습니다. (FOLLOWUP_ENABLED)", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If m_Busy Then
        MsgBoxU "지금 다른 검사를 하는 중입니다. 잠시 후 다시 실행하세요.", vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    EnsureMaps True
    SaveSetting SETTINGS_APP, "Followup", "LastRun", Format$(Date, "yyyy-mm-dd")
    nNew = ScanFollowups(waiting)
    m_FuDirty = False
    m_Busy = False
    MsgBoxU "보낸 지 " & FOLLOWUP_DAYS & "일이 지났는데 답장이 없는 메일: " & waiting & "건" & vbCrLf & _
           "(최근 " & FOLLOWUP_LOOKBACK_DAYS & "일 동안 보낸 메일 기준, 새로 추가 " & nNew & "건)", _
           vbInformation, "답장 대기"
    If CREATE_LISTS Then ShowNoReplyList
End Sub

Public Sub ScanRecentInbox()
    Dim n As Long, attn As Long
    EnsureInit
    If m_Busy Then
        MsgBoxU "지금 다른 검사를 하는 중입니다. 잠시 후 다시 실행하세요.", vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    EnsureMaps True
    ScanInboxTable SCAN_DAYS, False, n, attn
    m_Busy = False
    MsgBoxU "최근 " & SCAN_DAYS & "일 받은 메일에서 새로 표시했습니다." & vbCrLf & vbCrLf & _
           "  내 메일에 대한 회신: " & n & "건" & vbCrLf & _
           "  내 앞으로 온 메일: " & attn & "건" & vbCrLf & vbCrLf & _
           "(이미 표시된 메일, 이미 회신한 메일은 제외)", vbInformation, APP_TITLE
    If CREATE_LISTS Then ShowReplyList
End Sub

' 메일을 하나 선택하고 실행하면, 알림을 클릭했을 때 그 메일이 열리는지까지 확인 가능
Public Sub TestToast()
    Dim itm As Object, openUri As String, subj As String
    EnsureInit
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
    EnsureInit

    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBoxU "메일을 하나 선택한 뒤 실행하세요.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    If itm.Class <> olMail Then
        MsgBoxU "선택한 항목이 일반 메일이 아닙니다.", vbExclamation, APP_TITLE
        Exit Sub
    End If
    EnsureMaps False

    msg = "제목: " & itm.Subject & vbCrLf & _
          "보낸 사람 주소: " & GetSenderSmtp(itm) & vbCrLf & _
          "내 주소 목록: " & MyAddresses() & vbCrLf & _
          "완료 처리됨: " & YesNo(IsDone(itm)) & vbCrLf & _
          "대화 ID 있음: " & YesNo(Len(SafeConvId(itm)) > 0) & vbCrLf & vbCrLf

    If IsFromMe(itm) Then
        st = FollowupState(itm)
        Select Case st
            Case FU_WAITING: stText = "답장 없음 (보낸 지 " & FOLLOWUP_DAYS & "일 지나면 [답장 대기] 에 표시)"
            Case FU_ANSWERED: stText = "답장 받음"
            Case FU_SUPERSEDED: stText = "같은 스레드에 내가 나중에 보낸 메일이 있음 (그 메일 기준으로 판단)"
            Case FU_IGNORED: stText = "완료 처리됨 / 추적 안 함"
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
    MsgBoxU msg, vbInformation, APP_TITLE & " - " & "진단"
End Sub


'==========================================================================
'  정기 검사 / 리마인드 (새 메일이 오거나 메일을 보낼 때마다 조금씩)
'==========================================================================
Private Sub RunScheduledChecks()
    Dim today As String, slotKey As String, doRemind As Boolean, remindNow As Boolean, todoRan As Boolean
    Dim todoN As Long, unreadN As Long, fuN As Long, dailyDue As Boolean
    Dim line1 As String, line2 As String, launch As String
    On Error GoTo Done
    If m_Busy Then Exit Sub
    If m_StartTime <> 0 Then
        If DateDiff("s", m_StartTime, Now) < STARTUP_DELAY_SECONDS Then Exit Sub
    End If
    m_Busy = True
    StartSlice
    EnsureMaps False

    ' 0) 예전 버전이 깨진 이름으로 만든 범주/폴더 정리 (한 번만)
    If GetSetting(SETTINGS_APP, "Cleanup", "Garbled", "") = "" Then
        SaveSetting SETTINGS_APP, "Cleanup", "Garbled", "done"
        CleanupGarbled
        GoTo Done
    End If

    ' 1) 새로 도착한 메일
    If Not ProcessPendingNew() Then GoTo Done

    ' 2) Outlook 이 꺼져 있는 동안 온 메일 (시작 후 한 번)
    If m_CatchUpPending Then
        If Not CatchUpStep() Then GoTo Done
    End If
    If SliceOver() Then GoTo Done

    today = Format$(Date, "yyyy-mm-dd")

    ' 3) [답장 대기] 하루 한 번 + 변화가 있을 때 (목록 정보만 보므로 빠름)
    If FOLLOWUP_ENABLED Then
        dailyDue = (GetSetting(SETTINGS_APP, "Followup", "LastRun", "") <> today)
        If dailyDue Or m_FuDirty Then
            ScanFollowups fuN
            m_FuDirty = False
            If dailyDue Then
                SaveSetting SETTINGS_APP, "Followup", "LastRun", today
                m_FuToastCount = fuN
                m_FuToastPending = (fuN > 0)
            End If
        End If
    End If
    If SliceOver() Then GoTo Done

    ' 4) [내 답장 필요] 리마인드 시각 / 갱신 주기 / 변화가 있을 때
    slotKey = CurrentRemindSlot()
    If Len(slotKey) > 0 Then doRemind = (GetSetting(SETTINGS_APP, "Todo", "LastRemind", "") <> slotKey)
    If TODO_ENABLED Then
        If doRemind Or m_TodoDirty Or m_TodoIncomplete Or m_LastTodoScan = 0 Then
            todoRan = True
        ElseIf DateDiff("n", m_LastTodoScan, Now) >= TODO_REFRESH_MINUTES Then
            todoRan = True
        End If
        If todoRan Then
            todoN = ScanTodo(unreadN, True)
            m_TodoDirty = False
        End If
    End If
    remindNow = doRemind And (todoRan Or Not TODO_ENABLED) And Not m_TodoIncomplete
    If remindNow Then SaveSetting SETTINGS_APP, "Todo", "LastRemind", slotKey

    ' 5) 알림 한 번에 모아서
    If SHOW_TOAST Then
        If remindNow And TODO_ENABLED And todoN > 0 Then
            line1 = "내 답장 필요 " & todoN & "건" & IIf(unreadN > 0, " (안 읽음 " & unreadN & "건)", "")
            launch = URI_SCHEME & ":todo"
        End If
        If m_FuToastPending And (remindNow Or Not doRemind) Then
            line2 = "답장 대기 " & m_FuToastCount & "건 (보낸 지 " & FOLLOWUP_DAYS & "일 넘게 답 없음)"
            If Len(launch) = 0 Then launch = URI_SCHEME & ":followup"
            m_FuToastPending = False
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

' 새 메일 처리. 다 처리했으면 True
Private Function ProcessPendingNew() As Boolean
    Dim eid As String, itm As Object, attn As Boolean, marked As Boolean
    On Error GoTo Done
    If m_PendingNew Is Nothing Then ProcessPendingNew = True: Exit Function
    If m_PendingNew.Count > 3 Then m_PendingSummary = True     ' 한꺼번에 많이 오면 알림을 모아서 한 번만
    Do While m_PendingNew.Count > 0
        eid = m_PendingNew(1)
        m_PendingNew.Remove 1
        m_Seen(eid) = True
        Set itm = GetItemSafe(eid)
        If Not itm Is Nothing Then
            UpdateMapsWithItem itm
            marked = ProcessMail(itm, Not m_PendingSummary, attn)
            If marked And m_PendingSummary Then m_SumReplies = m_SumReplies + 1
        End If
        If SliceOver() Then Exit Function
    Loop
    If m_PendingSummary Then
        m_PendingSummary = False
        If m_SumReplies > 0 And SHOW_TOAST Then
            ShowToast "내 메일에 회신 " & m_SumReplies & "건", _
                      "한꺼번에 도착한 메일 중 내 메일에 대한 회신입니다.", _
                      "클릭하면 회신 목록을 엽니다.", URI_SCHEME & ":list"
        End If
        m_SumReplies = 0
    End If
    ProcessPendingNew = True
Done:
End Function

' Outlook 이 꺼져 있는 동안 온 메일. 다 처리했으면 True
Private Function CatchUpStep() As Boolean
    Dim reps As Long, attns As Long
    If ScanInboxTable(STARTUP_SCAN_DAYS, True, reps, attns) Then
        m_CatchUpReplies = m_CatchUpReplies + reps
        m_CatchUpPending = False
        If m_CatchUpReplies > 0 And SHOW_TOAST Then
            ShowToast "내 메일에 회신 " & m_CatchUpReplies & "건", _
                      "Outlook 이 꺼져 있는 동안 도착한 회신입니다.", _
                      "클릭하면 회신 목록을 엽니다.", URI_SCHEME & ":list"
        End If
        m_CatchUpReplies = 0
        CatchUpStep = True
    Else
        m_CatchUpReplies = m_CatchUpReplies + reps
    End If
End Function

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

Private Sub StartSlice()
    m_SliceStart = Timer
End Sub

' 이번 차례에 쓸 시간을 다 썼는지
Private Function SliceOver() As Boolean
    Dim t As Single
    t = Timer
    If t < m_SliceStart Then t = t + 86400!
    SliceOver = ((t - m_SliceStart) > SLICE_SECONDS)
End Function


'==========================================================================
'  메일함 요약 정보 (대화별 마지막 시각, 보낸 메일 ID) - 목록(Table)으로 한 번에 읽음
'==========================================================================
Private Sub EnsureDicts()
    If m_SentIds Is Nothing Then Set m_SentIds = NewDict()
    If m_MyLastTime Is Nothing Then Set m_MyLastTime = NewDict()
    If m_MyLastId Is Nothing Then Set m_MyLastId = NewDict()
    If m_OtherLastTime Is Nothing Then Set m_OtherLastTime = NewDict()
    If m_RepliedIds Is Nothing Then Set m_RepliedIds = NewDict()
    If m_SmtpCache Is Nothing Then Set m_SmtpCache = NewDict()
    If m_TodoChecked Is Nothing Then Set m_TodoChecked = NewDict()
    If m_Seen Is Nothing Then Set m_Seen = NewDict()
End Sub

Private Sub EnsureMaps(ByVal force As Boolean)
    EnsureDicts
    If force Or m_MapsTime = 0 Or DateDiff("h", m_MapsTime, Now) >= 6 Then
        BuildMaps
    ElseIf m_SentIdsRefreshAt <> 0 And Now >= m_SentIdsRefreshAt Then
        m_SentIdsRefreshAt = 0
        AddMapsFromSent Now - REPLY_TRACK_DAYS
    End If
End Sub

Private Sub BuildMaps()
    Dim days As Long, folders As New Collection, fld As Object
    On Error Resume Next
    Set m_SentIds = NewDict()
    Set m_MyLastTime = NewDict()
    Set m_MyLastId = NewDict()
    Set m_OtherLastTime = NewDict()
    Set m_RepliedIds = NewDict()

    days = MaxL(MaxL(FOLLOWUP_LOOKBACK_DAYS, TODO_LOOKBACK_DAYS), SCAN_DAYS) + 2
    AddMapsFromSent Now - MaxL(REPLY_TRACK_DAYS, days)
    CollectInboxFolders folders
    For Each fld In folders
        AddMapsFromReceived fld, Now - days
    Next fld
    AddMapsFromReceived Application.Session.GetDefaultFolder(olFolderDeletedItems), Now - days
    m_MapsTime = Now
    m_SentIdsRefreshAt = 0
End Sub

Private Sub AddMapsFromSent(ByVal cutoff As Date)
    Dim C_EID As Long, C_CLASS As Long, C_TIME As Long, C_MSGID As Long, C_CONV As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, msgId As String, eid As String
    On Error GoTo Done
    Set tbl = Application.Session.GetDefaultFolder(olFolderSentMail).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_TIME = AddCol(tbl, "SentOn")
    C_MSGID = AddCol(tbl, PR_INTERNET_MESSAGE_ID)
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    If C_EID < 0 Or C_TIME < 0 Then GoTo Done
    tbl.Sort "[SentOn]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_TIME))
            If t <> 0 And t < cutoff Then GoTo Done
            If Left$(RowClass(arr, r, C_CLASS), 8) = "ipm.note" Then
                eid = VStr(AV(arr, r, C_EID))
                msgId = NormId(VStr(AV(arr, r, C_MSGID)))
                If Len(msgId) > 0 Then m_SentIds(msgId) = eid
                NoteMine VHex(AV(arr, r, C_CONV)), t, eid
            End If
        Next r
        DoEvents
    Loop
Done:
End Sub

Private Sub AddMapsFromReceived(ByVal fld As Object, ByVal cutoff As Date)
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_SENT As Long
    Dim C_SENDER As Long, C_IRT As Long, C_CONV As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, mc As String, conv As String, irt As String
    On Error GoTo Done
    Set tbl = fld.GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_SENT = AddCol(tbl, "SentOn")
    C_SENDER = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    C_IRT = AddCol(tbl, PR_IN_REPLY_TO_ID)
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    If C_EID < 0 Or C_RCV < 0 Then GoTo Done
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Done
            mc = RowClass(arr, r, C_CLASS)
            If Left$(mc, 8) = "ipm.note" And Left$(mc, 14) <> "ipm.note.rules" Then
                conv = VHex(AV(arr, r, C_CONV))
                If IsMeAddress(VStr(AV(arr, r, C_SENDER))) Then
                    NoteMine conv, VDate(AV(arr, r, C_SENT)), VStr(AV(arr, r, C_EID))
                Else
                    NoteOther conv, t
                    irt = NormId(VStr(AV(arr, r, C_IRT)))
                    If Len(irt) > 0 Then m_RepliedIds(irt) = True
                End If
            End If
        Next r
        DoEvents
    Loop
Done:
End Sub

Private Sub UpdateMapsWithItem(ByVal itm As Object)
    Dim conv As String, irt As String
    On Error Resume Next
    If itm.Class <> olMail Then Exit Sub
    conv = SafeConvId(itm)
    If IsFromMe(itm) Then
        NoteMine conv, SafeSentOn(itm), SafeEntryId(itm)
    Else
        NoteOther conv, ItemTime(itm)
        irt = NormId(GetPropStr(itm.PropertyAccessor, PR_IN_REPLY_TO_ID))
        If Len(irt) > 0 Then m_RepliedIds(irt) = True
        If Len(conv) > 0 Then
            If m_MyLastTime.Exists(conv) Then m_FuDirty = True   ' 내 메일 스레드에 답이 옴
        End If
    End If
End Sub

Private Sub NoteMine(ByVal conv As String, ByVal t As Date, ByVal eid As String)
    If Len(conv) = 0 Or t = 0 Then Exit Sub
    If m_MyLastTime.Exists(conv) Then
        If m_MyLastTime(conv) >= t Then Exit Sub
    End If
    m_MyLastTime(conv) = t
    m_MyLastId(conv) = eid
End Sub

Private Sub NoteOther(ByVal conv As String, ByVal t As Date)
    If Len(conv) = 0 Or t = 0 Then Exit Sub
    If m_OtherLastTime.Exists(conv) Then
        If m_OtherLastTime(conv) >= t Then Exit Sub
    End If
    m_OtherLastTime(conv) = t
End Sub

' 이 메일을 받은 뒤 같은 스레드에 내가 보낸 메일이 있는지 (휴대폰/다른 메일에서 답한 경우 포함)
Private Function IRepliedInThread(ByVal conv As String, ByVal rcv As Date) As Boolean
    If Len(conv) = 0 Or rcv = 0 Then Exit Function
    If m_MyLastTime Is Nothing Then Exit Function
    If m_MyLastTime.Exists(conv) Then IRepliedInThread = (m_MyLastTime(conv) > rcv)
End Function

Private Sub CollectInboxFolders(ByVal c As Collection)
    Dim inbox As Object
    On Error Resume Next
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    c.Add inbox
    If TODO_INCLUDE_SUBFOLDERS Then AddSubFolders inbox, c, 0
End Sub

Private Sub AddSubFolders(ByVal fld As Object, ByVal c As Collection, ByVal depth As Long)
    Dim child As Object
    On Error Resume Next
    If depth > 5 Then Exit Sub
    For Each child In fld.Folders
        If child.DefaultItemType = olMailItem Then
            c.Add child
            AddSubFolders child, c, depth + 1
        End If
    Next child
End Sub


'==========================================================================
'  [내 메일 회신] / [내 앞으로 온 메일] - 받은 메일 한 통 처리
'==========================================================================
' 새로 [내 메일 회신] 표시했으면 True. attnMarked = 이번에 [내 앞으로 온 메일] 표시했는지
Private Function ProcessMail(ByVal itm As Object, ByVal notify As Boolean, Optional ByRef attnMarked As Boolean) As Boolean
    Dim mine As Object, lv As Long
    On Error GoTo Fail
    attnMarked = False
    If itm.Class <> olMail Then Exit Function
    If IsDone(itm) Then Exit Function
    If IsFromMe(itm) Then Exit Function
    If SKIP_AUTO_REPLIES Then
        If IsAutoReply(itm) Then Exit Function
    End If

    ' 이미 내가 답한 메일이면 표시하지 않음 (예전 메일을 다시 검사할 때)
    lv = GetPropLong(itm.PropertyAccessor, PR_LAST_VERB_EXECUTED)
    If lv = VERB_REPLY Or lv = VERB_REPLY_ALL Or lv = VERB_FORWARD Then Exit Function
    If IRepliedInThread(SafeConvId(itm), GetReceived(itm)) Then Exit Function

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

' 받은 편지함을 목록(Table)으로 훑으며 ProcessMail. sliced = 시간 나눠서. 다 끝났으면 True
Private Function ScanInboxTable(ByVal days As Long, ByVal sliced As Boolean, _
                                ByRef replies As Long, ByRef attns As Long) As Boolean
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date
    Dim eid As String, itm As Object, attn As Boolean, cnt As Long
    On Error GoTo Fail
    EnsureDicts
    cutoff = Now - days
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    If C_EID < 0 Or C_RCV < 0 Then GoTo Finished
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Finished
            eid = VStr(AV(arr, r, C_EID))
            If Left$(RowClass(arr, r, C_CLASS), 8) <> "ipm.note" Then GoTo NextRow
            If sliced Then
                If m_Seen.Exists(eid) Then GoTo NextRow
                If SliceOver() Then Exit Function          ' 다음에 이어서
            End If
            Set itm = GetItemSafe(eid)
            If Not itm Is Nothing Then
                If ProcessMail(itm, False, attn) Then replies = replies + 1
                If attn Then attns = attns + 1
            End If
            m_Seen(eid) = True
            cnt = cnt + 1
            If cnt Mod 10 = 0 Then DoEvents
NextRow:
        Next r
    Loop
Finished:
    ScanInboxTable = True
    Exit Function
Fail:
    Debug.Print "ReplyAlert.ScanInboxTable 오류: " & Err.Description
    ScanInboxTable = True
End Function

' 받은 메일이 답장한 "내 메일"을 찾아서 반환 (없으면 Nothing)
Private Function FindRepliedMine(ByVal mail As Object) As Object
    Dim pa As Object, inReplyTo As String, refs As String, ids As Collection, i As Long
    Dim key As String, f As Object, conv As String
    On Error GoTo Done
    EnsureMaps False
    Set pa = mail.PropertyAccessor
    inReplyTo = GetPropStr(pa, PR_IN_REPLY_TO_ID)
    If MATCH_WHOLE_THREAD Then refs = GetPropStr(pa, PR_INTERNET_REFERENCES)

    ' 1) 답장 대상 메일 ID 가 내가 보낸 메일인지
    Set ids = SplitIds(inReplyTo & " " & refs)
    For i = 1 To ids.Count
        key = LCase$(ids(i))
        If m_SentIds.Exists(key) Then
            Set f = GetItemSafe(m_SentIds(key))
            If f Is Nothing Then Set f = FindSentByMessageId(key)   ' 다른 폴더로 옮긴 경우
            If Not f Is Nothing Then
                Set FindRepliedMine = f
                Exit Function
            End If
        End If
    Next i

    ' 2) 답장 정보가 없는 메일: 같은 대화에서 이 메일 전에 내가 보낸 메일이 있으면
    If ids.Count = 0 Then
        conv = SafeConvId(mail)
        If Len(conv) > 0 Then
            If m_MyLastTime.Exists(conv) Then
                If m_MyLastTime(conv) < GetReceived(mail) Then Set FindRepliedMine = GetItemSafe(m_MyLastId(conv))
            End If
        End If
    End If
Done:
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
'  [내 앞으로 온 메일] 판단
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

Private Sub NotifyAttention(ByVal mail As Object)
    Dim openUri As String
    On Error Resume Next
    openUri = URI_SCHEME & ":open/" & mail.EntryID
    ShowToast "내 앞으로 온 메일", mail.SenderName, mail.Subject, openUri, _
              "메일 열기|" & openUri & ";목록 보기|" & URI_SCHEME & ":attn"
End Sub


'==========================================================================
'  [내 답장 필요]
'==========================================================================
' 받은 편지함을 검사해서 [내 답장 필요] 범주를 붙이거나 뗌. 현재 총수 반환
Private Function ScanTodo(ByRef unreadCount As Long, ByVal sliced As Boolean) As Long
    Dim folders As New Collection, fld As Object, total As Long
    On Error Resume Next
    unreadCount = 0
    EnsureDicts
    EnsureCategory TODO_CATEGORY, olCategoryColorBlue
    m_TodoIncomplete = False
    CollectInboxFolders folders
    For Each fld In folders
        total = total + ScanTodoFolder(fld, unreadCount, sliced)
        If m_TodoIncomplete Then Exit For
    Next fld
    m_LastTodoScan = Now
    ScanTodo = total
End Function

Private Function ScanTodoFolder(ByVal fld As Object, ByRef unreadCount As Long, ByVal sliced As Boolean) As Long
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_UNREAD As Long, C_CATS As Long
    Dim C_VERB As Long, C_FLAG As Long, C_TOME As Long, C_TO As Long, C_CC As Long
    Dim C_SENDER As Long, C_CONV As Long
    Dim tbl As Object, arr As Variant, r As Long, itm As Object
    Dim newCut As Date, oldCut As Date, purgeCut As Date, rcv As Date
    Dim eid As String, cats As String, mc As String, reason As String, toMe As Variant
    Dim hasTodo As Boolean, hasReply As Boolean, hasAttn As Boolean, handled As Boolean, need As Boolean
    Dim lv As Long, total As Long

    On Error GoTo Done
    newCut = Now - TODO_MIN_HOURS / 24#
    oldCut = Now - TODO_LOOKBACK_DAYS
    purgeCut = Now - (MaxL(TODO_LOOKBACK_DAYS, SCAN_DAYS) + 30)

    Set tbl = fld.GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_UNREAD = AddCol(tbl, "UnRead")
    C_CATS = AddCol(tbl, "Categories")
    C_VERB = AddCol(tbl, PR_LAST_VERB_EXECUTED)
    C_FLAG = AddCol(tbl, PR_FLAG_STATUS)
    C_TOME = AddCol(tbl, PR_MESSAGE_TO_ME)
    C_TO = AddCol(tbl, PR_DISPLAY_TO)
    C_CC = AddCol(tbl, PR_DISPLAY_CC)
    C_SENDER = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    If C_EID < 0 Or C_RCV < 0 Then GoTo Done
    tbl.Sort "[ReceivedTime]", True

    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            Set itm = Nothing
            rcv = VDate(AV(arr, r, C_RCV))
            If rcv <> 0 And rcv < purgeCut Then GoTo Done
            mc = RowClass(arr, r, C_CLASS)
            If Left$(mc, 8) <> "ipm.note" Or Left$(mc, 14) = "ipm.note.rules" Then GoTo NextRow

            eid = VStr(AV(arr, r, C_EID))
            cats = VStr(AV(arr, r, C_CATS))
            hasTodo = CatListHas(cats, TODO_CATEGORY)
            hasReply = CatListHas(cats, CATEGORY_NAME)
            hasAttn = CatListHas(cats, ATTENTION_CATEGORY)
            lv = VLong(AV(arr, r, C_VERB))
            handled = (lv = VERB_REPLY Or lv = VERB_REPLY_ALL Or lv = VERB_FORWARD)
            If Not handled Then handled = IRepliedInThread(VHex(AV(arr, r, C_CONV)), rcv)

            ' 내가 답한 메일은 빨강/보라 표시도 정리
            If handled Then
                If hasReply And CLEAR_REPLY_MARK_WHEN_I_REPLY Then RemoveCategoryById eid, CATEGORY_NAME
                If hasAttn And CLEAR_ATTENTION_WHEN_I_REPLY Then RemoveCategoryById eid, ATTENTION_CATEGORY
            End If

            If rcv = 0 Or rcv > newCut Then GoTo NextRow              ' 아직 기다리는 시간
            If rcv < oldCut Then                                     ' 추적 기간 지남 -> 목록에서 제거
                If hasTodo Then RemoveCategoryById eid, TODO_CATEGORY
                GoTo NextRow
            End If

            ' 목록 정보만으로 빠르게 판단
            need = Not handled
            If CatListHas(cats, DONE_CATEGORY) Then need = False
            If VLong(AV(arr, r, C_FLAG)) = FLAG_COMPLETE Then need = False
            If need And IsMeAddress(VStr(AV(arr, r, C_SENDER))) Then need = False
            toMe = AV(arr, r, C_TOME)
            If need And Not hasAttn Then
                If TODO_ONLY_TO_ME And VarType(toMe) = vbBoolean Then
                    If Not toMe Then need = False                     ' 참조(CC)/단체 주소로 받음
                End If
                If CountNames(VStr(AV(arr, r, C_TO))) + CountNames(VStr(AV(arr, r, C_CC))) > TODO_MAX_RECIPIENTS Then need = False
            End If

            ' 처음 보는 메일만 열어서 자세히 확인 (자동 발송/뉴스레터 등). 결과는 기억해 둠
            If need And Not hasTodo Then
                If m_TodoChecked.Exists(eid) Then
                    need = (Len(m_TodoChecked(eid)) = 0)
                Else
                    If sliced Then
                        If SliceOver() Then
                            m_TodoIncomplete = True                   ' 다음에 이어서
                            GoTo Done
                        End If
                    End If
                    Set itm = GetItemSafe(eid)
                    reason = TodoItemCheck(itm, hasAttn, VarType(toMe) <> vbBoolean)
                    m_TodoChecked(eid) = reason
                    need = (Len(reason) = 0)
                End If
            End If

            If need Then
                total = total + 1
                If VBool(AV(arr, r, C_UNREAD)) Then unreadCount = unreadCount + 1
                If Not hasTodo Then
                    If itm Is Nothing Then Set itm = GetItemSafe(eid)
                    AddCategoryName itm, TODO_CATEGORY
                End If
            ElseIf hasTodo Then
                RemoveCategoryById eid, TODO_CATEGORY
            End If
            If Not sliced Then
                If total Mod 20 = 0 Then DoEvents
            End If
NextRow:
        Next r
        DoEvents
    Loop
Done:
    If Err.Number <> 0 Then Debug.Print "ReplyAlert.ScanTodoFolder 오류: " & Err.Description
    ScanTodoFolder = total
End Function

' 메일을 열어야만 알 수 있는 것 확인. 필요하면 "", 아니면 이유
Private Function TodoItemCheck(ByVal itm As Object, ByVal attn As Boolean, ByVal checkToMe As Boolean) As String
    On Error GoTo Fail
    If itm Is Nothing Then TodoItemCheck = "메일을 열 수 없음": Exit Function
    If IsAutoReply(itm) Then TodoItemCheck = "자동회신": Exit Function
    If IsBulkMail(itm) Then TodoItemCheck = "자동 발송/단체 메일": Exit Function
    If checkToMe And TODO_ONLY_TO_ME And Not attn Then
        If Not IsToMe(itm) Then TodoItemCheck = "받는 사람(To)에 내가 없음 (참조/단체 주소)": Exit Function
    End If
    Exit Function
Fail:
    TodoItemCheck = "확인 중 오류"
End Function

' 진단용: 받은 메일이 [내 답장 필요] 인지. 필요하면 "", 아니면 이유
Private Function TodoReason(ByVal itm As Object) As String
    Dim lv As Long, n As Long, attn As Boolean
    On Error GoTo Fail
    EnsureMaps False
    If itm.Class <> olMail Then TodoReason = "일반 메일 아님": Exit Function
    If Left$(LCase$(itm.MessageClass), 14) = "ipm.note.rules" Then TodoReason = "자동 메일": Exit Function
    If IsDone(itm) Then TodoReason = "완료 처리됨": Exit Function
    If IsFromMe(itm) Then TodoReason = "내가 보낸 메일": Exit Function

    lv = GetPropLong(itm.PropertyAccessor, PR_LAST_VERB_EXECUTED)
    If lv = VERB_REPLY Or lv = VERB_REPLY_ALL Then TodoReason = "이미 회신함": Exit Function
    If lv = VERB_FORWARD Then TodoReason = "이미 전달함": Exit Function
    If GetPropLong(itm.PropertyAccessor, PR_FLAG_STATUS) = FLAG_COMPLETE Then TodoReason = "플래그 완료": Exit Function
    If IRepliedInThread(SafeConvId(itm), GetReceived(itm)) Then TodoReason = "같은 스레드에 내가 나중에 보낸 메일 있음": Exit Function
    If IsAutoReply(itm) Then TodoReason = "자동회신": Exit Function
    If IsBulkMail(itm) Then TodoReason = "자동 발송/단체 메일": Exit Function

    attn = CatListHas(itm.Categories, ATTENTION_CATEGORY)
    If Not attn And ATTENTION_ENABLED Then attn = (Len(AttentionMatch(itm)) > 0)
    If TODO_ONLY_TO_ME And Not attn Then
        If Not IsToMe(itm) Then TodoReason = "받는 사람(To)에 내가 없음 (참조/단체 주소)": Exit Function
    End If
    n = itm.Recipients.Count
    If n > TODO_MAX_RECIPIENTS And Not attn Then TodoReason = "받는 사람이 많은 메일 (" & n & "명)": Exit Function
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
            If IsMeAddress(RecipientSmtp(r)) Then
                IsToMe = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function RecipientSmtp(ByVal r As Object) As String
    Dim s As String, key As String, eu As Object
    On Error Resume Next
    s = GetPropStr(r.PropertyAccessor, PR_SMTP_ADDRESS)
    If Len(s) = 0 Then
        EnsureDicts
        key = LCase$(r.Address)
        If m_SmtpCache.Exists(key) Then
            s = m_SmtpCache(key)
        Else
            Set eu = r.AddressEntry.GetExchangeUser
            If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
            If Len(s) = 0 Then s = r.Address
            m_SmtpCache(key) = s
        End If
    End If
    RecipientSmtp = LCase$(Trim$(s))
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

' "홍길동; Kim, Tom; ..." 의 이름 수
Private Function CountNames(ByVal s As String) As Long
    If Len(Trim$(s)) = 0 Then Exit Function
    CountNames = UBound(Split(s, ";")) + 1
End Function


'==========================================================================
'  [답장 대기]
'==========================================================================
' 보낸 편지함을 목록(Table)으로 검사해서 [답장 대기] 범주를 붙이거나 뗌.
' 새로 붙인 개수 반환, waitingTotal = 현재 대기 총수
Private Function ScanFollowups(ByRef waitingTotal As Long) As Long
    Dim C_EID As Long, C_CLASS As Long, C_SENT As Long, C_CATS As Long
    Dim C_MSGID As Long, C_IRT As Long, C_CONV As Long
    Dim tbl As Object, arr As Variant, r As Long, itm As Object
    Dim newCut As Date, oldCut As Date, purgeCut As Date, sentOn As Date
    Dim eid As String, cats As String, hasCat As Boolean, st As Long, nNew As Long

    On Error GoTo Done
    waitingTotal = 0
    EnsureDicts
    newCut = Now - FOLLOWUP_DAYS
    oldCut = Now - FOLLOWUP_LOOKBACK_DAYS
    purgeCut = oldCut - 90
    EnsureCategory FOLLOWUP_CATEGORY, olCategoryColorOrange

    Set tbl = Application.Session.GetDefaultFolder(olFolderSentMail).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_SENT = AddCol(tbl, "SentOn")
    C_CATS = AddCol(tbl, "Categories")
    C_MSGID = AddCol(tbl, PR_INTERNET_MESSAGE_ID)
    C_IRT = AddCol(tbl, PR_IN_REPLY_TO_ID)
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    If C_EID < 0 Or C_SENT < 0 Then GoTo Done
    tbl.Sort "[SentOn]", True

    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            sentOn = VDate(AV(arr, r, C_SENT))
            If sentOn <> 0 And sentOn < purgeCut Then GoTo Done
            If Left$(RowClass(arr, r, C_CLASS), 8) <> "ipm.note" Then GoTo NextRow
            If sentOn = 0 Or sentOn >= newCut Then GoTo NextRow    ' 아직 기다리는 기간

            eid = VStr(AV(arr, r, C_EID))
            cats = VStr(AV(arr, r, C_CATS))
            hasCat = CatListHas(cats, FOLLOWUP_CATEGORY)

            If sentOn < oldCut Then                                ' 추적 기간이 지난 메일은 목록에서 제거
                If hasCat Then RemoveCategoryById eid, FOLLOWUP_CATEGORY
                GoTo NextRow
            End If

            st = FollowupStateOf(VHex(AV(arr, r, C_CONV)), sentOn, NormId(VStr(AV(arr, r, C_MSGID))), _
                                 Len(VStr(AV(arr, r, C_IRT))) > 0, CatListHas(cats, DONE_CATEGORY))
            If st = FU_WAITING Then
                waitingTotal = waitingTotal + 1
                If Not hasCat Then
                    Set itm = GetItemSafe(eid)
                    If Not itm Is Nothing Then
                        AddCategoryName itm, FOLLOWUP_CATEGORY
                        nNew = nNew + 1
                    End If
                End If
            ElseIf hasCat Then
                RemoveCategoryById eid, FOLLOWUP_CATEGORY
            End If
NextRow:
        Next r
        DoEvents
    Loop
Done:
    If Err.Number <> 0 Then Debug.Print "ReplyAlert.ScanFollowups 오류: " & Err.Description
    ScanFollowups = nNew
End Function

' 내가 보낸 메일 하나의 답장 대기 상태 (메일함 요약 정보로 판단)
Private Function FollowupStateOf(ByVal conv As String, ByVal sentOn As Date, ByVal msgId As String, _
                                 ByVal isReply As Boolean, ByVal isDone As Boolean) As Long
    Dim after As Date
    If isDone Then FollowupStateOf = FU_IGNORED: Exit Function
    If FOLLOWUP_SKIP_MY_REPLIES And isReply Then FollowupStateOf = FU_IGNORED: Exit Function
    If Len(msgId) > 0 Then
        If m_RepliedIds.Exists(msgId) Then FollowupStateOf = FU_ANSWERED: Exit Function
    End If
    If Len(conv) > 0 Then
        after = DateAdd("s", 5, sentOn)
        If m_OtherLastTime.Exists(conv) Then
            If m_OtherLastTime(conv) > after Then FollowupStateOf = FU_ANSWERED: Exit Function
        End If
        If m_MyLastTime.Exists(conv) Then
            If m_MyLastTime(conv) > after Then FollowupStateOf = FU_SUPERSEDED: Exit Function
        End If
    End If
    FollowupStateOf = FU_WAITING
End Function

' 진단용
Private Function FollowupState(ByVal sentItem As Object) As Long
    On Error Resume Next
    FollowupState = FU_ANSWERED
    If sentItem Is Nothing Then Exit Function
    EnsureMaps False
    FollowupState = FollowupStateOf(SafeConvId(sentItem), SafeSentOn(sentItem), _
                        NormId(GetPropStr(sentItem.PropertyAccessor, PR_INTERNET_MESSAGE_ID)), _
                        Len(GetPropStr(sentItem.PropertyAccessor, PR_IN_REPLY_TO_ID)) > 0, IsDone(sentItem))
End Function


'==========================================================================
'  보낸 사람 / 내 주소
'==========================================================================
Private Function IsFromMe(ByVal mail As Object) As Boolean
    IsFromMe = IsMeAddress(GetSenderSmtp(mail))
End Function

Private Function IsMeAddress(ByVal addr As String) As Boolean
    addr = LCase$(Trim$(addr))
    If Len(addr) = 0 Then Exit Function
    IsMeAddress = (InStr(1, MyAddresses(), ";" & addr & ";", vbTextCompare) > 0)
End Function

Private Function GetSenderSmtp(ByVal mail As Object) As String
    Dim s As String, key As String, ae As Object, eu As Object
    On Error Resume Next
    s = GetPropStr(mail.PropertyAccessor, PR_SENDER_SMTP_ADDRESS)   ' 빠름 (서버에 묻지 않음)
    If Len(s) = 0 Then
        If mail.SenderEmailType = "EX" Then
            EnsureDicts
            key = LCase$(mail.SenderEmailAddress)
            If m_SmtpCache.Exists(key) Then
                s = m_SmtpCache(key)
            Else
                Set ae = mail.Sender
                If Not ae Is Nothing Then
                    Set eu = ae.GetExchangeUser
                    If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
                End If
                m_SmtpCache(key) = s
            End If
        Else
            s = mail.SenderEmailAddress
        End If
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
        MsgBoxU "내 메일에 회신이 왔습니다." & vbCrLf & vbCrLf & _
               "보낸 사람: " & sender & vbCrLf & "제목: " & subj, _
               vbInformation + vbSystemModal, APP_TITLE
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
    SaveSetting SETTINGS_APP, "Favorites2", FavKey(folderName), "pending"
    TryPendingFavorites
Done:
End Sub

Private Sub TryPendingFavorites()
    Dim names As Variant, i As Long
    On Error Resume Next
    names = Array(LIST_FOLDER_NAME, ATTENTION_LIST_NAME, TODO_LIST_NAME, FOLLOWUP_LIST_NAME)
    For i = LBound(names) To UBound(names)
        If GetSetting(SETTINGS_APP, "Favorites2", FavKey(CStr(names(i))), "") = "pending" Then
            If AddToFavorites(FindSearchFolder(CStr(names(i)))) Then
                SaveSetting SETTINGS_APP, "Favorites2", FavKey(CStr(names(i))), "done"
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
        MsgBoxU "[" & folderName & "] 목록 폴더를 찾지 못했습니다.", vbExclamation, APP_TITLE
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
'  Strings / Unicode helpers
'==========================================================================
Private Sub EnsureInit()
    On Error Resume Next      ' 시작할 때 오류 창이 뜨지 않도록
    If m_Inited Then Exit Sub
    APP_TITLE = "메일 알리미"
    CATEGORY_NAME = "내 메일 회신"
    LIST_FOLDER_NAME = CATEGORY_NAME
    ATTENTION_CATEGORY = "내 앞으로 온 메일"
    ATTENTION_LIST_NAME = ATTENTION_CATEGORY
    TODO_CATEGORY = "내 답장 필요"
    TODO_LIST_NAME = TODO_CATEGORY
    FOLLOWUP_CATEGORY = "답장 대기"
    FOLLOWUP_LIST_NAME = FOLLOWUP_CATEGORY
    DONE_CATEGORY = "완료 처리"
    ATTENTION_GREETINGS = "hi;hello;hey;dear;good morning;good afternoon;good evening;morning;greetings;attn;attention;to;" & _
                          "안녕하세요;안녕하십니까;수신;받는 사람;존경하는"
    ATTENTION_NAMES = LoadNames()
    m_Inited = True
End Sub

' 내 이름 목록: %APPDATA%\ReplyAlert\names.txt (메모장으로 수정, 한 줄에 하나)
Private Function NamesFilePath() As String
    NamesFilePath = Environ$("APPDATA") & "\ReplyAlert\names.txt"
End Function

Private Function LoadNames() As String
    Dim txt As String, defaults As String, lines() As String, i As Long, ln As String, result As String
    defaults = "Jennifer" & vbCrLf & "김나현" & vbCrLf & "김 나현" & vbCrLf
    txt = ReadUtf8(NamesFilePath())
    If Len(Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))) = 0 Then
        WriteUtf8 NamesFilePath(), defaults
        txt = defaults
    End If
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(Replace(lines(i), ChrW(65279), ""))
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then result = result & ln & ";"
    Next i
    LoadNames = result & ATTENTION_EXTRA_NAMES
End Function

' 내 이름 목록 파일을 메모장으로 열기 (저장 후 Outlook 재시작하면 적용)
Public Sub EditMyNames()
    EnsureInit
    Shell "notepad.exe """ & NamesFilePath() & """", vbNormalFocus
End Sub

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    On Error GoTo Done
    If Len(Dir$(path)) = 0 Then Exit Function
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
Done:
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal txt As String)
    Dim st As Object, fso As Object, folderPath As String
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.GetParentFolderName(path)
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
Done:
End Sub

' "\uAE40\uB098" 같은 표기를 실제 글자로 (빌드 때 한글을 이렇게 바꿔 둠)
Private Function U(ByVal s As String) As String
    Dim i As Long, j As Long, v As Long, d As Long, out As String, ok As Boolean
    i = 1
    Do While i <= Len(s)
        ok = False
        If Mid$(s, i, 2) = "\u" And i + 5 <= Len(s) Then
            ok = True
            v = 0
            For j = 0 To 3
                d = InStr(1, "0123456789ABCDEF", UCase$(Mid$(s, i + 2 + j, 1)), vbBinaryCompare) - 1
                If d < 0 Then
                    ok = False
                    Exit For
                End If
                v = v * 16 + d
            Next j
        End If
        If ok Then
            out = out & ChrW(v)
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function

' 한글이 깨지지 않는 메시지 박스
Private Function MsgBoxU(ByVal prompt As String, Optional ByVal buttons As Long = 0, Optional ByVal title As String = "") As Long
    If Len(title) = 0 Then title = APP_TITLE
    MsgBoxU = MessageBoxW(0, StrPtr(prompt), StrPtr(title), buttons Or &H10000)
End Function

Private Function FavKey(ByVal folderName As String) As String
    Select Case folderName
        Case LIST_FOLDER_NAME: FavKey = "list"
        Case ATTENTION_LIST_NAME: FavKey = "attn"
        Case TODO_LIST_NAME: FavKey = "todo"
        Case FOLLOWUP_LIST_NAME: FavKey = "followup"
        Case Else: FavKey = "other"
    End Select
End Function

' 예전 버전이 (한국어가 아닌 Windows 에서) 깨진 이름으로 만든 범주/목록 폴더를 정리
Private Sub CleanupGarbled()
    Dim targets As Object, garbled As New Collection, eids As New Collection, delFolders As New Collection
    Dim cat As Object, f As Object, i As Long, eid As Variant, nm As Variant, itm As Object, n As Long
    On Error Resume Next
    Set targets = NewDict()
    targets("B3BB20B8DEC0CF20C8B8BDC5") = True          ' 내 메일 회신 (CP949)
    targets("B3BB20BED5C0B8B7CE20BFC220B8DEC0CF") = True ' 내 앞으로 온 메일
    targets("B3BB20B4E4C0E520C7CABFE4") = True          ' 내 답장 필요
    targets("B4E4C0E520B4EBB1E2") = True                ' 답장 대기
    targets("BFCFB7E120C3B3B8AE") = True                ' 완료 처리

    For Each cat In Application.Session.Categories
        If IsGarbled(cat.Name, targets) Then garbled.Add cat.Name
    Next cat

    For Each f In Application.Session.GetDefaultFolder(olFolderInbox).Store.GetSearchFolders
        If IsGarbled(f.Name, targets) Then
            For i = 1 To f.Items.Count
                eids.Add f.Items(i).EntryID
            Next i
            delFolders.Add f
        End If
    Next f

    If garbled.Count > 0 Then
        For Each eid In eids
            Set itm = Nothing
            Set itm = GetItemSafe(CStr(eid))
            If Not itm Is Nothing Then
                For Each nm In garbled
                    RemoveCategoryName itm, CStr(nm), False
                Next nm
                itm.Save
                n = n + 1
                If n Mod 20 = 0 Then DoEvents
            End If
        Next eid
    End If
    For Each f In delFolders
        f.Delete
    Next f
    For Each nm In garbled
        Application.Session.Categories.Remove CStr(nm)
    Next nm
End Sub

' 이름이 한글 이름을 잘못 읽은(깨진) 형태인지: 이 PC 의 기본 코드페이지로 되돌린 바이트가 한글(CP949) 바이트와 같으면
Private Function IsGarbled(ByVal nm As String, ByVal targets As Object) As Boolean
    Dim b() As Byte, i As Long, h As String
    On Error GoTo Done
    If Len(nm) = 0 Then Exit Function
    For i = 1 To Len(nm)
        If AscW(Mid$(nm, i, 1)) > 255 Or AscW(Mid$(nm, i, 1)) < 0 Then Exit Function   ' 진짜 한글 이름은 건드리지 않음
    Next i
    b = StrConv(nm, vbFromUnicode)
    For i = LBound(b) To UBound(b)
        h = h & Right$("0" & Hex$(b(i)), 2)
    Next i
    IsGarbled = targets.Exists(h)
Done:
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

' MarkDone 으로 완료 처리한 메일인지
Private Function IsDone(ByVal itm As Object) As Boolean
    IsDone = CatListHas(SafeCategories(itm), DONE_CATEGORY)
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

Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
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

' 대화(스레드) ID (대문자 16진수)
Private Function SafeConvId(ByVal itm As Object) As String
    On Error Resume Next
    SafeConvId = UCase$(itm.ConversationID)
End Function

' 목록(Table) 에 열 추가. 지원하지 않는 열이면 -1 (그 정보 없이 계속 진행)
Private Function AddCol(ByVal tbl As Object, ByVal colName As String) As Long
    On Error GoTo Fail
    tbl.Columns.Add colName
    AddCol = tbl.Columns.Count - 1
    Exit Function
Fail:
    AddCol = -1
End Function

Private Function AV(ByRef arr As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If c >= 0 Then AV = arr(r, c)
End Function

' 메일 종류 (열이 없으면 일반 메일로 간주)
Private Function RowClass(ByRef arr As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 0 Then
        RowClass = "ipm.note"
    Else
        RowClass = LCase$(VStr(arr(r, c)))
    End If
End Function

' 목록(Table) 값 변환
Private Function VStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsArray(v) Then
        VStr = Join(v, ", ")
    ElseIf Not IsNull(v) And Not IsEmpty(v) Then
        VStr = CStr(v)
    End If
End Function

Private Function VDate(ByVal v As Variant) As Date
    On Error Resume Next
    If IsDate(v) Then VDate = CDate(v)
End Function

Private Function VLong(ByVal v As Variant) As Long
    On Error Resume Next
    If IsNumeric(v) Then VLong = CLng(v)
End Function

Private Function VBool(ByVal v As Variant) As Boolean
    On Error Resume Next
    If VarType(v) = vbBoolean Then VBool = v
End Function

' 이진 값(바이트 배열) -> 대문자 16진수 문자열
Private Function VHex(ByVal v As Variant) As String
    Dim i As Long, s As String
    On Error Resume Next
    If VarType(v) = vbString Then
        VHex = UCase$(v)
    ElseIf VarType(v) = (vbArray + vbByte) Then
        For i = LBound(v) To UBound(v)
            s = s & Right$("0" & Hex$(v(i)), 2)
        Next i
        VHex = s
    End If
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
