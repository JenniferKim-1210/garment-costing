Attribute VB_Name = "ReplyAlert"
'==========================================================================
'  내 메일 회신 알림 + 답장 대기 목록 (Outlook VBA)
'
'  [회신 알림]
'    받은 메일이 "내가 보낸 메일에 대한 회신"이면
'      1) 빨간색 범주 [내 메일 회신] 을 붙여 목록에서 눈에 띄게 표시
'      2) Windows 알림 (클릭하면 그 메일이 열림)
'      3) 회신만 모아 보는 검색 폴더 [내 메일 회신] (즐겨찾기에 자동 추가)
'
'  [답장 대기]
'    보낸 지 N일(기본 3일)이 지났는데 아무도 답장하지 않은 메일에
'    주황색 범주 [답장 대기] 를 붙이고 검색 폴더 [답장 대기] 로 모아 보여줌
'      - 하루 한 번 자동 검사 (Outlook 시작 시 또는 그날 첫 메일 도착 시)
'      - 답장이 오면 자동으로 목록에서 빠짐
'      - 답장이 필요 없는 메일은 MarkNoReplyNeeded 로 제외
'
'  판단 방법
'    받은 메일의 In-Reply-To 헤더(어떤 메일에 대한 답장인지)를 보낸 편지함의
'    메일 ID 와 비교하고, 못 찾으면 대화(스레드) 기록으로 보조 판단
'
'  설치 방법은 README.md 참고.
'  ThisOutlookSession 에 아래 코드를 붙여넣어야 자동으로 동작합니다.
'
'    Private Sub Application_Startup()
'        ReplyAlert.OnStartup
'    End Sub
'
'    Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'        ReplyAlert.OnNewMail EntryIDCollection
'    End Sub
'
'  수동 실행 매크로 (Alt+F8)
'    ShowReplyList      : 회신 목록 열기
'    ShowNoReplyList    : 답장 대기 목록 열기
'    CheckNoReply       : 답장 대기 메일 지금 다시 검사
'    MarkNoReplyNeeded  : 선택한 보낸 메일을 답장 대기에서 제외 (답장 불필요)
'    ScanRecentInbox    : 최근 받은 메일을 검사해서 회신 표시 (기본 14일)
'    TestSelectedMail   : 선택한 메일의 판단 근거 진단
'    TestToast          : Windows 알림 테스트 (메일을 선택하고 실행하면 클릭 시 그 메일이 열림)
'    ClearReplyMark     : 선택한 메일에서 [내 메일 회신] 표시 제거
'==========================================================================
Option Explicit

'------------------------------ 설정 ------------------------------
' [회신 알림]
Public Const CATEGORY_NAME As String = "내 메일 회신"   ' 회신 표시 범주 (빨간색)
Public Const LIST_FOLDER_NAME As String = "내 메일 회신" ' 회신 목록 검색 폴더 이름
Private Const CREATE_REPLY_LIST As Boolean = True   ' 검색 폴더 자동 생성 + 즐겨찾기 추가
Private Const CLICK_TO_OPEN As Boolean = True       ' 알림 클릭 시 메일 열기
Private Const SHOW_TOAST As Boolean = True          ' Windows 알림 표시
Private Const SHOW_POPUP As Boolean = False         ' 메시지 박스 표시 (확인 누를 때까지 Outlook 멈춤)
Private Const SET_FLAG As Boolean = False           ' 플래그(오늘) 추가 -> 할 일 목록에도 표시
Private Const SET_HIGH_IMPORTANCE As Boolean = False ' 중요도 "높음" 으로 변경

' False : 내 메일에 "직접" 답장한 경우만 (권장)
' True  : 스레드 안에 내 메일이 있으면, 다른 사람 메일에 단 답장이어도 표시
Private Const MATCH_WHOLE_THREAD As Boolean = False

Private Const USE_CONVERSATION_FALLBACK As Boolean = True ' 대화(스레드) 기반 보조 판단
Private Const SKIP_AUTO_REPLIES As Boolean = True   ' 부재중 자동회신 등은 제외

Private Const SCAN_DAYS As Long = 14                ' ScanRecentInbox 검사 기간(일)
Private Const STARTUP_SCAN_DAYS As Long = 1         ' Outlook 시작 시 검사 기간(일), 0 = 안 함

' [답장 대기]
Private Const FOLLOWUP_ENABLED As Boolean = True
Public Const FOLLOWUP_CATEGORY As String = "답장 대기"   ' 답장 대기 표시 범주 (주황색)
Public Const FOLLOWUP_LIST_NAME As String = "답장 대기"  ' 답장 대기 검색 폴더 이름
Private Const FOLLOWUP_DAYS As Long = 3             ' 보낸 지 며칠 지나도 답장 없으면 표시
Private Const FOLLOWUP_LOOKBACK_DAYS As Long = 30   ' 최근 며칠 동안 보낸 메일까지 추적 (지나면 목록에서 빠짐)
Private Const FOLLOWUP_SKIP_MY_REPLIES As Boolean = False ' True = 내가 보낸 "답장(RE:)" 은 추적 안 함 (새 메일/전달만)
Private Const FOLLOWUP_DAILY_TOAST As Boolean = True ' 하루 한 번 "답장 대기 N건" 알림

' 내 주소가 여러 개(별칭)면 세미콜론으로 추가: "a@company.com;b@company.com"
Private Const EXTRA_MY_ADDRESSES As String = ""
'------------------------------------------------------------------

Private Const PR_IN_REPLY_TO_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
Private Const PR_INTERNET_REFERENCES As String = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
Private Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"

Private Const TOAST_APP_ID As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
Private Const URI_SCHEME As String = "replyalert"   ' replyalert:open/<EntryID>, replyalert:list, replyalert:followup
Private Const NOFOLLOWUP_PROP As String = "ReplyAlertNoFollowup"
Private Const SETTINGS_APP As String = "ReplyAlert"

Private Const FU_WAITING As Long = 0      ' 답장 대기
Private Const FU_ANSWERED As Long = 1     ' 답장 받음
Private Const FU_SUPERSEDED As Long = 2   ' 같은 스레드에 내가 더 나중에 보낸 메일이 있음
Private Const FU_IGNORED As Long = 3      ' 답장 불필요로 제외됨

Private m_MyAddresses As String   ' ";a@x.com;b@y.com;" 형태 캐시
Private m_HandlerReady As Boolean ' 클릭 처리기 설치 여부
Private m_FollowupRunning As Boolean


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
    RunDailyFollowupIfNeeded       ' Outlook 을 며칠씩 켜 두는 경우를 위해
End Sub

Public Sub OnStartup()
    Dim n As Long
    On Error Resume Next
    If CLICK_TO_OPEN Then InstallClickHandler
    If CREATE_REPLY_LIST Then
        EnsureSearchFolder LIST_FOLDER_NAME, olFolderInbox, CATEGORY_NAME, olCategoryColorRed
        If FOLLOWUP_ENABLED Then
            EnsureSearchFolder FOLLOWUP_LIST_NAME, olFolderSentMail, FOLLOWUP_CATEGORY, olCategoryColorOrange
        End If
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
    RunDailyFollowupIfNeeded
End Sub


'==========================================================================
'  수동 실행 매크로
'==========================================================================
Public Sub ShowReplyList()
    OpenSearchFolder LIST_FOLDER_NAME, olFolderInbox, CATEGORY_NAME, olCategoryColorRed
End Sub

Public Sub ShowNoReplyList()
    OpenSearchFolder FOLLOWUP_LIST_NAME, olFolderSentMail, FOLLOWUP_CATEGORY, olCategoryColorOrange
End Sub

Public Sub ScanRecentInbox()
    Dim n As Long
    n = ScanInbox(SCAN_DAYS)
    MsgBox "최근 " & SCAN_DAYS & "일 받은 메일 중 내 메일에 대한 회신 " & n & "건을 새로 표시했습니다." & vbCrLf & _
           "(이미 표시된 메일은 제외)", vbInformation, "회신 알림"
    If CREATE_REPLY_LIST Then ShowReplyList
End Sub

Public Sub CheckNoReply()
    Dim waiting As Long, nNew As Long
    If Not FOLLOWUP_ENABLED Then
        MsgBox "답장 대기 기능이 꺼져 있습니다. (FOLLOWUP_ENABLED)", vbExclamation, "회신 알림"
        Exit Sub
    End If
    SaveSetting SETTINGS_APP, "Followup", "LastRun", Format$(Date, "yyyy-mm-dd")
    nNew = ScanFollowups(waiting)
    MsgBox "보낸 지 " & FOLLOWUP_DAYS & "일이 지났는데 답장이 없는 메일: " & waiting & "건" & vbCrLf & _
           "(최근 " & FOLLOWUP_LOOKBACK_DAYS & "일 동안 보낸 메일 기준, 새로 추가 " & nNew & "건)", _
           vbInformation, "답장 대기"
    If CREATE_REPLY_LIST Then ShowNoReplyList
End Sub

Public Sub MarkNoReplyNeeded()
    Dim sel As Object, itm As Object, up As Object, n As Long
    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    For Each itm In sel
        If itm.Class = olMail Then
            Set up = Nothing
            Set up = itm.UserProperties.Find(NOFOLLOWUP_PROP)
            If up Is Nothing Then Set up = itm.UserProperties.Add(NOFOLLOWUP_PROP, olYesNo, False)
            up.Value = True
            RemoveCategoryName itm, FOLLOWUP_CATEGORY, False
            itm.Save
            n = n + 1
        End If
    Next itm
    If n = 0 Then MsgBox "보낸 메일을 선택한 뒤 실행하세요.", vbExclamation, "답장 대기"
End Sub

Public Sub ClearReplyMark()
    Dim sel As Object, itm As Object
    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    For Each itm In sel
        RemoveCategoryName itm, CATEGORY_NAME, True
    Next itm
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
        ShowToast "내 메일에 회신이 왔습니다", "(테스트) 클릭하면 선택한 메일이 열립니다", subj, openUri, openUri
    End If
End Sub

Public Sub TestSelectedMail()
    Dim itm As Object, pa As Object, msg As String, mine As Object
    Dim inReplyTo As String, refs As String, st As Long, stText As String

    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBox "메일을 하나 선택한 뒤 실행하세요.", vbExclamation, "회신 알림"
        Exit Sub
    End If
    If itm.Class <> olMail Then
        MsgBox "선택한 항목이 일반 메일이 아닙니다.", vbExclamation, "회신 알림"
        Exit Sub
    End If

    Set pa = itm.PropertyAccessor
    inReplyTo = GetPropStr(pa, PR_IN_REPLY_TO_ID)
    refs = GetPropStr(pa, PR_INTERNET_REFERENCES)

    msg = "제목: " & itm.Subject & vbCrLf & _
          "보낸 사람 주소: " & GetSenderSmtp(itm) & vbCrLf & _
          "내 주소 목록: " & MyAddresses() & vbCrLf & vbCrLf

    If IsFromMe(itm) Then
        st = FollowupState(itm)
        Select Case st
            Case FU_WAITING: stText = "답장 없음 (보낸 지 " & FOLLOWUP_DAYS & "일 지나면 답장 대기에 표시)"
            Case FU_ANSWERED: stText = "답장 받음"
            Case FU_SUPERSEDED: stText = "같은 스레드에 내가 나중에 보낸 메일이 있음 (그 메일 기준으로 판단)"
            Case FU_IGNORED: stText = "답장 불필요로 제외됨"
        End Select
        msg = msg & "내가 보낸 메일입니다." & vbCrLf & _
              "보낸 시각: " & itm.SentOn & vbCrLf & _
              "==> 답장 대기 판단: " & stText
    Else
        Set mine = FindRepliedMine(itm)
        msg = msg & "자동회신인가: " & YesNo(IsAutoReply(itm)) & vbCrLf & _
              "In-Reply-To: " & IIf(Len(inReplyTo) = 0, "(없음)", inReplyTo) & vbCrLf & _
              "References 개수: " & SplitIds(refs).Count & vbCrLf & vbCrLf
        If mine Is Nothing Then
            msg = msg & "==> 최종 판단: 회신 아님 X"
        Else
            msg = msg & "답장 대상(내 메일): " & mine.Subject & " / " & mine.SentOn & vbCrLf & _
                  "==> 최종 판단: 내 메일에 대한 회신 O"
        End If
    End If
    MsgBox msg, vbInformation, "회신 알림 - 진단"
End Sub


'==========================================================================
'  회신 알림 - 핵심 로직
'==========================================================================
Private Function ScanInbox(ByVal days As Long) As Long
    Dim inbox As Object, its As Object, itm As Object
    Dim cutoff As Date, rcv As Date, n As Long, total As Long

    cutoff = Now - days
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    Set its = inbox.Items
    its.Sort "[ReceivedTime]", True     ' 최신 메일부터

    For Each itm In its
        rcv = GetReceived(itm)
        If rcv <> 0 And rcv < cutoff Then Exit For
        If ProcessMail(itm, False) Then n = n + 1
        total = total + 1
        If total Mod 25 = 0 Then DoEvents
    Next itm
    ScanInbox = n
End Function

' 새로 표시했으면 True
Private Function ProcessMail(ByVal itm As Object, ByVal notify As Boolean) As Boolean
    Dim mine As Object
    On Error GoTo Fail
    If itm.Class <> olMail Then Exit Function
    If CatListHas(itm.Categories, CATEGORY_NAME) Then Exit Function
    If SKIP_AUTO_REPLIES Then
        If IsAutoReply(itm) Then Exit Function
    End If
    If IsFromMe(itm) Then Exit Function

    Set mine = FindRepliedMine(itm)
    If mine Is Nothing Then Exit Function

    MarkMail itm
    If FOLLOWUP_ENABLED Then RemoveCategoryName mine, FOLLOWUP_CATEGORY, True   ' 답장 왔으니 대기 목록에서 제거
    If notify Then NotifyReply itm
    ProcessMail = True
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
'  답장 대기 - 핵심 로직
'==========================================================================
Private Sub RunDailyFollowupIfNeeded()
    Dim today As String, waiting As Long
    On Error Resume Next
    If Not FOLLOWUP_ENABLED Then Exit Sub
    If m_FollowupRunning Then Exit Sub
    today = Format$(Date, "yyyy-mm-dd")
    If GetSetting(SETTINGS_APP, "Followup", "LastRun", "") = today Then Exit Sub
    SaveSetting SETTINGS_APP, "Followup", "LastRun", today
    ScanFollowups waiting
    If waiting > 0 And SHOW_TOAST And FOLLOWUP_DAILY_TOAST Then
        ShowToast "답장 대기 메일 " & waiting & "건", _
                  "보낸 지 " & FOLLOWUP_DAYS & "일이 지났는데 답장이 없는 메일입니다.", _
                  "클릭하면 답장 대기 목록을 엽니다.", _
                  URI_SCHEME & ":followup"
    End If
End Sub

' 보낸 편지함을 검사해서 [답장 대기] 범주를 붙이거나 뗌. 새로 붙인 개수 반환, waitingTotal = 현재 대기 총수
Private Function ScanFollowups(ByRef waitingTotal As Long) As Long
    Dim replied As Object, tbl As Object, row As Object, itm As Object
    Dim newCut As Date, oldCut As Date, purgeCut As Date, sentOn As Date
    Dim eid As String, msgId As String, hasCat As Boolean, isReply As Boolean
    Dim st As Long, nNew As Long, cnt As Long

    On Error GoTo Done
    m_FollowupRunning = True
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
    m_FollowupRunning = False
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
    If IsNoFollowup(sentItem) Then
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

Private Function IsNoFollowup(ByVal itm As Object) As Boolean
    Dim up As Object
    On Error Resume Next
    Set up = itm.UserProperties.Find(NOFOLLOWUP_PROP)
    If Not up Is Nothing Then
        IsNoFollowup = (up.Value = True)
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
    If SHOW_TOAST Then ShowToast "내 메일에 회신이 왔습니다", sender, subj, openUri, openUri
    If SHOW_POPUP Then
        MsgBox "내 메일에 회신이 왔습니다." & vbCrLf & vbCrLf & _
               "보낸 사람: " & sender & vbCrLf & "제목: " & subj, _
               vbInformation + vbSystemModal, "회신 알림"
    End If
End Sub

' Windows 10/11 알림 (PowerShell 사용, 별도 설치 불필요)
'   launchUri : 알림 본문을 클릭했을 때 열 주소
'   openUri   : 지정하면 [메일 열기] [회신 목록 보기] 버튼 표시
Private Sub ShowToast(ByVal title As String, ByVal line1 As String, ByVal line2 As String, _
                      Optional ByVal launchUri As String = "", Optional ByVal openUri As String = "")
    Dim xml As String, ps As String, head As String, actions As String
    On Error Resume Next

    If CLICK_TO_OPEN And Len(launchUri) > 0 Then
        If Not m_HandlerReady Then InstallClickHandler
    End If

    If CLICK_TO_OPEN And m_HandlerReady And Len(launchUri) > 0 Then
        head = "<toast activationType=""protocol"" launch=""" & XmlEsc(launchUri) & """ duration=""long"">"
        If Len(openUri) > 0 Then
            actions = "<actions>" & _
                      "<action content=""메일 열기"" activationType=""protocol"" arguments=""" & XmlEsc(openUri) & """/>" & _
                      "<action content=""회신 목록 보기"" activationType=""protocol"" arguments=""" & URI_SCHEME & ":list""/>" & _
                      "</actions>"
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
'  검색 폴더 (회신 목록 / 답장 대기 목록)
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
    names = Array(LIST_FOLDER_NAME, FOLLOWUP_LIST_NAME)
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
        MsgBox "[" & folderName & "] 목록 폴더를 찾지 못했습니다.", vbExclamation, "회신 알림"
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
