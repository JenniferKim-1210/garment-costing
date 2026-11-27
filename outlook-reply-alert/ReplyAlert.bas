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
'    ShowNoReplyList, CheckMyTodo, CheckNoReply, ScanRecentInbox, RecheckReplyMarks, TestSelectedMail,
'    TestToast, EditMyNames, FixBrokenCategories, ReplyReport, EditPeople, TagPeopleNow,
'    ShowUrgentList, EditKeywords, SetDeadline
'==========================================================================
Option Explicit

'============================== SETTINGS ==============================
' [Reply to my mail] (red)
' True  = every mail I receive in a conversation (thread) where I sent mail earlier ("mail I exchange")
' False = only direct replies to my mail (and mid-thread replies that put me in To)
Private Const REPLY_THREAD_MODE As Boolean = True
Private Const REPLY_REQUIRE_TO As Boolean = True        ' True = only when I am in To (not CC only)
Private Const REPLY_REQUIRE_MY_RECIPIENT As Boolean = True ' (thread mode) True = sender must be someone I sent that thread to
Private Const MATCH_WHOLE_THREAD As Boolean = False ' (direct mode only) also use the References header
Private Const CLEAR_REPLY_MARK_WHEN_I_REPLY As Boolean = False ' True = remove the mark after I reply to it
Private Const REPLY_TRACK_DAYS As Long = 60         ' alert replies to mail I sent in the last N days
Private Const SCAN_DAYS As Long = 14                ' ScanRecentInbox range (days)
Private Const STARTUP_SCAN_DAYS As Long = 1         ' check mail that arrived while Outlook was closed (days), 0 = off

' [Addressed to me] (purple) - body starts with "Hi Jennifer", "Dear Jennifer", Korean name + title ...
Private Const ATTENTION_ENABLED As Boolean = True
Private Const ATTENTION_EXTRA_NAMES As String = ""  ' extra English names, e.g. "Jenny;Ms. Kim" (Korean names: EditMyNames)
Private Const ATTENTION_LINES As Long = 8           ' how many lines from the top of the body to check
Private Const ATTENTION_AUTO_NAMES As Boolean = True ' also use my name from Outlook (e.g. "Na Hyun", "Nahyun", "Ms. Kim")
Private Const ATTENTION_TOAST As Boolean = False    ' True = Windows notification on arrival
Private Const CLEAR_ATTENTION_WHEN_I_REPLY As Boolean = False ' True = remove the mark after I reply

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

' [Urgent] (yellow) - subject / top of body contains a keyword (edit with EditKeywords)
Private Const URGENT_ENABLED As Boolean = True
Private Const URGENT_POPUP As Boolean = True        ' silent Windows notification on arrival
Private Const URGENT_FLAG As Boolean = False        ' True = also flag the mail (today)

' [Deadline] - SetDeadline: find dates in the selected mail and set a due date + reminder
Private Const DATE_ORDER_MD As Boolean = True       ' 11/12 = Nov 12 (False = 11 Dec)
Private Const DEADLINE_REMIND_DAYS_BEFORE As Long = 1 ' reminder N days before the due date
Private Const DEADLINE_REMIND_HOUR As Long = 9      ' reminder time (hour)
Private Const DEADLINE_AUTO As Boolean = False      ' True = new mail with "by/until/deadline + date" gets a due date automatically

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
Private Const SLICE_SECONDS As Single = 0.4         ' work at most N seconds at a time, continue on the next mail
Private Const BACKGROUND_INTERVAL_SECONDS As Long = 180 ' update the lists at most once every N seconds (reminders still on time)
Private Const SUBFOLDER_SWEEP_MINUTES As Long = 30  ' check Inbox subfolders (mail moved by rules) every N minutes, 0 = off

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
Private URGENT_CATEGORY As String, URGENT_LIST_NAME As String
Private m_UrgentKw As String       ' "|kw|kw|" ()
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
Private Const PR_CONVERSATION_TOPIC As String = "http://schemas.microsoft.com/mapi/proptag/0x0070001F"

Private Const VERB_REPLY As Long = 102
Private Const VERB_REPLY_ALL As Long = 103
Private Const VERB_FORWARD As Long = 104
Private Const FLAG_COMPLETE As Long = 1

Private Const TOAST_APP_ID As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
Private Const URI_SCHEME As String = "replyalert"   ' replyalert:open/<EntryID>, :list, :attn, :todo, :followup
Private Const SETTINGS_APP As String = "ReplyAlert"
Private Const TABLE_CHUNK As Long = 200

Private Const FU_WAITING As Long = 0
Private Const FU_ANSWERED As Long = 1
Private Const FU_SUPERSEDED As Long = 2
Private Const FU_IGNORED As Long = 3

' (Outlook )
Private m_MyAddresses As String   ' ";a@x.com;b@y.com;"
Private m_HandlerReady As Boolean
Private m_Busy As Boolean
Private m_StartTime As Date
Private m_SliceStart As Single

Private m_MapsTime As Date
Private m_SentIdsRefreshAt As Date
Private m_SentIds As Object        ' ID() -> EntryID
Private m_MyLastTime As Object     ' ID ->
Private m_MyLastId As Object       ' ID -> EntryID
Private m_MyFirstTime As Object    ' ID ->
Private m_MyThreadFirst As Object  ' "c:ID" / "t:" ->
Private m_MyThreadIds As Object    ' "c:ID" / "t:" -> EntryID ("|" , 10)
Private m_ThreadRecips As Object   ' -> ("|a|b|")
Private m_OtherLastTime As Object  ' ID ->
Private m_RepliedIds As Object     ' ID() ( In-Reply-To)
Private m_SmtpCache As Object      ' Exchange -> SMTP
Private m_TodoChecked As Object    ' EntryID -> "" ( )
Private m_Seen As Object           ' EntryID

Private m_PendingNew As Collection ' EntryID
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
Private m_LastBackground As Date
Private m_FuChecked As Object      ' EntryID -> (FU_*)
Private m_FuCheckedTime As Object  ' EntryID ->
Private m_FuIncomplete As Boolean
Private m_FuDailyPending As Boolean
Private m_SubFolders As Collection
Private m_SubSweepIdx As Long
Private m_LastSubSweep As Date
Private m_SweepReplies As Long
Private m_PeopleLabel() As String
Private m_PeopleKw() As String
Private m_PeopleCount As Long
Private m_PeopleSet() As Object    ' : ("c:ID", "t:")
Private m_PeopleSetsReady As Boolean


'==========================================================================
' (ThisOutlookSession )
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
        EnsurePeopleFolders
        If URGENT_ENABLED Then EnsureSearchFolder URGENT_LIST_NAME, olFolderInbox, URGENT_CATEGORY, olCategoryColorYellow
    End If
    ' ( STARTUP_DELAY_SECONDS , )
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
            m_SentIdsRefreshAt = Now + TimeSerial(0, 5, 0)  ' ID 5
        End If
        NoteMine SafeConvId(sentItem), t, eid
        NoteMyThread SafeConvId(sentItem), NormTopic(sentItem.ConversationTopic), t, eid
    End If
    TagPeople sentItem
    m_TodoDirty = True
    m_FuDirty = True
    RunScheduledChecks
End Sub


'==========================================================================
'==========================================================================
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
            RemoveCategoryName itm, URGENT_CATEGORY, False
            If Not CatListHas(itm.Categories, DONE_CATEGORY) Then
                If Len(itm.Categories) = 0 Then
                    itm.Categories = DONE_CATEGORY
                Else
                    itm.Categories = itm.Categories & ", " & DONE_CATEGORY
                End If
            End If
            itm.Save
            m_TodoChecked(SafeEntryId(itm)) = U("\uC644\uB8CC \uCC98\uB9AC\uB428")
            n = n + 1
        End If
    Next itm
    If n = 0 Then MsgBoxU U("\uC644\uB8CC \uCC98\uB9AC\uD560 \uBA54\uC77C\uC744 \uC120\uD0DD\uD55C \uB4A4 \uC2E4\uD589\uD558\uC138\uC694."), vbExclamation, APP_TITLE
End Sub

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

Public Sub ShowUrgentList()
    EnsureInit
    OpenSearchFolder URGENT_LIST_NAME, olFolderInbox, URGENT_CATEGORY, olCategoryColorYellow
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
        MsgBoxU U("[\uB0B4 \uB2F5\uC7A5 \uD544\uC694] \uAE30\uB2A5\uC774 \uAEBC\uC838 \uC788\uC2B5\uB2C8\uB2E4. (TODO_ENABLED)"), vbExclamation, APP_TITLE
        Exit Sub
    End If
    If m_Busy Then
        MsgBoxU U("\uC9C0\uAE08 \uB2E4\uB978 \uAC80\uC0AC\uB97C \uD558\uB294 \uC911\uC785\uB2C8\uB2E4. \uC7A0\uC2DC \uD6C4 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694."), vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    EnsureMaps True
    total = ScanTodo(unread, False)
    m_Busy = False
    MsgBoxU U("\uBC1B\uC740 \uC9C0 ") & TODO_MIN_HOURS & U("\uC2DC\uAC04\uC774 \uC9C0\uB0AC\uB294\uB370 \uC544\uC9C1 \uD68C\uC2E0\uD558\uC9C0 \uC54A\uC740 \uBA54\uC77C: ") & total & U("\uAC74") & vbCrLf & _
           U("(\uADF8\uC911 \uC548 \uC77D\uC740 \uBA54\uC77C ") & unread & U("\uAC74, \uCD5C\uADFC ") & TODO_LOOKBACK_DAYS & U("\uC77C \uAE30\uC900)"), _
           vbInformation, U("\uB0B4 \uB2F5\uC7A5 \uD544\uC694")
    If CREATE_LISTS Then ShowMyTodoList
End Sub

Public Sub CheckNoReply()
    Dim waiting As Long, nNew As Long
    EnsureInit
    If Not FOLLOWUP_ENABLED Then
        MsgBoxU U("[\uB2F5\uC7A5 \uB300\uAE30] \uAE30\uB2A5\uC774 \uAEBC\uC838 \uC788\uC2B5\uB2C8\uB2E4. (FOLLOWUP_ENABLED)"), vbExclamation, APP_TITLE
        Exit Sub
    End If
    If m_Busy Then
        MsgBoxU U("\uC9C0\uAE08 \uB2E4\uB978 \uAC80\uC0AC\uB97C \uD558\uB294 \uC911\uC785\uB2C8\uB2E4. \uC7A0\uC2DC \uD6C4 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694."), vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    EnsureMaps True
    SaveSetting SETTINGS_APP, "Followup", "LastRun", Format$(Date, "yyyy-mm-dd")
    nNew = ScanFollowups(waiting, False)
    m_FuDirty = False
    m_Busy = False
    MsgBoxU U("\uBCF4\uB0B8 \uC9C0 ") & FOLLOWUP_DAYS & U("\uC77C\uC774 \uC9C0\uB0AC\uB294\uB370 \uB2F5\uC7A5\uC774 \uC5C6\uB294 \uBA54\uC77C: ") & waiting & U("\uAC74") & vbCrLf & _
           U("(\uCD5C\uADFC ") & FOLLOWUP_LOOKBACK_DAYS & U("\uC77C \uB3D9\uC548 \uBCF4\uB0B8 \uBA54\uC77C \uAE30\uC900, \uC0C8\uB85C \uCD94\uAC00 ") & nNew & U("\uAC74)"), _
           vbInformation, U("\uB2F5\uC7A5 \uB300\uAE30")
    If CREATE_LISTS Then ShowNoReplyList
End Sub

Public Sub ScanRecentInbox()
    Dim n As Long, attn As Long, v As String, days As Long
    EnsureInit
    If m_Busy Then
        MsgBoxU U("\uC9C0\uAE08 \uB2E4\uB978 \uAC80\uC0AC\uB97C \uD558\uB294 \uC911\uC785\uB2C8\uB2E4. \uC7A0\uC2DC \uD6C4 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694."), vbInformation, APP_TITLE
        Exit Sub
    End If
    ' (InputBox )
    v = InputBox("How many days back should I check the Inbox?" & vbCrLf & "(e.g. 14, 30, 60)", "ReplyAlert", CStr(SCAN_DAYS))
    If Len(Trim$(v)) = 0 Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    days = CLng(v)
    If days < 1 Then Exit Sub
    m_Busy = True
    EnsureMaps True
    Dim folders As New Collection, fld As Object
    folders.Add Application.Session.GetDefaultFolder(olFolderInbox)
    AddSubFolders Application.Session.GetDefaultFolder(olFolderInbox), folders, 0
    For Each fld In folders
        ScanInboxTable days, False, n, attn, fld
    Next fld
    m_Busy = False
    MsgBoxU U("\uCD5C\uADFC ") & days & U("\uC77C \uBC1B\uC740 \uBA54\uC77C\uC5D0\uC11C \uC0C8\uB85C \uD45C\uC2DC\uD588\uC2B5\uB2C8\uB2E4.") & vbCrLf & vbCrLf & _
           U("  \uB0B4 \uBA54\uC77C\uC5D0 \uB300\uD55C \uD68C\uC2E0: ") & n & U("\uAC74") & vbCrLf & _
           U("  \uB0B4 \uC55E\uC73C\uB85C \uC628 \uBA54\uC77C: ") & attn & U("\uAC74") & vbCrLf & vbCrLf & _
           U("(\uC774\uBBF8 \uD45C\uC2DC\uB41C \uBA54\uC77C\uC740 \uC81C\uC678)"), vbInformation, APP_TITLE
    If CREATE_LISTS Then ShowAttentionList
End Sub

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
        ShowToast U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0\uC774 \uC654\uC2B5\uB2C8\uB2E4"), U("\uD64D\uAE38\uB3D9 (\uD14C\uC2A4\uD2B8)"), U("RE: \uC54C\uB9BC \uD14C\uC2A4\uD2B8 \uBA54\uC77C"), URI_SCHEME & ":list"
    Else
        ShowToast U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0\uC774 \uC654\uC2B5\uB2C8\uB2E4"), U("(\uD14C\uC2A4\uD2B8) \uD074\uB9AD\uD558\uBA74 \uC120\uD0DD\uD55C \uBA54\uC77C\uC774 \uC5F4\uB9BD\uB2C8\uB2E4"), subj, openUri, _
                  U("\uBA54\uC77C \uC5F4\uAE30|") & openUri & U(";\uD68C\uC2E0 \uBAA9\uB85D \uBCF4\uAE30|") & URI_SCHEME & ":list"
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
        MsgBoxU U("\uBA54\uC77C\uC744 \uD558\uB098 \uC120\uD0DD\uD55C \uB4A4 \uC2E4\uD589\uD558\uC138\uC694."), vbExclamation, APP_TITLE
        Exit Sub
    End If
    If itm.Class <> olMail Then
        MsgBoxU U("\uC120\uD0DD\uD55C \uD56D\uBAA9\uC774 \uC77C\uBC18 \uBA54\uC77C\uC774 \uC544\uB2D9\uB2C8\uB2E4."), vbExclamation, APP_TITLE
        Exit Sub
    End If
    EnsureMaps False

    msg = U("\uC81C\uBAA9: ") & itm.Subject & vbCrLf & _
          U("\uBCF4\uB0B8 \uC0AC\uB78C \uC8FC\uC18C: ") & GetSenderSmtp(itm) & vbCrLf & _
          U("\uB0B4 \uC8FC\uC18C \uBAA9\uB85D: ") & MyAddresses() & vbCrLf & _
          U("\uC644\uB8CC \uCC98\uB9AC\uB428: ") & YesNo(IsDone(itm)) & vbCrLf & _
          U("\uB300\uD654 ID \uC788\uC74C: ") & YesNo(Len(SafeConvId(itm)) > 0) & vbCrLf & vbCrLf

    If IsFromMe(itm) Then
        st = FollowupState(itm)
        If st = FU_WAITING Then st = ConversationFollowupState(itm, stText)
        If st = FU_ANSWERED And Len(stText) > 0 Then stText = U(" (\uD68C\uC2E0 \uC704\uCE58: ") & stText & ")" Else stText = ""
        reason = stText
        Select Case st
            Case FU_WAITING: stText = U("\uB2F5\uC7A5 \uC5C6\uC74C - \uBAA8\uB4E0 \uD3F4\uB354\uC758 \uAC19\uC740 \uC2A4\uB808\uB4DC\uB97C \uD655\uC778\uD574\uB3C4 \uB2E4\uB978 \uC0AC\uB78C \uBA54\uC77C\uC774 \uC5C6\uC74C (\uBCF4\uB0B8 \uC9C0 ") & FOLLOWUP_DAYS & U("\uC77C \uC9C0\uB098\uBA74 [\uB2F5\uC7A5 \uB300\uAE30] \uC5D0 \uD45C\uC2DC)")
            Case FU_ANSWERED: stText = U("\uB2F5\uC7A5 \uBC1B\uC74C") & reason
            Case FU_SUPERSEDED: stText = U("\uAC19\uC740 \uC2A4\uB808\uB4DC\uC5D0 \uB0B4\uAC00 \uB098\uC911\uC5D0 \uBCF4\uB0B8 \uBA54\uC77C\uC774 \uC788\uC74C (\uADF8 \uBA54\uC77C \uAE30\uC900\uC73C\uB85C \uD310\uB2E8)")
            Case FU_IGNORED: stText = U("\uC644\uB8CC \uCC98\uB9AC\uB428 / \uCD94\uC801 \uC548 \uD568")
        End Select
        msg = msg & U("[\uB0B4\uAC00 \uBCF4\uB0B8 \uBA54\uC77C]") & vbCrLf & _
              U("\uBCF4\uB0B8 \uC2DC\uAC01: ") & itm.SentOn & vbCrLf & _
              U("==> \uB2F5\uC7A5 \uB300\uAE30 \uD310\uB2E8: ") & stText
    Else
        inReplyTo = GetPropStr(itm.PropertyAccessor, PR_IN_REPLY_TO_ID)
        Set mine = FindRepliedMine(itm)
        msg = msg & U("[\uBC1B\uC740 \uBA54\uC77C]") & vbCrLf & _
              "In-Reply-To: " & IIf(Len(inReplyTo) = 0, U("(\uC5C6\uC74C)"), inReplyTo) & vbCrLf
        If mine Is Nothing Then
            msg = msg & U("==> \uB0B4 \uBA54\uC77C\uC5D0 \uB300\uD55C \uD68C\uC2E0: \uC544\uB2D8") & vbCrLf
        Else
            msg = msg & U("==> \uB0B4 \uBA54\uC77C\uC5D0 \uB300\uD55C \uD68C\uC2E0: \uB9DE\uC74C (") & mine.Subject & " / " & mine.SentOn & ")" & vbCrLf
        End If
        stText = AttentionMatch(itm)
        If Len(stText) = 0 Then
            msg = msg & U("==> \uB0B4 \uC55E\uC73C\uB85C \uC628 \uBA54\uC77C: \uC544\uB2D8 (\uBCF8\uBB38 \uC55E ") & ATTENTION_LINES & U("\uC904\uC5D0 \uC544\uB798 \uC774\uB984\uC744 \uBD80\uB974\uB294 \uC778\uC0AC\uAC00 \uC5C6\uC74C)") & vbCrLf & _
                  U("    \uCC3E\uB294 \uC774\uB984: ") & Replace(ATTENTION_NAMES, ";", ", ") & vbCrLf & _
                  U("    \uBCF8\uBB38 \uCCAB \uC904\uB4E4: ") & FirstLines(itm, 3) & vbCrLf
        Else
            msg = msg & U("==> \uB0B4 \uC55E\uC73C\uB85C \uC628 \uBA54\uC77C: \uB9DE\uC74C [") & Left$(stText, 60) & "]" & vbCrLf
            If Not CatListHas(itm.Categories, ATTENTION_CATEGORY) Then
                If IsBulkMail(itm) Then
                    msg = msg & U("    (\uD45C\uC2DC \uC548 \uB428: \uC790\uB3D9 \uBC1C\uC1A1/\uB2E8\uCCB4 \uBA54\uC77C\uB85C \uD310\uB2E8)") & vbCrLf
                Else
                    msg = msg & U("    (\uC544\uC9C1 \uD45C\uC2DC \uC804: ScanRecentInbox \uB97C \uC2E4\uD589\uD558\uBA74 \uD45C\uC2DC\uB428)") & vbCrLf
                End If
            End If
        End If
        reason = TodoReason(itm)
        If Len(reason) = 0 Then
            msg = msg & U("==> \uB0B4 \uB2F5\uC7A5 \uD544\uC694: \uC608 (\uBC1B\uC740 \uC9C0 ") & TODO_MIN_HOURS & U("\uC2DC\uAC04 \uC9C0\uB098\uBA74 \uBAA9\uB85D\uC5D0 \uD45C\uC2DC)")
        Else
            msg = msg & U("==> \uB0B4 \uB2F5\uC7A5 \uD544\uC694: \uC544\uB2C8\uC624 - ") & reason
        End If
    End If
    MsgBoxU msg, vbInformation, APP_TITLE & " - " & U("\uC9C4\uB2E8")
End Sub


'==========================================================================
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

    ' 0) / ( )
    If GetSetting(SETTINGS_APP, "Cleanup", "Garbled", "") = "" Then
        SaveSetting SETTINGS_APP, "Cleanup", "Garbled", "done"
        CleanupGarbled
        GoTo Done
    End If
    ' 0-1) (// , )
    If GetSetting(SETTINGS_APP, "Cleanup", "Garbled2", "") = "" Then
        SaveSetting SETTINGS_APP, "Cleanup", "Garbled2", "done"
        CleanupGarbledItems
        GoTo Done
    End If
    ' 0-2) [ ] ( )
    If GetSetting(SETTINGS_APP, "Cleanup", "ReplyRecheck2", "") = "" Then
        SaveSetting SETTINGS_APP, "Cleanup", "ReplyRecheck2", "done"
        RecheckReplyMarksCore
        GoTo Done
    End If

    ' 1) ( )
    If Not ProcessPendingNew() Then GoTo Done

    ' ( Outlook ).
    slotKey = CurrentRemindSlot()
    If Len(slotKey) > 0 Then doRemind = (GetSetting(SETTINGS_APP, "Todo", "LastRemind", "") <> slotKey)
    If Not doRemind And m_LastBackground <> 0 Then
        If DateDiff("s", m_LastBackground, Now) < BACKGROUND_INTERVAL_SECONDS Then GoTo Done
    End If
    m_LastBackground = Now

    ' 2) Outlook ( )
    If m_CatchUpPending Then
        If Not CatchUpStep() Then GoTo Done
    End If
    If SliceOver() Then GoTo Done

    ' 2-2) (SUBFOLDER_SWEEP_MINUTES , )
    If Not SubfolderSweepStep() Then GoTo Done
    If SliceOver() Then GoTo Done

    today = Format$(Date, "yyyy-mm-dd")

    ' 3) [ ] + ( )
    If FOLLOWUP_ENABLED Then
        dailyDue = (GetSetting(SETTINGS_APP, "Followup", "LastRun", "") <> today)
        If dailyDue Then
            SaveSetting SETTINGS_APP, "Followup", "LastRun", today
            m_FuDailyPending = True
        End If
        If dailyDue Or m_FuDirty Or m_FuIncomplete Then
            ScanFollowups fuN, True
            m_FuDirty = False
            If m_FuDailyPending And Not m_FuIncomplete Then
                m_FuDailyPending = False
                m_FuToastCount = fuN
                m_FuToastPending = (fuN > 0)
            End If
        End If
    End If
    If SliceOver() Then GoTo Done

    ' 4) [ ] / /
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

    ' 5)
    If SHOW_TOAST Then
        If remindNow And TODO_ENABLED And todoN > 0 Then
            line1 = U("\uB0B4 \uB2F5\uC7A5 \uD544\uC694 ") & todoN & U("\uAC74") & IIf(unreadN > 0, U(" (\uC548 \uC77D\uC74C ") & unreadN & U("\uAC74)"), "")
            launch = URI_SCHEME & ":todo"
        End If
        If m_FuToastPending And (remindNow Or Not doRemind) Then
            line2 = U("\uB2F5\uC7A5 \uB300\uAE30 ") & m_FuToastCount & U("\uAC74 (\uBCF4\uB0B8 \uC9C0 ") & FOLLOWUP_DAYS & U("\uC77C \uB118\uAC8C \uB2F5 \uC5C6\uC74C)")
            If Len(launch) = 0 Then launch = URI_SCHEME & ":followup"
            m_FuToastPending = False
        End If
        If Len(line1) = 0 Then
            line1 = line2
            line2 = ""
        End If
        If Len(line1) > 0 Then
            ShowToast U("\uBA54\uC77C \uB9AC\uB9C8\uC778\uB4DC"), line1, line2, launch, _
                      U("\uB0B4 \uB2F5\uC7A5 \uD544\uC694|") & URI_SCHEME & U(":todo;\uB2F5\uC7A5 \uB300\uAE30|") & URI_SCHEME & ":followup"
        End If
    End If
Done:
    If Err.Number <> 0 Then Debug.Print U("ReplyAlert.RunScheduledChecks \uC624\uB958: ") & Err.Description
    m_Busy = False
End Sub

' . True
Private Function ProcessPendingNew() As Boolean
    Dim eid As String, itm As Object, attn As Boolean, marked As Boolean
    On Error GoTo Done
    If m_PendingNew Is Nothing Then ProcessPendingNew = True: Exit Function
    If m_PendingNew.Count > 3 Then m_PendingSummary = True
    Do While m_PendingNew.Count > 0
        eid = m_PendingNew(1)
        m_PendingNew.Remove 1
        m_Seen(eid) = True
        Set itm = GetItemSafe(eid)
        If Not itm Is Nothing Then
            UpdateMapsWithItem itm
            TagPeople itm
            marked = ProcessMail(itm, Not m_PendingSummary, attn)
            If marked And m_PendingSummary Then m_SumReplies = m_SumReplies + 1
        End If
        If SliceOver() Then Exit Function
    Loop
    If m_PendingSummary Then
        m_PendingSummary = False
        If m_SumReplies > 0 And SHOW_TOAST Then
            ShowToast U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0 ") & m_SumReplies & U("\uAC74"), _
                      U("\uD55C\uAEBC\uBC88\uC5D0 \uB3C4\uCC29\uD55C \uBA54\uC77C \uC911 \uB0B4 \uBA54\uC77C\uC5D0 \uB300\uD55C \uD68C\uC2E0\uC785\uB2C8\uB2E4."), _
                      U("\uD074\uB9AD\uD558\uBA74 \uD68C\uC2E0 \uBAA9\uB85D\uC744 \uC5FD\uB2C8\uB2E4."), URI_SCHEME & ":list"
        End If
        m_SumReplies = 0
    End If
    ProcessPendingNew = True
Done:
End Function

' 1 ( ). True
Private Function SubfolderSweepStep() As Boolean
    Dim reps As Long, attns As Long
    On Error GoTo Fail
    If SUBFOLDER_SWEEP_MINUTES <= 0 Then SubfolderSweepStep = True: Exit Function
    If m_SubSweepIdx = 0 Then
        If m_LastSubSweep <> 0 Then
            If DateDiff("n", m_LastSubSweep, Now) < SUBFOLDER_SWEEP_MINUTES Then SubfolderSweepStep = True: Exit Function
        End If
        Set m_SubFolders = New Collection
        AddSubFolders Application.Session.GetDefaultFolder(olFolderInbox), m_SubFolders, 0
        m_SubSweepIdx = 1
        m_SweepReplies = 0
    End If
    Do While m_SubSweepIdx <= m_SubFolders.Count
        reps = 0
        If Not ScanInboxTable(1, True, reps, attns, m_SubFolders(m_SubSweepIdx)) Then
            m_SweepReplies = m_SweepReplies + reps
            Exit Function
        End If
        m_SweepReplies = m_SweepReplies + reps
        m_SubSweepIdx = m_SubSweepIdx + 1
        If SliceOver() Then Exit Function
    Loop
    m_SubSweepIdx = 0
    m_LastSubSweep = Now
    Set m_SubFolders = Nothing
    If m_SweepReplies > 0 And SHOW_TOAST Then
        ShowToast U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0 ") & m_SweepReplies & U("\uAC74"), _
                  U("\uBC1B\uC740 \uD3B8\uC9C0\uD568\uC758 \uD558\uC704 \uD3F4\uB354\uC5D0 \uB4E4\uC5B4\uC628 \uD68C\uC2E0\uC785\uB2C8\uB2E4."), _
                  U("\uD074\uB9AD\uD558\uBA74 \uD68C\uC2E0 \uBAA9\uB85D\uC744 \uC5FD\uB2C8\uB2E4."), URI_SCHEME & ":list"
    End If
    m_SweepReplies = 0
    SubfolderSweepStep = True
    Exit Function
Fail:
    m_SubSweepIdx = 0
    m_LastSubSweep = Now
    SubfolderSweepStep = True
End Function

' Outlook . True
Private Function CatchUpStep() As Boolean
    Dim reps As Long, attns As Long
    If ScanInboxTable(STARTUP_SCAN_DAYS, True, reps, attns) Then
        m_CatchUpReplies = m_CatchUpReplies + reps
        m_CatchUpPending = False
        If m_CatchUpReplies > 0 And SHOW_TOAST Then
            ShowToast U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0 ") & m_CatchUpReplies & U("\uAC74"), _
                      U("Outlook \uC774 \uAEBC\uC838 \uC788\uB294 \uB3D9\uC548 \uB3C4\uCC29\uD55C \uD68C\uC2E0\uC785\uB2C8\uB2E4."), _
                      U("\uD074\uB9AD\uD558\uBA74 \uD68C\uC2E0 \uBAA9\uB85D\uC744 \uC5FD\uB2C8\uB2E4."), URI_SCHEME & ":list"
        End If
        m_CatchUpReplies = 0
        CatchUpStep = True
    Else
        m_CatchUpReplies = m_CatchUpReplies + reps
    End If
End Function

' ("2026-11-26 14"), ""
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

Private Function SliceOver() As Boolean
    Dim t As Single
    t = Timer
    If t < m_SliceStart Then t = t + 86400!
    SliceOver = ((t - m_SliceStart) > SLICE_SECONDS)
End Function


'==========================================================================
' ( , ID) - (Table)
'==========================================================================
Private Sub EnsureDicts()
    If m_SentIds Is Nothing Then Set m_SentIds = NewDict()
    If m_MyLastTime Is Nothing Then Set m_MyLastTime = NewDict()
    If m_MyLastId Is Nothing Then Set m_MyLastId = NewDict()
    If m_MyFirstTime Is Nothing Then Set m_MyFirstTime = NewDict()
    If m_MyThreadFirst Is Nothing Then Set m_MyThreadFirst = NewDict()
    If m_MyThreadIds Is Nothing Then Set m_MyThreadIds = NewDict()
    If m_ThreadRecips Is Nothing Then Set m_ThreadRecips = NewDict()
    If m_OtherLastTime Is Nothing Then Set m_OtherLastTime = NewDict()
    If m_RepliedIds Is Nothing Then Set m_RepliedIds = NewDict()
    If m_SmtpCache Is Nothing Then Set m_SmtpCache = NewDict()
    If m_TodoChecked Is Nothing Then Set m_TodoChecked = NewDict()
    If m_Seen Is Nothing Then Set m_Seen = NewDict()
    If m_FuChecked Is Nothing Then Set m_FuChecked = NewDict()
    If m_FuCheckedTime Is Nothing Then Set m_FuCheckedTime = NewDict()
End Sub

Private Sub EnsureMaps(ByVal force As Boolean)
    EnsureDicts
    If force Or m_MapsTime = 0 Or DateDiff("h", m_MapsTime, Now) >= 24 Then
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
    Set m_MyFirstTime = NewDict()
    Set m_MyThreadFirst = NewDict()
    Set m_MyThreadIds = NewDict()
    Set m_ThreadRecips = NewDict()
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
    Dim C_EID As Long, C_CLASS As Long, C_TIME As Long, C_MSGID As Long, C_CONV As Long, C_TOPIC As Long, C_SUBJ As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, msgId As String, eid As String, topic As String
    On Error GoTo Done
    Set tbl = Application.Session.GetDefaultFolder(olFolderSentMail).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_TIME = AddCol(tbl, "SentOn")
    C_MSGID = AddCol(tbl, PR_INTERNET_MESSAGE_ID)
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    C_TOPIC = AddCol(tbl, PR_CONVERSATION_TOPIC)
    C_SUBJ = AddCol(tbl, "Subject")
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
                topic = VStr(AV(arr, r, C_TOPIC))
                If Len(topic) = 0 Then topic = VStr(AV(arr, r, C_SUBJ))
                NoteMyThread VHex(AV(arr, r, C_CONV)), NormTopic(topic), t, eid
            End If
        Next r
        DoEvents
    Loop
Done:
End Sub

' ( ID / )
Private Sub NoteMyThread(ByVal conv As String, ByVal topic As String, ByVal t As Date, ByVal eid As String)
    If t = 0 Then Exit Sub
    If Len(conv) > 0 Then NoteMyThreadKey "c:" & conv, t, eid
    If Len(topic) >= 4 Then NoteMyThreadKey "t:" & topic, t, eid
End Sub

Private Sub NoteMyThreadKey(ByVal key As String, ByVal t As Date, ByVal eid As String)
    Dim ids As String
    If Not m_MyThreadFirst.Exists(key) Then
        m_MyThreadFirst(key) = t
    ElseIf m_MyThreadFirst(key) > t Then
        m_MyThreadFirst(key) = t
    End If
    If m_MyThreadIds.Exists(key) Then ids = m_MyThreadIds(key)
    If InStr(1, ids, "|" & eid & "|", vbBinaryCompare) = 0 And UBound(Split(ids, "|")) < 20 Then
        If Len(ids) = 0 Then ids = "|"
        m_MyThreadIds(key) = ids & eid & "|"
        If m_ThreadRecips.Exists(key) Then m_ThreadRecips.Remove key
    End If
End Sub

' (To/CC) addr
Private Function ThreadRecipHas(ByVal key As String, ByVal addr As String) As Boolean
    Dim ids() As String, i As Long, itm As Object, r As Object, lst As String
    On Error Resume Next
    addr = LCase$(Trim$(addr))
    If Len(addr) = 0 Then Exit Function
    If Not m_ThreadRecips.Exists(key) Then
        lst = "|"
        ids = Split(m_MyThreadIds(key), "|")
        For i = LBound(ids) To UBound(ids)
            If Len(ids(i)) > 0 Then
                Set itm = Nothing
                Set itm = GetItemSafe(ids(i))
                If Not itm Is Nothing Then
                    For Each r In itm.Recipients
                        lst = lst & RecipientSmtp(r) & "|"
                    Next r
                End If
            End If
        Next i
        m_ThreadRecips(key) = lst
    End If
    ThreadRecipHas = (InStr(1, m_ThreadRecips(key), "|" & addr & "|", vbTextCompare) > 0)
End Function

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
            If m_MyLastTime.Exists(conv) Then m_FuDirty = True
        End If
    End If
End Sub

Private Sub NoteMine(ByVal conv As String, ByVal t As Date, ByVal eid As String)
    If Len(conv) = 0 Or t = 0 Then Exit Sub
    If Not m_MyFirstTime.Exists(conv) Then
        m_MyFirstTime(conv) = t
    ElseIf m_MyFirstTime(conv) > t Then
        m_MyFirstTime(conv) = t
    End If
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
'==========================================================================
' [ ] True. attnMarked = [ ]
Private Function ProcessMail(ByVal itm As Object, ByVal notify As Boolean, Optional ByRef attnMarked As Boolean) As Boolean
    Dim mine As Object, lv As Long, handled As Boolean, urgentKw As String, urgentHit As Boolean, cands As Collection
    On Error GoTo Fail
    attnMarked = False
    If itm.Class <> olMail Then Exit Function
    If IsDone(itm) Then Exit Function
    If IsFromMe(itm) Then Exit Function
    If SKIP_AUTO_REPLIES Then
        If IsAutoReply(itm) Then Exit Function
    End If

    If URGENT_ENABLED Then
        If Not CatListHas(itm.Categories, URGENT_CATEGORY) Then
            urgentKw = UrgentMatch(itm)
            If Len(urgentKw) > 0 Then
                If Not IsBulkMail(itm) Then
                    EnsureCategory URGENT_CATEGORY, olCategoryColorYellow
                    AddCategoryName itm, URGENT_CATEGORY
                    If URGENT_FLAG Then
                        itm.MarkAsTask olMarkToday
                        itm.Save
                    End If
                    urgentHit = True
                End If
            End If
        End If
    End If

    ' [ ] "by / until / deadline / + " ( )
    If DEADLINE_AUTO And notify Then
        If GetPropLong(itm.PropertyAccessor, PR_FLAG_STATUS) = 0 Then
            Set cands = FindDeadlines(itm, True)
            If cands.Count > 0 Then RegisterDeadline itm, CDate(cands(1)(0))
        End If
    End If

    lv = GetPropLong(itm.PropertyAccessor, PR_LAST_VERB_EXECUTED)
    handled = (lv = VERB_REPLY Or lv = VERB_REPLY_ALL Or lv = VERB_FORWARD)
    If Not handled Then handled = IRepliedInThread(SafeConvId(itm), GetReceived(itm))

    ' [ ] (CLEAR_ATTENTION_WHEN_I_REPLY = False )
    If ATTENTION_ENABLED And (Not handled Or Not CLEAR_ATTENTION_WHEN_I_REPLY) Then
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

    If (Not handled Or Not CLEAR_REPLY_MARK_WHEN_I_REPLY) And Not CatListHas(itm.Categories, CATEGORY_NAME) Then
        Set mine = FindRepliedMine(itm)
        If Not mine Is Nothing Then
            MarkMail itm
            If FOLLOWUP_ENABLED Then RemoveCategoryName mine, FOLLOWUP_CATEGORY, True
            If notify And Not (urgentHit And URGENT_POPUP) Then NotifyReply itm
            ProcessMail = True
        End If
    End If

    If notify And attnMarked And Not ProcessMail And Not urgentHit And ATTENTION_TOAST Then NotifyAttention itm
    If notify And urgentHit And URGENT_POPUP Then NotifyUrgent itm, urgentKw, ProcessMail
    Exit Function
Fail:
    Debug.Print U("ReplyAlert.ProcessMail \uC624\uB958: ") & Err.Description
End Function

' (Table) ProcessMail. sliced = . True
Private Function ScanInboxTable(ByVal days As Long, ByVal sliced As Boolean, _
                                ByRef replies As Long, ByRef attns As Long, Optional ByVal fld As Object) As Boolean
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date
    Dim eid As String, itm As Object, attn As Boolean, cnt As Long
    On Error GoTo Fail
    EnsureDicts
    cutoff = Now - days
    If fld Is Nothing Then Set fld = Application.Session.GetDefaultFolder(olFolderInbox)
    Set tbl = fld.GetTable
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
                If SliceOver() Then Exit Function
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
    Debug.Print U("ReplyAlert.ScanInboxTable \uC624\uB958: ") & Err.Description
    ScanInboxTable = True
End Function

' " " ( Nothing)
Private Function FindRepliedMine(ByVal mail As Object) As Object
    Dim pa As Object, inReplyTo As String, refs As String, ids As Collection, i As Long
    Dim key As String, f As Object, conv As String, refIds As Collection
    On Error GoTo Done
    EnsureMaps False
    ' (To) ((CC) )
    If REPLY_REQUIRE_TO Then
        If Not IsToMe(mail) Then GoTo Done
    End If
    Set pa = mail.PropertyAccessor
    inReplyTo = GetPropStr(pa, PR_IN_REPLY_TO_ID)
    If MATCH_WHOLE_THREAD Then refs = GetPropStr(pa, PR_INTERNET_REFERENCES)

    ' 1) ID
    Set ids = SplitIds(inReplyTo & " " & refs)
    For i = 1 To ids.Count
        key = LCase$(ids(i))
        If m_SentIds.Exists(key) Then
            Set f = GetItemSafe(m_SentIds(key))
            If f Is Nothing Then Set f = FindSentByMessageId(key)
            If Not f Is Nothing Then
                Set FindRepliedMine = f
                Exit Function
            End If
        End If
    Next i

    ' 1-2) : (References) , (To)
    ' (: -> To )
    ' (CC)
    If Not MATCH_WHOLE_THREAD And ids.Count > 0 Then
        Set refIds = SplitIds(GetPropStr(pa, PR_INTERNET_REFERENCES))
        For i = 1 To refIds.Count
            key = LCase$(refIds(i))
            If m_SentIds.Exists(key) Then
                If IsToMe(mail) Then
                    Set f = GetItemSafe(m_SentIds(key))
                    If f Is Nothing Then Set f = FindSentByMessageId(key)
                    If Not f Is Nothing Then
                        Set FindRepliedMine = f
                        Exit Function
                    End If
                End If
                Exit For
            End If
        Next i
    End If

    ' 2) [ ] ( RE:/FW: ) ,
    If REPLY_THREAD_MODE Then
        Dim keys(1 To 2) As String, kk As Long, rcv As Date, snd As String, topic As String, idList() As String
        conv = SafeConvId(mail)
        topic = mail.ConversationTopic
        If Len(topic) = 0 Then topic = mail.Subject
        topic = NormTopic(topic)
        If Len(conv) > 0 Then keys(1) = "c:" & conv
        If Len(topic) >= 4 Then keys(2) = "t:" & topic
        rcv = GetReceived(mail)
        snd = GetSenderSmtp(mail)
        For kk = 1 To 2
            If Len(keys(kk)) > 0 Then
                If m_MyThreadFirst.Exists(keys(kk)) Then
                    If m_MyThreadFirst(keys(kk)) < rcv Then
                        If Not REPLY_REQUIRE_MY_RECIPIENT Or ThreadRecipHas(keys(kk), snd) Then
                            idList = Split(m_MyThreadIds(keys(kk)), "|")
                            If UBound(idList) >= 1 Then Set FindRepliedMine = GetItemSafe(idList(1))
                            If FindRepliedMine Is Nothing Then Set FindRepliedMine = mail
                            GoTo Done
                        End If
                    End If
                End If
            End If
        Next kk
        GoTo Done
    End If

    ' 2) (In-Reply-To) ( ):
    ' , (To) ,
    If ids.Count = 0 Then
        conv = SafeConvId(mail)
        If Len(conv) > 0 Then
            If m_MyLastTime.Exists(conv) Then
                If m_MyLastTime(conv) < GetReceived(mail) Then
                    If IsToMe(mail) Then
                        Set f = GetItemSafe(m_MyLastId(conv))
                        If Not f Is Nothing Then
                            If WasRecipientOf(f, GetSenderSmtp(mail)) Then Set FindRepliedMine = f
                        End If
                    End If
                End If
            End If
        End If
    End If
Done:
End Function

' myMail (To/CC) addr
Private Function WasRecipientOf(ByVal myMail As Object, ByVal addr As String) As Boolean
    Dim r As Object
    On Error Resume Next
    addr = LCase$(Trim$(addr))
    If Len(addr) = 0 Then Exit Function
    For Each r In myMail.Recipients
        If RecipientSmtp(r) = addr Then
            WasRecipientOf = True
            Exit Function
        End If
    Next r
End Function

Public Sub RecheckReplyMarks()
    Dim n As Long
    EnsureInit
    If m_Busy Then
        MsgBoxU U("\uC9C0\uAE08 \uB2E4\uB978 \uAC80\uC0AC\uB97C \uD558\uB294 \uC911\uC785\uB2C8\uB2E4. \uC7A0\uC2DC \uD6C4 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694."), vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    EnsureMaps True
    n = RecheckReplyMarksCore()
    m_Busy = False
    MsgBoxU "[" & CATEGORY_NAME & U("] \uD45C\uC2DC\uB97C \uB2E4\uC2DC \uD655\uC778\uD574\uC11C ") & n & U("\uAC74\uC744 \uBAA9\uB85D\uC5D0\uC11C \uBE90\uC2B5\uB2C8\uB2E4."), vbInformation, APP_TITLE
End Sub

Private Function RecheckReplyMarksCore() As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date, itm As Object, n As Long
    Dim C_EID As Long, C_RCV As Long, C_CATS As Long
    On Error GoTo Done
    cutoff = Now - REPLY_TRACK_DAYS
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_CATS = AddCol(tbl, "Categories")
    If C_EID < 0 Or C_RCV < 0 Or C_CATS < 0 Then GoTo Done
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Done
            If CatListHas(VStr(AV(arr, r, C_CATS)), CATEGORY_NAME) Then
                Set itm = GetItemSafe(VStr(AV(arr, r, C_EID)))
                If Not itm Is Nothing Then
                    If FindRepliedMine(itm) Is Nothing Then
                        RemoveCategoryName itm, CATEGORY_NAME, True
                        n = n + 1
                    End If
                End If
            End If
        Next r
        DoEvents
    Loop
Done:
    RecheckReplyMarksCore = n
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
'==========================================================================
Private Function AttentionMatch(ByVal mail As Object) As String
    Dim body As String, lines() As String, i As Long, n As Long, ln As String
    On Error GoTo Done
    If Len(Trim$(ATTENTION_NAMES)) = 0 Then Exit Function
    body = Left$(mail.Body, 3000)
    body = Replace(body, ChrW(160), " ")
    body = Replace(body, ChrW(8203), "")
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

' - (Hi, Dear, ...) O "Hi Jennifer," " ,"
' - O "Hi Tom and Jennifer,"
Private Function LineAddressesMe(ByVal ln As String) As Boolean
    Dim s As String, hadGreeting As Boolean, changed As Boolean
    Dim greetings() As String, names() As String, i As Long, g As String, nm As String

    hadGreeting = (Left$(LTrim$(ln), 1) = "@")      ' @Jennifer
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
                    ' O "Jennifer, please.." " ," " "
                    ' X "Jennifer will send.." " .."
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

Private Function AddressTail(ByVal r As String) As Boolean
    Dim p As Long
    r = LTrim$(r)
    p = InStr(1, r, U("\uB2D8"), vbBinaryCompare)
    If p > 0 And p <= 6 Then
        r = LTrim$(Mid$(r, p + 1))
        If Left$(r, 1) = U("\uAED8") And Mid$(r, 2, 1) <> U("\uC11C") Then r = LTrim$(Mid$(r, 2))
    End If
    If Len(r) = 0 Then AddressTail = True: Exit Function
    If InStr(1, ",.!:;~-)" & ChrW(65292), Left$(r, 1), vbBinaryCompare) > 0 Then AddressTail = True: Exit Function
    If Left$(r, 2) = U("\uC548\uB155") Then AddressTail = True
End Function

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

' / (: "to" vs "today").
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
    ShowToast U("\uB0B4 \uC55E\uC73C\uB85C \uC628 \uBA54\uC77C"), mail.SenderName, mail.Subject, openUri, _
              U("\uBA54\uC77C \uC5F4\uAE30|") & openUri & U(";\uBAA9\uB85D \uBCF4\uAE30|") & URI_SCHEME & ":attn"
End Sub


'==========================================================================
'==========================================================================
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
    purgeCut = Now - (TODO_LOOKBACK_DAYS + 3)

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

            If handled Then
                If hasReply And CLEAR_REPLY_MARK_WHEN_I_REPLY Then RemoveCategoryById eid, CATEGORY_NAME
                If hasAttn And CLEAR_ATTENTION_WHEN_I_REPLY Then RemoveCategoryById eid, ATTENTION_CATEGORY
            End If

            If rcv = 0 Or rcv > newCut Then GoTo NextRow
            If rcv < oldCut Then
                If hasTodo Then RemoveCategoryById eid, TODO_CATEGORY
                GoTo NextRow
            End If

            need = Not handled
            If CatListHas(cats, DONE_CATEGORY) Then need = False
            If VLong(AV(arr, r, C_FLAG)) = FLAG_COMPLETE Then need = False
            If need And IsMeAddress(VStr(AV(arr, r, C_SENDER))) Then need = False
            toMe = AV(arr, r, C_TOME)
            If need And Not hasAttn Then
                If TODO_ONLY_TO_ME And VarType(toMe) = vbBoolean Then
                    If Not toMe Then need = False                     ' (CC)/
                End If
                If CountNames(VStr(AV(arr, r, C_TO))) + CountNames(VStr(AV(arr, r, C_CC))) > TODO_MAX_RECIPIENTS Then need = False
            End If

            If need And Not hasTodo Then
                If m_TodoChecked.Exists(eid) Then
                    need = (Len(m_TodoChecked(eid)) = 0)
                Else
                    If sliced Then
                        If SliceOver() Then
                            m_TodoIncomplete = True
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
    If Err.Number <> 0 Then Debug.Print U("ReplyAlert.ScanTodoFolder \uC624\uB958: ") & Err.Description
    ScanTodoFolder = total
End Function

Private Function TodoItemCheck(ByVal itm As Object, ByVal attn As Boolean, ByVal checkToMe As Boolean) As String
    On Error GoTo Fail
    If itm Is Nothing Then TodoItemCheck = U("\uBA54\uC77C\uC744 \uC5F4 \uC218 \uC5C6\uC74C"): Exit Function
    If IsAutoReply(itm) Then TodoItemCheck = U("\uC790\uB3D9\uD68C\uC2E0"): Exit Function
    If IsBulkMail(itm) Then TodoItemCheck = U("\uC790\uB3D9 \uBC1C\uC1A1/\uB2E8\uCCB4 \uBA54\uC77C"): Exit Function
    If checkToMe And TODO_ONLY_TO_ME And Not attn Then
        If Not IsToMe(itm) Then TodoItemCheck = U("\uBC1B\uB294 \uC0AC\uB78C(To)\uC5D0 \uB0B4\uAC00 \uC5C6\uC74C (\uCC38\uC870/\uB2E8\uCCB4 \uC8FC\uC18C)"): Exit Function
    End If
    Exit Function
Fail:
    TodoItemCheck = U("\uD655\uC778 \uC911 \uC624\uB958")
End Function

Private Function TodoReason(ByVal itm As Object) As String
    Dim lv As Long, n As Long, attn As Boolean
    On Error GoTo Fail
    EnsureMaps False
    If itm.Class <> olMail Then TodoReason = U("\uC77C\uBC18 \uBA54\uC77C \uC544\uB2D8"): Exit Function
    If Left$(LCase$(itm.MessageClass), 14) = "ipm.note.rules" Then TodoReason = U("\uC790\uB3D9 \uBA54\uC77C"): Exit Function
    If IsDone(itm) Then TodoReason = U("\uC644\uB8CC \uCC98\uB9AC\uB428"): Exit Function
    If IsFromMe(itm) Then TodoReason = U("\uB0B4\uAC00 \uBCF4\uB0B8 \uBA54\uC77C"): Exit Function

    lv = GetPropLong(itm.PropertyAccessor, PR_LAST_VERB_EXECUTED)
    If lv = VERB_REPLY Or lv = VERB_REPLY_ALL Then TodoReason = U("\uC774\uBBF8 \uD68C\uC2E0\uD568"): Exit Function
    If lv = VERB_FORWARD Then TodoReason = U("\uC774\uBBF8 \uC804\uB2EC\uD568"): Exit Function
    If GetPropLong(itm.PropertyAccessor, PR_FLAG_STATUS) = FLAG_COMPLETE Then TodoReason = U("\uD50C\uB798\uADF8 \uC644\uB8CC"): Exit Function
    If IRepliedInThread(SafeConvId(itm), GetReceived(itm)) Then TodoReason = U("\uAC19\uC740 \uC2A4\uB808\uB4DC\uC5D0 \uB0B4\uAC00 \uB098\uC911\uC5D0 \uBCF4\uB0B8 \uBA54\uC77C \uC788\uC74C"): Exit Function
    If IsAutoReply(itm) Then TodoReason = U("\uC790\uB3D9\uD68C\uC2E0"): Exit Function
    If IsBulkMail(itm) Then TodoReason = U("\uC790\uB3D9 \uBC1C\uC1A1/\uB2E8\uCCB4 \uBA54\uC77C"): Exit Function

    attn = CatListHas(itm.Categories, ATTENTION_CATEGORY)
    If Not attn And ATTENTION_ENABLED Then attn = (Len(AttentionMatch(itm)) > 0)
    If TODO_ONLY_TO_ME And Not attn Then
        If Not IsToMe(itm) Then TodoReason = U("\uBC1B\uB294 \uC0AC\uB78C(To)\uC5D0 \uB0B4\uAC00 \uC5C6\uC74C (\uCC38\uC870/\uB2E8\uCCB4 \uC8FC\uC18C)"): Exit Function
    End If
    n = itm.Recipients.Count
    If n > TODO_MAX_RECIPIENTS And Not attn Then TodoReason = U("\uBC1B\uB294 \uC0AC\uB78C\uC774 \uB9CE\uC740 \uBA54\uC77C (") & n & U("\uBA85)"): Exit Function
    TodoReason = ""
    Exit Function
Fail:
    TodoReason = U("\uD655\uC778 \uC911 \uC624\uB958")
End Function

Private Function IsToMe(ByVal itm As Object) As Boolean
    Dim r As Object, v As Variant
    On Error Resume Next
    v = Empty
    v = itm.PropertyAccessor.GetProperty(PR_MESSAGE_TO_ME)
    If VarType(v) = vbBoolean Then
        If v Then IsToMe = True: Exit Function
    End If
    ' False ( )
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
    snd = GetSenderSmtp(itm)
    ' (loungepm@ Microsoft 365 )
    If Not IsMyDomain(snd) Then
        If InStr(1, hdr, "list-unsubscribe:", vbBinaryCompare) > 0 Then IsBulkMail = True
        If InStr(1, hdr, "precedence: list", vbBinaryCompare) > 0 Then IsBulkMail = True
    End If
    If InStr(1, hdr, "precedence: bulk", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, hdr, "precedence: junk", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "noreply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "no-reply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "donotreply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "do-not-reply", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "mailer-daemon", vbBinaryCompare) > 0 Then IsBulkMail = True
    If InStr(1, snd, "postmaster", vbBinaryCompare) > 0 Then IsBulkMail = True
End Function

Private Function IsMyDomain(ByVal addr As String) As Boolean
    Dim p As Long, dom As String
    addr = LCase$(Trim$(addr))
    p = InStrRev(addr, "@")
    If p = 0 Then
        IsMyDomain = (Len(addr) > 0)       ' SMTP (Exchange)
        Exit Function
    End If
    dom = Mid$(addr, p)                    ' "@globalhansoll.com"
    IsMyDomain = (InStr(1, MyAddresses(), dom & ";", vbTextCompare) > 0)
End Function

' "; Kim, Tom; ..."
Private Function CountNames(ByVal s As String) As Long
    If Len(Trim$(s)) = 0 Then Exit Function
    CountNames = UBound(Split(s, ";")) + 1
End Function


'==========================================================================
'==========================================================================
' (Table) [ ] .
' , waitingTotal =
Private Function ScanFollowups(ByRef waitingTotal As Long, ByVal sliced As Boolean) As Long
    Dim C_EID As Long, C_CLASS As Long, C_SENT As Long, C_CATS As Long
    Dim C_MSGID As Long, C_IRT As Long, C_CONV As Long
    Dim tbl As Object, arr As Variant, r As Long, itm As Object
    Dim newCut As Date, oldCut As Date, purgeCut As Date, sentOn As Date
    Dim eid As String, cats As String, hasCat As Boolean, st As Long, nNew As Long

    On Error GoTo Done
    waitingTotal = 0
    m_FuIncomplete = False
    EnsureDicts
    newCut = Now - FOLLOWUP_DAYS
    oldCut = Now - FOLLOWUP_LOOKBACK_DAYS
    purgeCut = oldCut - 3
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
            If sentOn = 0 Or sentOn >= newCut Then GoTo NextRow

            eid = VStr(AV(arr, r, C_EID))
            cats = VStr(AV(arr, r, C_CATS))
            hasCat = CatListHas(cats, FOLLOWUP_CATEGORY)

            If sentOn < oldCut Then
                If hasCat Then RemoveCategoryById eid, FOLLOWUP_CATEGORY
                GoTo NextRow
            End If

            st = FollowupStateOf(VHex(AV(arr, r, C_CONV)), sentOn, NormId(VStr(AV(arr, r, C_MSGID))), _
                                 Len(VStr(AV(arr, r, C_IRT))) > 0, CatListHas(cats, DONE_CATEGORY))

            ' 6
            If st = FU_WAITING Then
                If m_FuChecked.Exists(eid) Then
                    If DateDiff("h", m_FuCheckedTime(eid), Now) < 6 Or m_FuChecked(eid) <> FU_WAITING Then
                        st = m_FuChecked(eid)
                        GoTo Decided
                    End If
                End If
                If sliced Then
                    If SliceOver() Then
                        m_FuIncomplete = True
                        GoTo Done
                    End If
                End If
                Set itm = GetItemSafe(eid)
                If Not itm Is Nothing Then st = ConversationFollowupState(itm)
                m_FuChecked(eid) = st
                m_FuCheckedTime(eid) = Now
                Set itm = Nothing
            End If
Decided:
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
    If Err.Number <> 0 Then Debug.Print U("ReplyAlert.ScanFollowups \uC624\uB958: ") & Err.Description
    ScanFollowups = nNew
End Function

' () ( ). FU_ANSWERED / FU_SUPERSEDED / FU_WAITING
' foundWhere =
Private Function ConversationFollowupState(ByVal sentItem As Object, Optional ByRef foundWhere As String) As Long
    Dim conv As Object, roots As Object, rt As Object, c As New Collection, it As Object
    Dim sent As Date, after As Date, selfId As String, superseded As Boolean
    On Error GoTo Done
    ConversationFollowupState = FU_WAITING
    foundWhere = ""
    sent = SafeSentOn(sentItem)
    If sent = 0 Then Exit Function
    after = DateAdd("s", 5, sent)
    selfId = SafeEntryId(sentItem)
    Set conv = sentItem.GetConversation
    If conv Is Nothing Then Exit Function
    Set roots = conv.GetRootItems
    For Each rt In roots
        CollectConversation conv, rt, c, 0
    Next rt
    For Each it In c
        If IsOtherMail(it, selfId) Then
            If ItemTime(it) > after Then
                If IsFromMe(it) Then
                    superseded = True
                ElseIf Not IsAutoReply(it) Then
                    foundWhere = SafeFolderName(it)
                    ConversationFollowupState = FU_ANSWERED
                    Exit Function
                End If
            End If
        End If
    Next it
    If superseded Then ConversationFollowupState = FU_SUPERSEDED
Done:
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

Private Function IsOtherMail(ByVal it As Object, ByVal selfId As String) As Boolean
    On Error GoTo Done
    If it.Class <> olMail Then Exit Function
    If it.EntryID = selfId Then Exit Function
    IsOtherMail = True
Done:
End Function

Private Function SafeFolderName(ByVal itm As Object) As String
    On Error Resume Next
    SafeFolderName = itm.Parent.Name
End Function

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
    s = GetPropStr(mail.PropertyAccessor, PR_SENDER_SMTP_ADDRESS)
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
        ShowToast U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0\uC774 \uC654\uC2B5\uB2C8\uB2E4"), sender, subj, openUri, _
                  U("\uBA54\uC77C \uC5F4\uAE30|") & openUri & U(";\uD68C\uC2E0 \uBAA9\uB85D \uBCF4\uAE30|") & URI_SCHEME & ":list"
    End If
    If SHOW_POPUP Then
        MsgBoxU U("\uB0B4 \uBA54\uC77C\uC5D0 \uD68C\uC2E0\uC774 \uC654\uC2B5\uB2C8\uB2E4.") & vbCrLf & vbCrLf & _
               U("\uBCF4\uB0B8 \uC0AC\uB78C: ") & sender & vbCrLf & U("\uC81C\uBAA9: ") & subj, _
               vbInformation + vbSystemModal, APP_TITLE
    End If
End Sub

' Windows 10/11 (PowerShell , )
' launchUri :
' buttons : "|;|" ( )
Private Sub ShowToast(ByVal title As String, ByVal line1 As String, ByVal line2 As String, _
                      Optional ByVal launchUri As String = "", Optional ByVal buttons As String = "", _
                      Optional ByVal silent As Boolean = False)
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
          IIf(silent, "<audio silent=""true""/>", "<audio src=""ms-winsoundevent:Notification.Mail""/>") & "</toast>"
    ps = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;" & _
         "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null;" & _
         "$x = New-Object Windows.Data.Xml.Dom.XmlDocument;" & _
         "$x.LoadXml('" & xml & "');" & _
         "$t = New-Object Windows.UI.Notifications.ToastNotification $x;" & _
         "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('" & TOAST_APP_ID & "').Show($t)"
    Shell "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " & B64Utf16(ps), vbHide
End Sub


'==========================================================================
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

' scopeFolderType ( ) category
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
    ' Outlook
    SaveSetting SETTINGS_APP, "Favorites2", FavKey(folderName), "pending"
    TryPendingFavorites
Done:
End Sub

Private Sub TryPendingFavorites()
    Dim names As Variant, i As Long
    On Error Resume Next
    names = Array(LIST_FOLDER_NAME, ATTENTION_LIST_NAME, TODO_LIST_NAME, FOLLOWUP_LIST_NAME, URGENT_LIST_NAME)
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
        MsgBoxU "[" & folderName & U("] \uBAA9\uB85D \uD3F4\uB354\uB97C \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4."), vbExclamation, APP_TITLE
        Exit Sub
    End If
    If Application.ActiveExplorer Is Nothing Then
        sf.Display
    Else
        Set Application.ActiveExplorer.CurrentFolder = sf
    End If
    TryPendingFavorites
End Sub

' True
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
' - replyalert: (HKCU) ( )
' - %APPDATA%\ReplyAlert\open.vbs Outlook /
'==========================================================================
Private Sub InstallClickHandler()
    Dim fso As Object, ts As Object, sh As Object
    Dim folderPath As String, scriptPath As String, regKey As String
    On Error GoTo Done
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Environ$("APPDATA") & "\ReplyAlert"
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    scriptPath = folderPath & "\open.vbs"

    Set ts = fso.CreateTextFile(scriptPath, True, True)
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

' open.vbs (~ )
Private Function HandlerScript() As String
    Dim s As String
    s = U("' ReplyAlert \uC54C\uB9BC \uD074\uB9AD \uCC98\uB9AC\uAE30 - Outlook \uB9E4\uD06C\uB85C ReplyAlert \uAC00 \uC790\uB3D9 \uC0DD\uC131") & vbCrLf & _
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
        U("' \uBA54\uC77C\uC744 \uBABB \uCC3E\uC558\uAC70\uB098(\uB2E4\uB978 \uD3F4\uB354\uB85C \uC62E\uAE40 \uB4F1) \uBAA9\uB85D \uBC84\uD2BC\uC774\uBA74 \uBAA9\uB85D \uD3F4\uB354 \uC5F4\uAE30") & vbCrLf & _
        "If Not found Then" & vbCrLf & _
        "  target = ~<<LIST>>~" & vbCrLf & _
        "  If LCase(uri) = ~attn~ Then target = ~<<ATTN>>~" & vbCrLf & _
        "  If LCase(uri) = ~todo~ Then target = ~<<TODO>>~" & vbCrLf & _
        "  If LCase(uri) = ~urgent~ Then target = ~<<URGENT>>~" & vbCrLf & _
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
    s = Replace(s, "<<URGENT>>", Replace(URGENT_LIST_NAME, "~", ""))
    s = Replace(s, "<<FOLLOWUP>>", Replace(FOLLOWUP_LIST_NAME, "~", ""))
    HandlerScript = Replace(s, "~", Chr$(34))
End Function


'==========================================================================
'  Strings / Unicode helpers
'==========================================================================
Private Sub EnsureInit()
    On Error Resume Next
    If m_Inited Then Exit Sub
    APP_TITLE = U("\uBA54\uC77C \uC54C\uB9AC\uBBF8")
    URGENT_CATEGORY = U("\uAE34\uAE09")
    URGENT_LIST_NAME = URGENT_CATEGORY
    CATEGORY_NAME = U("\uB0B4 \uBA54\uC77C \uD68C\uC2E0")
    LIST_FOLDER_NAME = CATEGORY_NAME
    ATTENTION_CATEGORY = U("\uB0B4 \uC55E\uC73C\uB85C \uC628 \uBA54\uC77C")
    ATTENTION_LIST_NAME = ATTENTION_CATEGORY
    TODO_CATEGORY = U("\uB0B4 \uB2F5\uC7A5 \uD544\uC694")
    TODO_LIST_NAME = TODO_CATEGORY
    FOLLOWUP_CATEGORY = U("\uB2F5\uC7A5 \uB300\uAE30")
    FOLLOWUP_LIST_NAME = FOLLOWUP_CATEGORY
    DONE_CATEGORY = U("\uC644\uB8CC \uCC98\uB9AC")
    ATTENTION_GREETINGS = "hi;hello;hey;dear;good morning;good afternoon;good evening;morning;greetings;attn;attention;to;" & _
                          U("\uC548\uB155\uD558\uC138\uC694;\uC548\uB155\uD558\uC2ED\uB2C8\uAE4C;\uC218\uC2E0;\uBC1B\uB294 \uC0AC\uB78C;\uC874\uACBD\uD558\uB294")
    ATTENTION_NAMES = LoadNames()
    LoadPeople
    If ATTENTION_AUTO_NAMES Then ATTENTION_NAMES = ATTENTION_NAMES & ";" & AutoNames()
    LoadKeywords
    m_Inited = True
End Sub

' : %APPDATA%\ReplyAlert\names.txt ( , )
Private Function NamesFilePath() As String
    NamesFilePath = Environ$("APPDATA") & "\ReplyAlert\names.txt"
End Function

Private Function LoadNames() As String
    Dim txt As String, defaults As String, lines() As String, i As Long, ln As String, result As String
    defaults = "Jennifer" & vbCrLf & U("\uAE40\uB098\uD604") & vbCrLf & U("\uAE40 \uB098\uD604") & vbCrLf
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

' Outlook
' ) "Kim, Na Hyun (Jennifer)" -> Jennifer; Na Hyun; Nahyun; Na-Hyun; Ms. Kim; Ms Kim
Private Function AutoNames() As String
    Dim eu As Object, disp As String, firstN As String, lastN As String, nick As String, p As Long, q As Long, out As String
    On Error Resume Next
    disp = Application.Session.CurrentUser.Name
    Set eu = Application.Session.CurrentUser.AddressEntry.GetExchangeUser
    If Not eu Is Nothing Then
        firstN = Trim$(eu.FirstName)
        lastN = Trim$(eu.LastName)
    End If
    p = InStr(1, disp, "(")
    If p > 0 Then
        q = InStr(p + 1, disp, ")")
        If q > p Then nick = Trim$(Mid$(disp, p + 1, q - p - 1))
        disp = Trim$(Left$(disp, p - 1))
    End If
    If Len(firstN) = 0 Then
        p = InStr(1, disp, ",")
        If p > 0 Then
            lastN = Trim$(Left$(disp, p - 1))
            firstN = Trim$(Mid$(disp, p + 1))
        End If
    End If
    If IsAsciiText(nick) And Len(nick) > 1 Then out = out & nick & ";"
    If IsAsciiText(firstN) And Len(firstN) > 1 Then
        out = out & firstN & ";"
        If InStr(firstN, " ") > 0 Or InStr(firstN, "-") > 0 Then
            out = out & Replace(Replace(firstN, " ", ""), "-", "") & ";"
            out = out & Replace(firstN, " ", "-") & ";"
            out = out & Replace(firstN, "-", " ") & ";"
        End If
    End If
    If IsAsciiText(lastN) And Len(lastN) > 1 Then out = out & "Ms. " & lastN & ";Ms " & lastN & ";"
    AutoNames = out
End Function

Private Function IsAsciiText(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If AscW(Mid$(s, i, 1)) > 127 Or AscW(Mid$(s, i, 1)) < 0 Then Exit Function
    Next i
    IsAsciiText = True
End Function

Private Function FirstLines(ByVal mail As Object, ByVal n As Long) As String
    Dim lines() As String, i As Long, k As Long, ln As String, out As String
    On Error Resume Next
    lines = Split(Left$(mail.Body, 2000), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(Replace(Replace(lines(i), vbCr, ""), ChrW(160), " "))
        If Len(ln) > 0 Then
            out = out & " | " & Left$(ln, 40)
            k = k + 1
            If k >= n Then Exit For
        End If
    Next i
    FirstLines = Mid$(out, 4)
End Function

' ( Outlook )
' ==========================================================================
' ReplyReport : ()
' - Outlook ( )
' ==========================================================================
' ==========================================================================
' : people.txt /
' ==========================================================================
Private Function PeopleFilePath() As String
    PeopleFilePath = Environ$("APPDATA") & "\ReplyAlert\people.txt"
End Function

Private Sub LoadPeople()
    Dim txt As String, defaults As String, lines() As String, i As Long, ln As String, p As Long
    Dim lbl As String, kws() As String, j As Long, kw As String, joined As String
    On Error Resume Next
    defaults = U("# \uD55C \uC904\uC5D0 \uD55C \uC0AC\uB78C:  \uBAA9\uB85D \uC774\uB984 = \uCC3E\uC744 \uB2E8\uC5B4; \uCC3E\uC744 \uB2E8\uC5B4; ...") & vbCrLf & _
               U("# \uCC3E\uC744 \uB2E8\uC5B4\uB294 Outlook \uC5D0 \uBCF4\uC774\uB294 \uC774\uB984\uC774\uB098 \uBA54\uC77C \uC8FC\uC18C \uC77C\uBD80 (\uB300\uC18C\uBB38\uC790 \uBB34\uC2DC)") & vbCrLf & _
               U("\uD574\uB9AC \uACFC\uC7A5 = Hae Ri; Hailey") & vbCrLf & _
               U("\uC131\uC9C4 \uB300\uB9AC = Sung Jin; Seong Jin; Sungjin; Seongjin; \uC131\uC9C4") & vbCrLf
    txt = ReadUtf8(PeopleFilePath())
    If Len(Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))) = 0 Then
        WriteUtf8 PeopleFilePath(), defaults
        txt = defaults
    End If
    m_PeopleCount = 0
    m_PeopleSetsReady = False
    ReDim m_PeopleLabel(1 To 50)
    ReDim m_PeopleKw(1 To 50)
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(Replace(lines(i), ChrW(65279), ""))
        p = InStr(1, ln, "=")
        If Len(ln) > 0 And Left$(ln, 1) <> "#" And p > 1 And m_PeopleCount < 50 Then
            lbl = Trim$(Left$(ln, p - 1))
            kws = Split(Mid$(ln, p + 1), ";")
            joined = ""
            For j = LBound(kws) To UBound(kws)
                kw = LCase$(Trim$(kws(j)))
                If Len(kw) > 1 Then joined = joined & "|" & kw
            Next j
            If Len(lbl) > 0 And Len(joined) > 0 Then
                m_PeopleCount = m_PeopleCount + 1
                m_PeopleLabel(m_PeopleCount) = lbl
                m_PeopleKw(m_PeopleCount) = joined & "|"
            End If
        End If
    Next i
End Sub

' ( TagPeopleNow Outlook )
' ==========================================================================
' [] : %APPDATA%\ReplyAlert\keywords.txt ( , EditKeywords )
' ==========================================================================
Private Function KeywordsFilePath() As String
    KeywordsFilePath = Environ$("APPDATA") & "\ReplyAlert\keywords.txt"
End Function

Private Sub LoadKeywords()
    Dim txt As String, defaults As String, lines() As String, i As Long, ln As String
    On Error Resume Next
    defaults = U("# \uD55C \uC904\uC5D0 \uD558\uB098. \uC81C\uBAA9\uC774\uB098 \uBCF8\uBB38 \uC55E\uBD80\uBD84\uC5D0 \uC788\uC73C\uBA74 [\uAE34\uAE09] \uD45C\uC2DC + \uD31D\uC5C5 (\uB300\uC18C\uBB38\uC790 \uBB34\uC2DC)") & vbCrLf & _
               U("# \uC601\uC5B4 \uB2E8\uC5B4\uB294 \uB4A4\uC5D0 \uBD99\uB294 \uD615\uD0DC\uB3C4 \uC7A1\uC74C (delay -> delayed, reject -> rejected)") & vbCrLf & _
               "urgent" & vbCrLf & "asap" & vbCrLf & "immediately" & vbCrLf & "delay" & vbCrLf & "claim" & vbCrLf & _
               "shortage" & vbCrLf & "reject" & vbCrLf & "cancel" & vbCrLf & _
               U("\uAE34\uAE09") & vbCrLf & U("\uAE09\uD568") & vbCrLf & U("\uAE09\uAC74") & vbCrLf & U("\uC9C0\uC5F0") & vbCrLf & U("\uBD88\uB7C9") & vbCrLf & U("\uD074\uB808\uC784") & vbCrLf
    txt = ReadUtf8(KeywordsFilePath())
    If Len(Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))) = 0 Then
        WriteUtf8 KeywordsFilePath(), defaults
        txt = defaults
    End If
    m_UrgentKw = "|"
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = LCase$(Trim$(Replace(lines(i), ChrW(65279), "")))
        If Len(ln) > 1 And Left$(ln, 1) <> "#" Then m_UrgentKw = m_UrgentKw & ln & "|"
    Next i
End Sub

Public Sub EditKeywords()
    EnsureInit
    Shell "notepad.exe """ & KeywordsFilePath() & """", vbNormalFocus
End Sub

Private Function UrgentMatch(ByVal itm As Object) As String
    Dim s As String, kws() As String, i As Long
    On Error Resume Next
    If Len(m_UrgentKw) < 3 Then Exit Function
    s = LCase$(itm.Subject & vbLf & Left$(NewPartOfBody(itm), 1500))
    kws = Split(m_UrgentKw, "|")
    For i = LBound(kws) To UBound(kws)
        If Len(kws(i)) > 0 Then
            If KeywordHit(s, kws(i)) Then
                UrgentMatch = kws(i)
                Exit Function
            End If
        End If
    Next i
End Function

' (reclaim X, claims O)
Private Function KeywordHit(ByVal s As String, ByVal kw As String) As Boolean
    Dim p As Long, prevChar As String
    p = InStr(1, s, kw, vbBinaryCompare)
    Do While p > 0
        If p > 1 Then prevChar = Mid$(s, p - 1, 1) Else prevChar = ""
        If Not (LCase$(Left$(kw, 1)) Like "[a-z0-9]") Or Not (prevChar Like "[a-z0-9]") Then
            KeywordHit = True
            Exit Function
        End If
        p = InStr(p + 1, s, kw, vbBinaryCompare)
    Loop
End Function

Private Sub NotifyUrgent(ByVal mail As Object, ByVal kw As String, ByVal isReply As Boolean)
    Dim openUri As String
    On Error Resume Next
    If Not SHOW_TOAST Then Exit Sub
    openUri = URI_SCHEME & ":open/" & mail.EntryID
    ShowToast U("\uAE34\uAE09 \uBA54\uC77C") & IIf(isReply, U(" (\uB0B4 \uBA54\uC77C \uD68C\uC2E0)"), ""), mail.SenderName, "[" & kw & "] " & mail.Subject, openUri, _
              U("\uBA54\uC77C \uC5F4\uAE30|") & openUri & U(";\uAE34\uAE09 \uBAA9\uB85D|") & URI_SCHEME & ":urgent", True
End Sub

' (From: / : / Original Message)
Private Function NewPartOfBody(ByVal mail As Object, Optional ByVal maxLen As Long = 6000) As String
    Dim b As String, lb As String, markers() As String, i As Long, p As Long, best As Long
    On Error Resume Next
    b = Left$(mail.Body, maxLen * 2)
    lb = LCase$(b)
    markers = Split(vbLf & "from:|" & vbLf & U("\uBCF4\uB0B8 \uC0AC\uB78C:|") & vbLf & "sent:|-----original message|" & vbLf & U("\uBC1C\uC2E0:|________________________________"), "|")
    best = 0
    For i = LBound(markers) To UBound(markers)
        p = InStr(1, lb, markers(i), vbBinaryCompare)
        If p > 20 Then
            If best = 0 Or p < best Then best = p
        End If
    Next i
    If best > 0 Then b = Left$(b, best - 1)
    NewPartOfBody = Left$(b, maxLen)
End Function


' ==========================================================================
' [] SetDeadline : () +
' ==========================================================================
Public Sub SetDeadline()
    Dim itm As Object, cands As Collection, c As Variant, ans As Long, d As Date, v As String, i As Long
    EnsureInit
    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBoxU U("\uBA54\uC77C\uC744 \uD558\uB098 \uC120\uD0DD\uD55C \uB4A4 \uC2E4\uD589\uD558\uC138\uC694."), vbExclamation
        Exit Sub
    End If
    If itm.Class <> olMail Then
        MsgBoxU U("\uC120\uD0DD\uD55C \uD56D\uBAA9\uC774 \uC77C\uBC18 \uBA54\uC77C\uC774 \uC544\uB2D9\uB2C8\uB2E4."), vbExclamation
        Exit Sub
    End If

    Set cands = FindDeadlines(itm, False)
    For i = 1 To cands.Count
        c = cands(i)
        ans = MsgBoxU(U("\uC774 \uBA54\uC77C\uC758 \uB9C8\uAC10\uC77C\uC744 ") & FmtDateKo(CDate(c(0))) & U(" \uB85C \uB4F1\uB85D\uD560\uAE4C\uC694?") & vbCrLf & vbCrLf & _
                      U("\uCC3E\uC740 \uBD80\uBD84: ...") & c(2) & "..." & vbCrLf & vbCrLf & _
                      U("\uC608 = \uB4F1\uB85D / \uC544\uB2C8\uC624 = \uB2E4\uC74C \uD6C4\uBCF4 / \uCDE8\uC18C = \uADF8\uB9CC"), vbYesNoCancel + vbQuestion)
        If ans = 6 Then
            RegisterDeadline itm, CDate(c(0))
            MsgBoxU U("\uB9C8\uAC10\uC77C ") & FmtDateKo(CDate(c(0))) & U(" \uB85C \uB4F1\uB85D\uD588\uC2B5\uB2C8\uB2E4.") & vbCrLf & ReminderText(CDate(c(0))), vbInformation
            Exit Sub
        End If
        If ans = 2 Then Exit Sub
        If i >= 5 Then Exit For
    Next i

    ' (InputBox )
    v = InputBox(IIf(cands.Count = 0, "No date found in this mail." & vbCrLf, "") & _
                 "Type the due date (e.g. 11/30, 2026-12-05, Dec 5). Leave empty to cancel.", "ReplyAlert", "")
    If Len(Trim$(v)) = 0 Then Exit Sub
    Set cands = FindDatesInText(v, False)
    If cands.Count = 0 Then
        MsgBoxU U("\uB0A0\uC9DC\uB97C \uC54C\uC544\uBCFC \uC218 \uC5C6\uC2B5\uB2C8\uB2E4: ") & v, vbExclamation
        Exit Sub
    End If
    d = CDate(cands(1)(0))
    RegisterDeadline itm, d
    MsgBoxU U("\uB9C8\uAC10\uC77C ") & FmtDateKo(d) & U(" \uB85C \uB4F1\uB85D\uD588\uC2B5\uB2C8\uB2E4.") & vbCrLf & ReminderText(d), vbInformation
End Sub

Private Sub RegisterDeadline(ByVal itm As Object, ByVal d As Date)
    Dim remind As Date
    On Error Resume Next
    d = DateValue(d)
    itm.MarkAsTask olMarkNoDate
    If d < Date Then itm.TaskStartDate = d Else itm.TaskStartDate = Date
    itm.TaskDueDate = d
    itm.FlagRequest = U("\uB9C8\uAC10 ") & Format$(d, "mm/dd")
    remind = ReminderFor(d)
    If remind > Now Then
        itm.ReminderSet = True
        itm.ReminderTime = remind
    End If
    itm.Save
End Sub

Private Function ReminderFor(ByVal d As Date) As Date
    ReminderFor = DateAdd("d", -DEADLINE_REMIND_DAYS_BEFORE, d) + TimeSerial(DEADLINE_REMIND_HOUR, 0, 0)
    If ReminderFor < Now Then ReminderFor = d + TimeSerial(DEADLINE_REMIND_HOUR, 0, 0)
End Function

Private Function ReminderText(ByVal d As Date) As String
    Dim r As Date
    r = ReminderFor(d)
    If r > Now Then
        ReminderText = U("\uBBF8\uB9AC \uC54C\uB9BC: ") & FmtDateKo(r) & " " & Format$(r, "hh:nn")
    Else
        ReminderText = U("(\uBBF8\uB9AC \uC54C\uB9BC \uC2DC\uAC01\uC774 \uC774\uBBF8 \uC9C0\uB098\uC11C \uC54C\uB9BC\uC740 \uC124\uC815\uD558\uC9C0 \uC54A\uC558\uC2B5\uB2C8\uB2E4)")
    End If
End Function

Private Function FmtDateKo(ByVal d As Date) As String
    FmtDateKo = Format$(d, "yyyy-mm-dd") & " (" & Split(U("\uC77C,\uC6D4,\uD654,\uC218,\uBAA9,\uAE08,\uD1A0"), ",")(Weekday(d, vbSunday) - 1) & ")"
End Function

Private Function FindDeadlines(ByVal itm As Object, ByVal requireKeyword As Boolean) As Collection
    Dim txt As String
    On Error Resume Next
    txt = itm.Subject & vbLf & NewPartOfBody(itm, 4000)
    Set FindDeadlines = FindDatesInText(txt, requireKeyword)
End Function

' Array(, , ) ( , )
Private Function FindDatesInText(ByVal txt As String, ByVal requireKeyword As Boolean) As Collection
    Dim re As Object, ms As Object, m As Object, pats(1 To 6) As String, k As Long
    Dim d As Date, score As Long, byDate As Object, ctx As Object, scr As Object, key As Variant, out As New Collection
    Dim keysArr() As Date, scoreArr() As Long, n As Long, i As Long, j As Long, tD As Date, tS As Long
    Dim mon As String
    Set FindDatesInText = out
    On Error Resume Next
    Set re = CreateObject("VBScript.RegExp")
    If re Is Nothing Then Exit Function
    re.Global = True
    re.IgnoreCase = True
    re.MultiLine = True
    mon = "(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?"
    pats(1) = "(\d{4})[-./](\d{1,2})[-./](\d{1,2})"
    pats(2) = "(^|[^A-Za-z0-9/.-])(\d{1,2})/(\d{1,2})(/(\d{2,4}))?(?![A-Za-z0-9/])"
    pats(3) = "(^|[^A-Za-z0-9])(\d{1,2})(st|nd|rd|th)?[- ]?" & mon & "([- ,]+(\d{4}))?(?![A-Za-z0-9])"
    pats(4) = "(^|[^A-Za-z])" & mon & "[ -]?(\d{1,2})(st|nd|rd|th)?(?![0-9])(,? ?(\d{4}))?"
    pats(5) = U("(\d{1,2})\s*\uC6D4\s*(\d{1,2})\s*\uC77C")
    pats(6) = U("(^|[^A-Za-z])(today|tomorrow|eod|eow|this week|next week|monday|tuesday|wednesday|thursday|friday|mon|tue|wed|thu|fri)(?![A-Za-z])|(\uC624\uB298|\uB0B4\uC77C|\uBAA8\uB808|\uC774\uBC88 \uC8FC|\uC774\uBC88\uC8FC|\uB2E4\uC74C \uC8FC|\uB2E4\uC74C\uC8FC|\uC6D4\uC694\uC77C|\uD654\uC694\uC77C|\uC218\uC694\uC77C|\uBAA9\uC694\uC77C|\uAE08\uC694\uC77C)")
    Set byDate = NewDict(): Set ctx = NewDict(): Set scr = NewDict()

    For k = 1 To 6
        re.Pattern = pats(k)
        Set ms = re.Execute(txt)
        For Each m In ms
            d = MatchToDate(k, m)
            If d >= Date And d <= Date + 366 Then
                score = 1
                If DeadlineWordNear(txt, m.FirstIndex + 1, m.Length) Then score = 10
                If k = 6 And score < 10 Then score = 0          ' "Friday"
                If score > 0 And (score >= 10 Or Not requireKeyword) Then
                    key = CDbl(d)
                    If Not byDate.Exists(key) Then
                        byDate(key) = d
                        scr(key) = score
                        ctx(key) = Snippet(txt, m.FirstIndex + 1, m.Length)
                    ElseIf score > scr(key) Then
                        scr(key) = score
                        ctx(key) = Snippet(txt, m.FirstIndex + 1, m.Length)
                    End If
                End If
            End If
        Next m
    Next k

    n = byDate.Count
    If n = 0 Then Exit Function
    ReDim keysArr(1 To n)
    ReDim scoreArr(1 To n)
    i = 0
    For Each key In byDate.Keys
        i = i + 1
        keysArr(i) = byDate(key)
        scoreArr(i) = scr(key)
    Next key
    For i = 1 To n - 1
        For j = i + 1 To n
            If scoreArr(j) > scoreArr(i) Or (scoreArr(j) = scoreArr(i) And keysArr(j) < keysArr(i)) Then
                tD = keysArr(i): keysArr(i) = keysArr(j): keysArr(j) = tD
                tS = scoreArr(i): scoreArr(i) = scoreArr(j): scoreArr(j) = tS
            End If
        Next j
    Next i
    For i = 1 To n
        out.Add Array(keysArr(i), "", ctx(CDbl(keysArr(i))))
    Next i
End Function

Private Function MatchToDate(ByVal k As Long, ByVal m As Object) As Date
    Dim sm As Object, a As Long, b As Long, y As Long, w As String
    On Error GoTo Done
    Set sm = m.SubMatches
    Select Case k
        Case 1
            MatchToDate = MakeDate(CLng(sm(0)), CLng(sm(1)), CLng(sm(2)))
        Case 2
            a = CLng(sm(1)): b = CLng(sm(2))
            If Len(sm(4) & "") > 0 Then y = CLng(sm(4))
            If DATE_ORDER_MD Then
                If a > 12 And b <= 12 Then MatchToDate = MakeDate(y, b, a) Else MatchToDate = MakeDate(y, a, b)
            Else
                If b > 12 And a <= 12 Then MatchToDate = MakeDate(y, a, b) Else MatchToDate = MakeDate(y, b, a)
            End If
        Case 3
            If Len(sm(5) & "") > 0 Then y = CLng(sm(5))
            MatchToDate = MakeDate(y, MonthNum(sm(3)), CLng(sm(1)))
        Case 4
            If Len(sm(5) & "") > 0 Then y = CLng(sm(5))
            MatchToDate = MakeDate(y, MonthNum(sm(1)), CLng(sm(2)))
        Case 5
            MatchToDate = MakeDate(0, CLng(sm(0)), CLng(sm(1)))
        Case 6
            w = LCase$(Trim$(sm(1) & sm(2)))
            MatchToDate = RelativeDate(w)
    End Select
Done:
End Function

Private Function MonthNum(ByVal s As String) As Long
    MonthNum = (InStr(1, "janfebmaraprmayjunjulaugsepoctnovdec", Left$(LCase$(s), 3)) + 2) \ 3
End Function

' ( 60 )
Private Function MakeDate(ByVal y As Long, ByVal mo As Long, ByVal dd As Long) As Date
    Dim r As Date
    On Error GoTo Done
    If mo < 1 Or mo > 12 Or dd < 1 Or dd > 31 Then Exit Function
    If y > 0 And y < 100 Then y = y + 2000
    If y = 0 Then
        y = Year(Date)
        r = DateSerial(y, mo, dd)
        If r < Date - 60 Then y = y + 1
    End If
    r = DateSerial(y, mo, dd)
    If Day(r) <> dd Then Exit Function          ' 2/30
    MakeDate = r
Done:
End Function

Private Function RelativeDate(ByVal w As String) As Date
    Dim wd As Long, i As Long, names() As String
    Select Case w
        Case "today", "eod", U("\uC624\uB298"): RelativeDate = Date: Exit Function
        Case "tomorrow", U("\uB0B4\uC77C"): RelativeDate = Date + 1: Exit Function
        Case U("\uBAA8\uB808"): RelativeDate = Date + 2: Exit Function
        Case "eow", "this week", U("\uC774\uBC88 \uC8FC"), U("\uC774\uBC88\uC8FC"): RelativeDate = NextWeekday(6): Exit Function
        Case "next week", U("\uB2E4\uC74C \uC8FC"), U("\uB2E4\uC74C\uC8FC"): RelativeDate = NextWeekday(6) + 7: Exit Function
    End Select
    names = Split(U("monday|tuesday|wednesday|thursday|friday|mon|tue|wed|thu|fri|\uC6D4\uC694\uC77C|\uD654\uC694\uC77C|\uC218\uC694\uC77C|\uBAA9\uC694\uC77C|\uAE08\uC694\uC77C"), "|")
    For i = LBound(names) To UBound(names)
        If w = names(i) Then
            wd = (i Mod 5) + 2                      ' =2 ... =6 (vbSunday )
            RelativeDate = NextWeekday(wd)
            Exit Function
        End If
    Next i
End Function

Private Function NextWeekday(ByVal wd As Long) As Date
    NextWeekday = Date + ((wd - Weekday(Date, vbSunday) + 7) Mod 7)
End Function

' 25 / 8
Private Function DeadlineWordNear(ByVal txt As String, ByVal pos As Long, ByVal ln As Long) As Boolean
    Dim before As String, after As String, words() As String, i As Long, st As Long
    st = pos - 25
    If st < 1 Then st = 1
    before = LCase$(Mid$(txt, st, pos - st))
    after = LCase$(Mid$(txt, pos + ln, 8))
    words = Split("by|until|till|before|deadline|due|no later than|not later than|latest|etd|eta|within", "|")
    For i = LBound(words) To UBound(words)
        If KeywordHit(before, words(i)) Then DeadlineWordNear = True: Exit Function
    Next i
    words = Split(U("\uB9C8\uAC10|\uAE30\uD55C|\uAE4C\uC9C0|\uC804\uAE4C\uC9C0|\uC774\uC804|\uC774\uB0B4"), "|")
    For i = LBound(words) To UBound(words)
        If InStr(1, before, words(i), vbBinaryCompare) > 0 Or InStr(1, after, words(i), vbBinaryCompare) > 0 Then
            DeadlineWordNear = True
            Exit Function
        End If
    Next i
End Function

Private Function Snippet(ByVal txt As String, ByVal pos As Long, ByVal ln As Long) As String
    Dim st As Long, s As String
    st = pos - 30
    If st < 1 Then st = 1
    s = Mid$(txt, st, (pos - st) + ln + 20)
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Snippet = Trim$(s)
End Function

Public Sub EditPeople()
    EnsureInit
    Shell "notepad.exe """ & PeopleFilePath() & """", vbNormalFocus
End Sub

Private Function PersonHit(ByVal hay As String, ByVal idx As Long) As Boolean
    Dim kws() As String, j As Long
    kws = Split(m_PeopleKw(idx), "|")
    For j = LBound(kws) To UBound(kws)
        If Len(kws(j)) > 0 Then
            If InStr(1, hay, kws(j), vbBinaryCompare) > 0 Then
                PersonHit = True
                Exit Function
            End If
        End If
    Next j
End Function

' , (To) ( ID RE:/FW: )
' -> ((CC) )
Private Sub AddThread(ByVal idx As Long, ByVal conv As String, ByVal topic As String)
    If Len(conv) > 0 Then m_PeopleSet(idx)("c:" & conv) = True
    If Len(topic) >= 4 Then m_PeopleSet(idx)("t:" & topic) = True
End Sub

Private Function InThread(ByVal idx As Long, ByVal conv As String, ByVal topic As String) As Boolean
    If Len(conv) > 0 Then
        If m_PeopleSet(idx).Exists("c:" & conv) Then InThread = True: Exit Function
    End If
    If Len(topic) >= 4 Then InThread = m_PeopleSet(idx).Exists("t:" & topic)
End Function

' : RE:/FW:/:/:/[External]
Private Function NormTopic(ByVal s As String) As String
    Dim changed As Boolean, prefixes() As String, i As Long, pf As String
    s = LCase$(Trim$(s))
    prefixes = Split(U("re:|fw:|fwd:|re :|fw :|aw:|wg:|\uD68C\uC2E0:|\uC804\uB2EC:|\uB2F5\uC7A5:|[external]|[ext]|[\uC678\uBD80]"), "|")
    Do
        changed = False
        For i = LBound(prefixes) To UBound(prefixes)
            pf = prefixes(i)
            If Left$(s, Len(pf)) = pf Then
                s = Trim$(Mid$(s, Len(pf) + 1))
                changed = True
            End If
        Next i
    Loop While changed And Len(s) > 0
    NormTopic = s
End Function

' ( ) + . rows
Private Sub BuildPeopleSets(ByVal days As Long, ByVal rows As Collection)
    Dim folders As New Collection, fld As Object, inbox As Object, i As Long
    On Error Resume Next
    If m_PeopleCount = 0 Then Exit Sub
    ReDim m_PeopleSet(1 To m_PeopleCount)
    For i = 1 To m_PeopleCount
        Set m_PeopleSet(i) = NewDict()
    Next i
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    folders.Add inbox
    AddSubFolders inbox, folders, 0
    folders.Add Application.Session.GetDefaultFolder(olFolderSentMail)
    For Each fld In folders
        PeopleScanFolder fld, Now - days, rows
    Next fld
    m_PeopleSetsReady = True
End Sub

Private Sub PeopleScanFolder(ByVal fld As Object, ByVal cutoff As Date, ByVal rows As Collection)
    Dim tbl As Object, arr As Variant, r As Long, t As Date, sHay As String, tHay As String, conv As String, topic As String, i As Long
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_NAME As Long, C_SND As Long, C_TO As Long, C_CATS As Long
    Dim C_CONV As Long, C_TOPIC As Long, C_SUBJ As Long
    On Error GoTo Done
    Set tbl = fld.GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_NAME = AddCol(tbl, "SenderName")
    C_SND = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    C_TO = AddCol(tbl, PR_DISPLAY_TO)
    C_CATS = AddCol(tbl, "Categories")
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    C_TOPIC = AddCol(tbl, PR_CONVERSATION_TOPIC)
    C_SUBJ = AddCol(tbl, "Subject")
    If C_EID < 0 Or C_RCV < 0 Then GoTo Done
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Done
            If Left$(RowClass(arr, r, C_CLASS), 8) = "ipm.note" Then
                sHay = LCase$(VStr(AV(arr, r, C_NAME)) & "|" & VStr(AV(arr, r, C_SND)))
                tHay = LCase$(VStr(AV(arr, r, C_TO)))
                conv = VHex(AV(arr, r, C_CONV))
                topic = VStr(AV(arr, r, C_TOPIC))
                If Len(topic) = 0 Then topic = VStr(AV(arr, r, C_SUBJ))
                topic = NormTopic(topic)
                For i = 1 To m_PeopleCount
                    If PersonHit(sHay, i) Or PersonHit(tHay, i) Then AddThread i, conv, topic
                Next i
                If Not rows Is Nothing Then rows.Add Array(VStr(AV(arr, r, C_EID)), VStr(AV(arr, r, C_CATS)), conv, topic)
            End If
        Next r
        DoEvents
    Loop
Done:
End Sub

Private Sub TagPeople(ByVal itm As Object)
    Dim sHay As String, tHay As String, conv As String, topic As String, i As Long
    On Error Resume Next
    If m_PeopleCount = 0 Then Exit Sub
    If itm.Class <> olMail Then Exit Sub
    If Not m_PeopleSetsReady Then BuildPeopleSets 60, Nothing      ' Outlook
    sHay = LCase$(itm.SenderName & "|" & GetSenderSmtp(itm))
    tHay = LCase$(itm.To)
    conv = SafeConvId(itm)
    topic = itm.ConversationTopic
    If Len(topic) = 0 Then topic = itm.Subject
    topic = NormTopic(topic)
    For i = 1 To m_PeopleCount
        If PersonHit(sHay, i) Or PersonHit(tHay, i) Then AddThread i, conv, topic
        If InThread(i, conv, topic) Then
            EnsureCategory m_PeopleLabel(i), olCategoryColorGreen
            AddCategoryName itm, m_PeopleLabel(i)
        End If
    Next i
End Sub

Private Sub EnsurePeopleFolders()
    Dim i As Long, scope As String, sch As Object, sf As Object, flt As String
    On Error Resume Next
    If m_PeopleCount = 0 Then Exit Sub
    scope = "'" & Application.Session.GetDefaultFolder(olFolderInbox).FolderPath & "','" & _
            Application.Session.GetDefaultFolder(olFolderSentMail).FolderPath & "'"
    For i = 1 To m_PeopleCount
        If FindSearchFolder(m_PeopleLabel(i)) Is Nothing Then
            EnsureCategory m_PeopleLabel(i), olCategoryColorGreen
            flt = """urn:schemas-microsoft-com:office:office#Keywords"" = '" & Replace(m_PeopleLabel(i), "'", "''") & "'"
            Set sch = Nothing
            Set sch = Application.AdvancedSearch(scope, flt, True, m_PeopleLabel(i))
            Set sf = Nothing
            Set sf = sch.Save(m_PeopleLabel(i))
            AddToFavorites sf
        End If
    Next i
End Sub

Public Sub TagPeopleNow()
    Dim v As String, days As Long, rows As New Collection, rw As Variant, i As Long, msg As String
    Dim want As Boolean, has As Boolean, itm As Object, nAdd As Long, nDel As Long, cnt() As Long
    EnsureInit
    LoadPeople
    If m_PeopleCount = 0 Then
        MsgBoxU U("\uC0AC\uB78C \uBAA9\uB85D\uC774 \uBE44\uC5B4 \uC788\uC2B5\uB2C8\uB2E4. EditPeople \uB85C people.txt \uB97C \uBA3C\uC800 \uC801\uC5B4 \uC8FC\uC138\uC694."), vbExclamation
        Exit Sub
    End If
    v = InputBox("How many days back should I check (Inbox + Sent Items)?" & vbCrLf & "(e.g. 30, 60, 90)", "ReplyAlert", "60")
    If Len(Trim$(v)) = 0 Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    days = CLng(v)
    If days < 1 Then Exit Sub
    If m_Busy Then
        MsgBoxU U("\uC9C0\uAE08 \uB2E4\uB978 \uAC80\uC0AC\uB97C \uD558\uB294 \uC911\uC785\uB2C8\uB2E4. \uC7A0\uC2DC \uD6C4 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694."), vbInformation
        Exit Sub
    End If
    m_Busy = True
    On Error Resume Next
    EnsurePeopleFolders
    BuildPeopleSets days, rows
    ReDim cnt(1 To m_PeopleCount)
    For Each rw In rows
        Set itm = Nothing
        For i = 1 To m_PeopleCount
            want = InThread(i, CStr(rw(2)), CStr(rw(3)))
            has = CatListHas(CStr(rw(1)), m_PeopleLabel(i))
            If want Then cnt(i) = cnt(i) + 1
            If want <> has Then
                If itm Is Nothing Then Set itm = GetItemSafe(CStr(rw(0)))
                If Not itm Is Nothing Then
                    If want Then
                        EnsureCategory m_PeopleLabel(i), olCategoryColorGreen
                        AddCategoryName itm, m_PeopleLabel(i)
                        nAdd = nAdd + 1
                    Else
                        RemoveCategoryName itm, m_PeopleLabel(i), True
                        nDel = nDel + 1
                    End If
                End If
            End If
        Next i
        If (nAdd + nDel) Mod 20 = 1 Then DoEvents
    Next rw
    On Error GoTo 0
    m_Busy = False
    msg = U("\uCD5C\uADFC ") & days & U("\uC77C, \uC8FC\uACE0\uBC1B\uC740 \uC2A4\uB808\uB4DC \uAE30\uC900\uC73C\uB85C \uC815\uB9AC\uD588\uC2B5\uB2C8\uB2E4.") & vbCrLf & _
          U("(\uC0C8\uB85C \uD45C\uC2DC ") & nAdd & U("\uAC74, \uD45C\uC2DC \uD574\uC81C ") & nDel & U("\uAC74)") & vbCrLf
    For i = 1 To m_PeopleCount
        msg = msg & vbCrLf & "  - " & m_PeopleLabel(i) & ": " & cnt(i) & U("\uAC74")
    Next i
    msg = msg & vbCrLf & vbCrLf & U("(\uC990\uACA8\uCC3E\uAE30 \uB610\uB294 \uD3F4\uB354 \uBAA9\uB85D > \uAC80\uC0C9 \uD3F4\uB354 \uC5D0\uC11C \uBCF4\uC138\uC694)")
    MsgBoxU msg, vbInformation
End Sub

Public Sub ReplyReport()
    Dim v As String, days As Long, cutoff As Date
    Dim convs As Object, myLast As Object, myLastId As Object, myFirst As Object, mySubj As Object, myTo As Object
    Dim cnt As Object, lastT As Object, lastWho As Object, lastId As Object, whoSet As Object
    Dim folders As New Collection, fld As Object, inbox As Object
    Dim tbl As Object, arr As Variant, r As Long, t As Date, conv As String, k As Variant
    Dim C_EID As Long, C_CLASS As Long, C_TIME As Long, C_SUBJ As Long, C_TO As Long, C_CONV As Long, C_SND As Long, C_SNDNAME As Long
    Dim rowsNew As String, rowsWait As String, rowsNone As String, nNew As Long, nWait As Long, nNone As Long
    Dim html As String, path As String, row As String, status As String

    EnsureInit
    v = InputBox("How many days back should I check the mail I sent?" & vbCrLf & "(e.g. 7, 14, 30)", "ReplyAlert", "14")
    If Len(Trim$(v)) = 0 Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    days = CLng(v)
    If days < 1 Then Exit Sub
    cutoff = Now - days
    EnsureDicts

    Set convs = NewDict(): Set myLast = NewDict(): Set myLastId = NewDict(): Set myFirst = NewDict()
    Set mySubj = NewDict(): Set myTo = NewDict()
    Set cnt = NewDict(): Set lastT = NewDict(): Set lastWho = NewDict(): Set lastId = NewDict(): Set whoSet = NewDict()

    On Error Resume Next
    ' 1) ->
    Set tbl = Application.Session.GetDefaultFolder(olFolderSentMail).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_TIME = AddCol(tbl, "SentOn")
    C_SUBJ = AddCol(tbl, "Subject")
    C_TO = AddCol(tbl, PR_DISPLAY_TO)
    C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
    tbl.Sort "[SentOn]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_TIME))
            If t <> 0 And t < cutoff Then GoTo SentDone
            If Left$(RowClass(arr, r, C_CLASS), 8) = "ipm.note" Then
                conv = VHex(AV(arr, r, C_CONV))
                If Len(conv) = 0 Then conv = "eid:" & VStr(AV(arr, r, C_EID))
                If Not convs.Exists(conv) Then
                    convs(conv) = True
                    myLast(conv) = t
                    myLastId(conv) = VStr(AV(arr, r, C_EID))
                    mySubj(conv) = VStr(AV(arr, r, C_SUBJ))
                    myTo(conv) = VStr(AV(arr, r, C_TO))
                End If
                myFirst(conv) = t
            End If
        Next r
        DoEvents
    Loop
SentDone:
    If convs.Count = 0 Then
        MsgBoxU U("\uCD5C\uADFC ") & days & U("\uC77C \uB3D9\uC548 \uBCF4\uB0B8 \uBA54\uC77C\uC774 \uC5C6\uC2B5\uB2C8\uB2E4."), vbInformation
        Exit Sub
    End If

    ' 2) ( ) +
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    folders.Add inbox
    AddSubFolders inbox, folders, 0
    folders.Add Application.Session.GetDefaultFolder(olFolderDeletedItems)
    For Each fld In folders
        Set tbl = Nothing
        Set tbl = fld.GetTable
        If Not tbl Is Nothing Then
            tbl.Columns.RemoveAll
            C_EID = AddCol(tbl, "EntryID")
            C_CLASS = AddCol(tbl, "MessageClass")
            C_TIME = AddCol(tbl, "ReceivedTime")
            C_SNDNAME = AddCol(tbl, "SenderName")
            C_SND = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
            C_CONV = AddCol(tbl, PR_CONVERSATION_ID)
            tbl.Sort "[ReceivedTime]", True
            Do Until tbl.EndOfTable
                arr = tbl.GetArray(TABLE_CHUNK)
                If Not IsArray(arr) Then Exit Do
                For r = LBound(arr, 1) To UBound(arr, 1)
                    t = VDate(AV(arr, r, C_TIME))
                    If t <> 0 And t < cutoff Then Exit Do
                    conv = VHex(AV(arr, r, C_CONV))
                    If Len(conv) > 0 Then
                        If convs.Exists(conv) Then
                            If t > myFirst(conv) And Left$(RowClass(arr, r, C_CLASS), 14) <> "ipm.note.rules" _
                               And Not IsMeAddress(VStr(AV(arr, r, C_SND))) Then
                                cnt(conv) = cnt(conv) + 1
                                If Not whoSet.Exists(conv & "|" & VStr(AV(arr, r, C_SNDNAME))) Then
                                    whoSet(conv & "|" & VStr(AV(arr, r, C_SNDNAME))) = True
                                End If
                                If Not lastT.Exists(conv) Then
                                    lastT(conv) = t
                                    lastWho(conv) = VStr(AV(arr, r, C_SNDNAME))
                                    lastId(conv) = VStr(AV(arr, r, C_EID))
                                ElseIf t > lastT(conv) Then
                                    lastT(conv) = t
                                    lastWho(conv) = VStr(AV(arr, r, C_SNDNAME))
                                    lastId(conv) = VStr(AV(arr, r, C_EID))
                                End If
                            End If
                        End If
                    End If
                Next r
                DoEvents
            Loop
        End If
    Next fld
    On Error GoTo 0

    ' 3)
    For Each k In convs.Keys
        conv = CStr(k)
        If lastT.Exists(conv) Then
            row = "<tr><td>" & ReportLink(lastId(conv), mySubj(conv)) & "</td><td>" & HtmlEsc(ShortList(myTo(conv))) & "</td>" & _
                  "<td>" & Format$(myLast(conv), "mm-dd hh:nn") & "</td><td class=n>" & cnt(conv) & "</td>" & _
                  "<td>" & HtmlEsc(lastWho(conv)) & "</td><td>" & Format$(lastT(conv), "mm-dd hh:nn") & "</td></tr>"
            If lastT(conv) > myLast(conv) Then
                rowsNew = rowsNew & row
                nNew = nNew + 1
            Else
                rowsWait = rowsWait & row
                nWait = nWait + 1
            End If
        Else
            row = "<tr><td>" & ReportLink(myLastId(conv), mySubj(conv)) & "</td><td>" & HtmlEsc(ShortList(myTo(conv))) & "</td>" & _
                  "<td>" & Format$(myLast(conv), "mm-dd hh:nn") & "</td><td class=n>0</td><td></td><td></td></tr>"
            rowsNone = rowsNone & row
            nNone = nNone + 1
        End If
    Next k

    html = "<!doctype html><html><head><meta charset=""utf-8""><title>" & U("\uB0B4 \uBA54\uC77C \uD68C\uC2E0 \uD604\uD669") & "</title><style>" & _
           "body{font-family:'Malgun Gothic','Segoe UI',sans-serif;margin:24px;color:#222}" & _
           "h1{font-size:20px}h2{font-size:16px;margin-top:28px}" & _
           "table{border-collapse:collapse;width:100%;font-size:13px}" & _
           "th,td{border-bottom:1px solid #ddd;padding:6px 8px;text-align:left;vertical-align:top}" & _
           "th{background:#f3f3f3}td.n{text-align:right}a{color:#0b57d0;text-decoration:none}" & _
           ".new{color:#c0392b}.wait{color:#2463b5}.none{color:#d35400}.note{color:#666;font-size:12px}" & _
           "</style></head><body>" & _
           "<h1>" & U("\uB0B4 \uBA54\uC77C \uD68C\uC2E0 \uD604\uD669") & " (" & U("\uCD5C\uADFC ") & days & U("\uC77C, ") & Format$(Now, "yyyy-mm-dd hh:nn") & ")</h1>" & _
           "<p class=note>" & U("\uB0B4\uAC00 \uBCF4\uB0B8 \uBA54\uC77C\uC744 \uC2A4\uB808\uB4DC(\uB300\uD654)\uBCC4\uB85C \uBB36\uC5C8\uC2B5\uB2C8\uB2E4. \uC81C\uBAA9\uC744 \uB204\uB974\uBA74 Outlook \uC5D0\uC11C \uBA54\uC77C\uC774 \uC5F4\uB9BD\uB2C8\uB2E4.") & "</p>"
    html = html & ReportSection("new", U("\uD68C\uC2E0\uC774 \uC654\uACE0 \uC544\uC9C1 \uB0B4\uAC00 \uB2E4\uC2DC \uB2F5\uD558\uC9C0 \uC54A\uC74C"), nNew, rowsNew, U("\uB9C8\uC9C0\uB9C9 \uD68C\uC2E0"))
    html = html & ReportSection("wait", U("\uD68C\uC2E0\uC744 \uBC1B\uC740 \uB4A4 \uB0B4\uAC00 \uB2E4\uC2DC \uBCF4\uB0C4 (\uC0C1\uB300 \uB2F5 \uAE30\uB2E4\uB9AC\uB294 \uC911)"), nWait, rowsWait, U("\uB9C8\uC9C0\uB9C9 \uD68C\uC2E0"))
    html = html & ReportSection("none", U("\uC544\uC9C1 \uD68C\uC2E0 \uC5C6\uC74C"), nNone, rowsNone, "")
    html = html & "</body></html>"

    path = Environ$("APPDATA") & "\ReplyAlert\ReplyReport.html"
    WriteUtf8 path, html
    CreateObject("Shell.Application").Open path
End Sub

Private Function ReportSection(ByVal cls As String, ByVal title As String, ByVal n As Long, ByVal rows As String, ByVal lastCol As String) As String
    ReportSection = "<h2 class=" & cls & ">" & HtmlEsc(title) & " - " & n & "</h2>"
    If n = 0 Then Exit Function
    ReportSection = ReportSection & "<table><tr><th>" & U("\uC81C\uBAA9") & "</th><th>" & U("\uBC1B\uB294 \uC0AC\uB78C") & "</th><th>" & U("\uB0B4\uAC00 \uB9C8\uC9C0\uB9C9 \uBCF4\uB0B8 \uC2DC\uAC01") & "</th>" & _
                    "<th>" & U("\uD68C\uC2E0 \uC218") & "</th><th>" & IIf(Len(lastCol) > 0, U("\uB9C8\uC9C0\uB9C9 \uD68C\uC2E0\uD55C \uC0AC\uB78C"), "") & "</th><th>" & IIf(Len(lastCol) > 0, U("\uB9C8\uC9C0\uB9C9 \uD68C\uC2E0 \uC2DC\uAC01"), "") & "</th></tr>" & _
                    rows & "</table>"
End Function

Private Function ReportLink(ByVal eid As String, ByVal subj As String) As String
    If Len(subj) = 0 Then subj = "(" & U("\uC81C\uBAA9 \uC5C6\uC74C") & ")"
    If Len(eid) = 0 Then
        ReportLink = HtmlEsc(subj)
    Else
        ReportLink = "<a href=""" & URI_SCHEME & ":open/" & eid & """>" & HtmlEsc(subj) & "</a>"
    End If
End Function

' "A; B; C; D" -> "A; B 2"
Private Function ShortList(ByVal s As String) As String
    Dim parts() As String
    If Len(Trim$(s)) = 0 Then Exit Function
    parts = Split(s, ";")
    If UBound(parts) <= 1 Then
        ShortList = s
    Else
        ShortList = Trim$(parts(0)) & "; " & Trim$(parts(1)) & " " & U("\uC678 ") & (UBound(parts) - 1) & U("\uBA85")
    End If
End Function

Private Function HtmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEsc = s
End Function

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

' "\uAE40\uB098" ( )
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
        Case URGENT_LIST_NAME: FavKey = "urgent"
        Case Else: FavKey = "other"
    End Select
End Function

' ( Windows ) /
Private Sub CleanupGarbled()
    Dim targets As Object, garbled As New Collection, eids As New Collection, delFolders As New Collection
    Dim cat As Object, f As Object, i As Long, eid As Variant, nm As Variant, itm As Object, n As Long
    On Error Resume Next
    Set targets = GarbledTargets()

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

Private Function GarbledTargets() As Object
    Dim t As Object
    Set t = NewDict()
    t("B3BB20B8DEC0CF20C8B8BDC5") = True          ' (CP949)
    t("B3BB20BED5C0B8B7CE20BFC220B8DEC0CF") = True
    t("B3BB20B4E4C0E520C7CABFE4") = True
    t("B4E4C0E520B4EBB1E2") = True
    t("BFCFB7E120C3B3B8AE") = True
    Set GarbledTargets = t
End Function

Public Sub FixBrokenCategories()
    Dim n As Long
    EnsureInit
    If m_Busy Then
        MsgBoxU U("\uC9C0\uAE08 \uB2E4\uB978 \uAC80\uC0AC\uB97C \uD558\uB294 \uC911\uC785\uB2C8\uB2E4. \uC7A0\uC2DC \uD6C4 \uB2E4\uC2DC \uC2E4\uD589\uD558\uC138\uC694."), vbInformation, APP_TITLE
        Exit Sub
    End If
    m_Busy = True
    CleanupGarbled
    n = CleanupGarbledItems()
    m_Busy = False
    MsgBoxU U("\uAE68\uC9C4 \uBC94\uC8FC\uB97C ") & n & U("\uAC74\uC758 \uBA54\uC77C\uC5D0\uC11C \uC9C0\uC6E0\uC2B5\uB2C8\uB2E4."), vbInformation, APP_TITLE
End Sub

Private Function CleanupGarbledItems() As Long
    Dim folders As New Collection, fld As Object, inbox As Object, targets As Object, n As Long, cat As Object
    Dim garbled As New Collection, nm As Variant
    On Error Resume Next
    Set targets = GarbledTargets()
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    folders.Add inbox
    AddSubFolders inbox, folders, 0
    folders.Add Application.Session.GetDefaultFolder(olFolderSentMail)
    folders.Add Application.Session.GetDefaultFolder(olFolderDeletedItems)
    For Each fld In folders
        n = n + CleanFolderGarbled(fld, targets, Now - 30)
    Next fld
    For Each cat In Application.Session.Categories
        If IsGarbled(cat.Name, targets) Then garbled.Add cat.Name
    Next cat
    For Each nm In garbled
        Application.Session.Categories.Remove CStr(nm)
    Next nm
    CleanupGarbledItems = n
End Function

Private Function CleanFolderGarbled(ByVal fld As Object, ByVal targets As Object, ByVal cutoff As Date) As Long
    Dim tbl As Object, arr As Variant, r As Long, t As Date, cats As String, parts() As String, i As Long
    Dim result As String, changed As Boolean, itm As Object, n As Long
    Dim C_EID As Long, C_MOD As Long, C_CATS As Long
    On Error GoTo Done
    Set tbl = fld.GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_MOD = AddCol(tbl, "LastModificationTime")
    C_CATS = AddCol(tbl, "Categories")
    If C_EID < 0 Or C_MOD < 0 Or C_CATS < 0 Then GoTo Done
    tbl.Sort "[LastModificationTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_MOD))
            If t <> 0 And t < cutoff Then GoTo Done
            cats = VStr(AV(arr, r, C_CATS))
            If Len(cats) > 0 Then
                parts = Split(Replace(cats, ";", ","), ",")
                result = ""
                changed = False
                For i = LBound(parts) To UBound(parts)
                    If IsGarbled(Trim$(parts(i)), targets) Then
                        changed = True
                    ElseIf Len(Trim$(parts(i))) > 0 Then
                        If Len(result) > 0 Then result = result & ", "
                        result = result & Trim$(parts(i))
                    End If
                Next i
                If changed Then
                    Set itm = GetItemSafe(VStr(AV(arr, r, C_EID)))
                    If Not itm Is Nothing Then
                        itm.Categories = result
                        itm.Save
                        n = n + 1
                    End If
                End If
            End If
        Next r
        DoEvents
    Loop
Done:
    CleanFolderGarbled = n
End Function

' () : PC (CP949)
Private Function IsGarbled(ByVal nm As String, ByVal targets As Object) As Boolean
    Dim b() As Byte, i As Long, h As String
    On Error GoTo Done
    If Len(nm) = 0 Then Exit Function
    For i = 1 To Len(nm)
        If AscW(Mid$(nm, i, 1)) > 255 Or AscW(Mid$(nm, i, 1)) < 0 Then Exit Function
    Next i
    b = StrConv(nm, vbFromUnicode)
    For i = LBound(b) To UBound(b)
        h = h & Right$("0" & Hex$(b(i)), 2)
    Next i
    IsGarbled = targets.Exists(h)
Done:
End Function


'==========================================================================
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

' MarkDone
Private Function IsDone(ByVal itm As Object) As Boolean
    IsDone = CatListHas(SafeCategories(itm), DONE_CATEGORY)
End Function


'==========================================================================
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
    If Left$(subj, 5) = U("\uC790\uB3D9 \uD68C\uC2E0") Then IsAutoReply = True
    If Left$(subj, 4) = U("\uC790\uB3D9\uD68C\uC2E0") Then IsAutoReply = True
    If Left$(subj, 3) = U("\uBD80\uC7AC\uC911") Then IsAutoReply = True
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

Private Function ItemTime(ByVal itm As Object) As Date
    ItemTime = GetReceived(itm)
    If ItemTime = 0 Then ItemTime = SafeSentOn(itm)
End Function

Private Function SafeEntryId(ByVal itm As Object) As String
    On Error Resume Next
    SafeEntryId = itm.EntryID
End Function

' () ID ( 16)
Private Function SafeConvId(ByVal itm As Object) As String
    On Error Resume Next
    SafeConvId = UCase$(itm.ConversationID)
End Function

' (Table) . -1 ( )
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

Private Function RowClass(ByRef arr As Variant, ByVal r As Long, ByVal c As Long) As String
    If c < 0 Then
        RowClass = "ipm.note"
    Else
        RowClass = LCase$(VStr(arr(r, c)))
    End If
End Function

' (Table)
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

' ( ) -> 16
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

' ID ( ID, , <> )
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

' PowerShell -EncodedCommand (UTF-16LE -> Base64)
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
    If v Then YesNo = U("\uC608") Else YesNo = U("\uC544\uB2C8\uC624")
End Function
