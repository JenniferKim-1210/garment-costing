Attribute VB_Name = "ReplyAlert"
'==========================================================================
'  내 메일 회신 알림 (Outlook VBA)
'
'  - 받은 메일이 "내가 보낸 메일에 대한 회신"이면
'      1) 빨간색 범주 [내 메일 회신] 을 붙여 목록에서 눈에 띄게 표시
'      2) Windows 알림(화면 오른쪽 아래)으로 보낸 사람 / 제목 표시
'      3) 알림을 클릭하면 해당 메일이 열림, [회신 목록 보기] 버튼은 목록 폴더 열기
'      4) 회신만 모아 보는 검색 폴더 [내 메일 회신] 을 자동 생성 + 즐겨찾기에 추가
'  - 판단 방법
'      1) 받은 메일의 In-Reply-To 헤더(어떤 메일에 대한 답장인지)를
'         보낸 편지함의 메일 ID 와 비교
'      2) 보낸 편지함에 없으면(다른 폴더로 옮긴 경우 등) 대화(스레드)에서 찾기
'      3) 헤더가 아예 없으면 같은 대화에 내가 먼저 보낸 메일이 있는지 확인
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
'    ShowReplyList     : 회신 목록(검색 폴더) 열기
'    ScanRecentInbox   : 최근 받은 메일을 검사해서 회신 표시 (기본 14일)
'    TestSelectedMail  : 선택한 메일이 왜 회신으로 판단됐는지/안 됐는지 진단
'    TestToast         : Windows 알림 테스트 (메일을 선택하고 실행하면 클릭 시 그 메일이 열림)
'    ClearReplyMark    : 선택한 메일에서 [내 메일 회신] 표시 제거
'==========================================================================
Option Explicit

'------------------------------ 설정 ------------------------------
Public Const CATEGORY_NAME As String = "내 메일 회신"   ' 표시용 범주 이름 (빨간색으로 자동 생성)
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

Private Const SCAN_DAYS As Long = 14               ' ScanRecentInbox 검사 기간(일)
Private Const STARTUP_SCAN_DAYS As Long = 1         ' Outlook 시작 시 검사 기간(일), 0 = 안 함

' 내 주소가 여러 개(별칭)면 세미콜론으로 추가: "a@company.com;b@company.com"
Private Const EXTRA_MY_ADDRESSES As String = ""
'------------------------------------------------------------------

Private Const PR_IN_REPLY_TO_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"
Private Const PR_INTERNET_REFERENCES As String = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
Private Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"

Private Const TOAST_APP_ID As String = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe"
Private Const URI_SCHEME As String = "replyalert"   ' 알림 클릭용 주소 (replyalert:open/<EntryID>, replyalert:list)

Private m_MyAddresses As String   ' ";a@x.com;b@y.com;" 형태 캐시
Private m_HandlerReady As Boolean ' 클릭 처리기 설치 여부


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
End Sub

Public Sub OnStartup()
    Dim n As Long
    On Error Resume Next
    If CLICK_TO_OPEN Then InstallClickHandler
    If CREATE_REPLY_LIST Then EnsureReplyList
    If STARTUP_SCAN_DAYS <= 0 Then Exit Sub
    n = ScanInbox(STARTUP_SCAN_DAYS)
    If n > 0 And SHOW_TOAST Then
        ShowToast "내 메일에 회신 " & n & "건", _
                  "Outlook 이 꺼져 있는 동안 도착한 회신입니다.", _
                  "클릭하면 회신 목록을 엽니다.", _
                  URI_SCHEME & ":list", ""
    End If
End Sub


'==========================================================================
'  수동 실행 매크로
'==========================================================================
Public Sub ScanRecentInbox()
    Dim n As Long
    n = ScanInbox(SCAN_DAYS)
    MsgBox "최근 " & SCAN_DAYS & "일 받은 메일 중 내 메일에 대한 회신 " & n & "건을 새로 표시했습니다." & vbCrLf & _
           "(이미 표시된 메일은 제외)", vbInformation, "회신 알림"
    If CREATE_REPLY_LIST Then ShowReplyList
End Sub

Public Sub ShowReplyList()
    Dim sf As Object
    On Error Resume Next
    Set sf = FindReplyList()
    If sf Is Nothing Then
        EnsureReplyList
        Set sf = FindReplyList()
    End If
    If sf Is Nothing Then
        MsgBox "회신 목록 폴더를 찾지 못했습니다.", vbExclamation, "회신 알림"
        Exit Sub
    End If
    If Application.ActiveExplorer Is Nothing Then
        sf.Display
    Else
        Set Application.ActiveExplorer.CurrentFolder = sf
    End If
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
        ShowToast "내 메일에 회신이 왔습니다", "홍길동 (테스트)", "RE: 알림 테스트 메일", _
                  URI_SCHEME & ":list", ""
    Else
        ShowToast "내 메일에 회신이 왔습니다", "(테스트) 클릭하면 선택한 메일이 열립니다", subj, _
                  openUri, openUri
    End If
End Sub

Public Sub ClearReplyMark()
    Dim sel As Object, itm As Object
    On Error Resume Next
    Set sel = Application.ActiveExplorer.Selection
    For Each itm In sel
        RemoveCategory itm
    Next itm
End Sub

Public Sub TestSelectedMail()
    Dim itm As Object, pa As Object, msg As String
    Dim inReplyTo As String, refs As String, found As String
    Dim c As Collection, i As Long

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

    Set c = SplitIds(inReplyTo)
    For i = 1 To c.Count
        If SentHasMessageId(c(i)) Then found = found & c(i) & " "
    Next i

    msg = "제목: " & itm.Subject & vbCrLf & _
          "보낸 사람 주소: " & GetSenderSmtp(itm) & vbCrLf & _
          "내 주소 목록: " & MyAddresses() & vbCrLf & vbCrLf & _
          "내가 보낸 메일인가: " & YesNo(IsFromMe(itm)) & vbCrLf & _
          "자동회신인가: " & YesNo(IsAutoReply(itm)) & vbCrLf & vbCrLf & _
          "In-Reply-To: " & IIf(Len(inReplyTo) = 0, "(없음)", inReplyTo) & vbCrLf & _
          "References 개수: " & SplitIds(refs).Count & vbCrLf & _
          "보낸 편지함에서 찾음: " & IIf(Len(found) = 0, "(없음)", found) & vbCrLf & vbCrLf & _
          "==> 최종 판단: " & IIf(IsReplyToMe(itm), "내 메일에 대한 회신 O", "회신 아님 X")
    MsgBox msg, vbInformation, "회신 알림 - 진단"
End Sub


'==========================================================================
'  핵심 로직
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
    On Error GoTo Fail
    If itm.Class <> olMail Then Exit Function
    If HasCategory(itm) Then Exit Function
    If SKIP_AUTO_REPLIES Then
        If IsAutoReply(itm) Then Exit Function
    End If
    If IsFromMe(itm) Then Exit Function
    If Not IsReplyToMe(itm) Then Exit Function

    MarkMail itm
    If notify Then NotifyReply itm
    ProcessMail = True
    Exit Function
Fail:
    Debug.Print "ReplyAlert.ProcessMail 오류: " & Err.Description
End Function

Private Function IsReplyToMe(ByVal mail As Object) As Boolean
    Dim pa As Object, inReplyTo As String, refs As String
    Dim c As Collection, i As Long, idSet As String

    Set pa = mail.PropertyAccessor
    inReplyTo = GetPropStr(pa, PR_IN_REPLY_TO_ID)
    If MATCH_WHOLE_THREAD Then refs = GetPropStr(pa, PR_INTERNET_REFERENCES)

    ' 1) 답장 대상 메일 ID 가 보낸 편지함에 있는지
    Set c = SplitIds(inReplyTo & " " & refs)
    For i = 1 To c.Count
        If SentHasMessageId(c(i)) Then
            IsReplyToMe = True
            Exit Function
        End If
        idSet = idSet & "|" & LCase$(c(i))
    Next i

    If Not USE_CONVERSATION_FALLBACK Then Exit Function

    If c.Count > 0 Then
        ' 2) 보낸 편지함에는 없지만(다른 폴더로 옮김 등) 대화 안에 그 메일이 있고 내가 보낸 것인지
        IsReplyToMe = ConversationHasMatch(mail, idSet & "|")
    Else
        ' 3) 헤더가 없으면: 같은 대화에 내가 먼저 보낸 메일이 있는지
        IsReplyToMe = ConversationHasMatch(mail, "")
    End If
End Function

Private Function SentHasMessageId(ByVal msgId As String) As Boolean
    Dim f As Object, res As Object, flt As String
    On Error GoTo Done
    Set f = Application.Session.GetDefaultFolder(olFolderSentMail)
    flt = "@SQL=""" & PR_INTERNET_MESSAGE_ID & """ = '" & Replace(msgId, "'", "''") & "'"
    Set res = f.Items.Restrict(flt)
    SentHasMessageId = (res.Count > 0)
Done:
End Function

' idSet = "|<id1>|<id2>|" (소문자) 이면 해당 ID 의 메일이 내가 보낸 것인지 확인,
' idSet = "" 이면 이 메일보다 먼저 내가 보낸 메일이 대화에 있는지 확인
Private Function ConversationHasMatch(ByVal mail As Object, ByVal idSet As String) As Boolean
    Dim conv As Object, roots As Object, r As Object
    On Error GoTo Done
    Set conv = mail.GetConversation
    If conv Is Nothing Then Exit Function
    Set roots = conv.GetRootItems
    For Each r In roots
        If WalkConversation(conv, r, mail.EntryID, mail.ReceivedTime, idSet, 0) Then
            ConversationHasMatch = True
            Exit Function
        End If
    Next r
Done:
End Function

Private Function WalkConversation(ByVal conv As Object, ByVal itm As Object, ByVal selfId As String, _
                                  ByVal before As Date, ByVal idSet As String, ByVal depth As Long) As Boolean
    Dim kids As Object, k As Object
    On Error GoTo Done
    If depth > 40 Then Exit Function
    If IsConversationMatch(itm, selfId, before, idSet) Then
        WalkConversation = True
        Exit Function
    End If
    Set kids = conv.GetChildren(itm)
    For Each k In kids
        If WalkConversation(conv, k, selfId, before, idSet, depth + 1) Then
            WalkConversation = True
            Exit Function
        End If
    Next k
Done:
End Function

Private Function IsConversationMatch(ByVal itm As Object, ByVal selfId As String, _
                                     ByVal before As Date, ByVal idSet As String) As Boolean
    Dim msgIdLc As String
    On Error GoTo Done
    If itm.Class <> olMail Then Exit Function
    If itm.EntryID = selfId Then Exit Function
    If Len(idSet) > 0 Then
        msgIdLc = LCase$(GetPropStr(itm.PropertyAccessor, PR_INTERNET_MESSAGE_ID))
        If Len(msgIdLc) = 0 Then Exit Function
        If InStr(1, idSet, "|" & msgIdLc & "|", vbBinaryCompare) = 0 Then Exit Function
    Else
        If itm.SentOn >= before Then Exit Function
    End If
    IsConversationMatch = IsFromMe(itm)
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
    EnsureCategory
    If Len(mail.Categories) = 0 Then
        mail.Categories = CATEGORY_NAME
    Else
        mail.Categories = mail.Categories & ", " & CATEGORY_NAME
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
'   launchUri : 알림 본문을 클릭했을 때 열 주소 (replyalert:open/<EntryID> 또는 replyalert:list)
'   openUri   : [메일 열기] 버튼 주소 (빈 문자열이면 버튼 없음)
Private Sub ShowToast(ByVal title As String, ByVal line1 As String, ByVal line2 As String, _
                      Optional ByVal launchUri As String = "", Optional ByVal openUri As String = "")
    Dim xml As String, ps As String, head As String, actions As String
    On Error Resume Next

    If CLICK_TO_OPEN And Len(launchUri) > 0 Then
        If Not m_HandlerReady Then InstallClickHandler
    End If

    If CLICK_TO_OPEN And m_HandlerReady And Len(launchUri) > 0 Then
        head = "<toast activationType=""protocol"" launch=""" & XmlEsc(launchUri) & """ duration=""long"">"
        actions = "<actions>"
        If Len(openUri) > 0 Then
            actions = actions & "<action content=""메일 열기"" activationType=""protocol"" arguments=""" & XmlEsc(openUri) & """/>"
        End If
        actions = actions & "<action content=""회신 목록 보기"" activationType=""protocol"" arguments=""" & URI_SCHEME & ":list""/>"
        actions = actions & "</actions>"
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
'  회신 목록 (검색 폴더)
'==========================================================================
Private Function FindReplyList() As Object
    Dim f As Object, nm As String
    On Error Resume Next
    For Each f In Application.Session.GetDefaultFolder(olFolderInbox).Store.GetSearchFolders
        nm = ""
        nm = f.Name
        If nm = LIST_FOLDER_NAME Then
            Set FindReplyList = f
            Exit Function
        End If
    Next f
End Function

' 받은 편지함(하위 폴더 포함)에서 [내 메일 회신] 범주가 붙은 메일만 모아 보는 검색 폴더
Private Sub EnsureReplyList()
    Dim inbox As Object, sch As Object, sf As Object, flt As String
    On Error GoTo Done
    If Not FindReplyList() Is Nothing Then Exit Sub
    Set inbox = Application.Session.GetDefaultFolder(olFolderInbox)
    flt = """urn:schemas-microsoft-com:office:office#Keywords"" = '" & Replace(CATEGORY_NAME, "'", "''") & "'"
    Set sch = Application.AdvancedSearch("'" & inbox.FolderPath & "'", flt, True, LIST_FOLDER_NAME)
    Set sf = sch.Save(LIST_FOLDER_NAME)
    AddToFavorites sf      ' 처음 만들 때만 즐겨찾기에 추가 (지워도 다시 추가하지 않음)
Done:
End Sub

Private Sub AddToFavorites(ByVal fld As Object)
    Dim grp As Object, nf As Object
    On Error GoTo Done
    If fld Is Nothing Then Exit Sub
    If Application.ActiveExplorer Is Nothing Then Exit Sub
    Set grp = Application.ActiveExplorer.NavigationPane.Modules _
                .GetNavigationModule(olModuleMail).NavigationGroups _
                .GetDefaultNavigationGroup(olFavoriteFoldersGroup)
    For Each nf In grp.NavigationFolders
        If nf.Folder.EntryID = fld.EntryID Then Exit Sub
    Next nf
    grp.NavigationFolders.Add fld
Done:
End Sub


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
        "Dim uri, arg, ol, ns, itm, f, ex, sh, re, found" & vbCrLf & _
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
        "' 메일을 못 찾았거나(다른 폴더로 옮김 등) 목록 버튼이면 회신 목록 열기" & vbCrLf & _
        "If Not found Then" & vbCrLf & _
        "  For Each f In ns.GetDefaultFolder(6).Store.GetSearchFolders" & vbCrLf & _
        "    If f.Name = ~<<LIST>>~ Then" & vbCrLf & _
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
    HandlerScript = Replace(s, "~", Chr$(34))
End Function


'==========================================================================
'  유틸리티
'==========================================================================
Private Sub EnsureCategory()
    Dim cat As Object
    On Error Resume Next
    Set cat = Application.Session.Categories.Item(CATEGORY_NAME)
    If cat Is Nothing Then
        Err.Clear
        Application.Session.Categories.Add CATEGORY_NAME, olCategoryColorRed
    End If
End Sub

Private Function HasCategory(ByVal itm As Object) As Boolean
    Dim parts() As String, i As Long, cats As String
    On Error Resume Next
    cats = itm.Categories
    If Len(cats) = 0 Then Exit Function
    parts = Split(cats, ",")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), CATEGORY_NAME, vbTextCompare) = 0 Then
            HasCategory = True
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveCategory(ByVal itm As Object)
    Dim parts() As String, i As Long, result As String, p As String
    If Len(itm.Categories) = 0 Then Exit Sub
    parts = Split(itm.Categories, ",")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If Len(p) > 0 And StrComp(p, CATEGORY_NAME, vbTextCompare) <> 0 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & p
        End If
    Next i
    If result <> itm.Categories Then
        itm.Categories = result
        itm.Save
    End If
End Sub

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

Private Function GetReceived(ByVal itm As Object) As Date
    On Error Resume Next
    GetReceived = itm.ReceivedTime
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
