Attribute VB_Name = "ReplyAlert"
'==========================================================================
'  내 메일 회신 알림 (Outlook VBA)
'
'  - 받은 메일이 "내가 보낸 메일에 대한 회신"이면
'      1) 빨간색 범주 [내 메일 회신] 을 붙여 목록에서 눈에 띄게 표시
'      2) Windows 알림(화면 오른쪽 아래)으로 보낸 사람 / 제목 표시
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
'    ScanRecentInbox   : 최근 받은 메일을 검사해서 회신 표시 (기본 7일)
'    TestSelectedMail  : 선택한 메일이 왜 회신으로 판단됐는지/안 됐는지 진단
'    TestToast         : Windows 알림이 뜨는지 테스트
'    ClearReplyMark    : 선택한 메일에서 [내 메일 회신] 표시 제거
'==========================================================================
Option Explicit

'------------------------------ 설정 ------------------------------
Public Const CATEGORY_NAME As String = "내 메일 회신"   ' 표시용 범주 이름 (빨간색으로 자동 생성)

Private Const SHOW_TOAST As Boolean = True          ' Windows 알림 표시
Private Const SHOW_POPUP As Boolean = False         ' 메시지 박스 표시 (확인 누를 때까지 Outlook 멈춤)
Private Const SET_FLAG As Boolean = False           ' 플래그(오늘) 추가 -> 할 일 목록에도 표시
Private Const SET_HIGH_IMPORTANCE As Boolean = False ' 중요도 "높음" 으로 변경

' False : 내 메일에 "직접" 답장한 경우만 (권장)
' True  : 스레드 안에 내 메일이 있으면, 다른 사람 메일에 단 답장이어도 표시
Private Const MATCH_WHOLE_THREAD As Boolean = False

Private Const USE_CONVERSATION_FALLBACK As Boolean = True ' 대화(스레드) 기반 보조 판단
Private Const SKIP_AUTO_REPLIES As Boolean = True   ' 부재중 자동회신 등은 제외

Private Const SCAN_DAYS As Long = 7                 ' ScanRecentInbox 검사 기간(일)
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

Private m_MyAddresses As String   ' ";a@x.com;b@y.com;" 형태 캐시


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
    If STARTUP_SCAN_DAYS <= 0 Then Exit Sub
    On Error Resume Next
    n = ScanInbox(STARTUP_SCAN_DAYS)
    If n > 0 And SHOW_TOAST Then
        ShowToast "내 메일에 회신 " & n & "건", _
                  "Outlook 이 꺼져 있는 동안 도착한 회신입니다.", _
                  "[" & CATEGORY_NAME & "] 범주로 표시했습니다."
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
End Sub

Public Sub TestToast()
    ShowToast "내 메일에 회신이 왔습니다", "홍길동 (테스트)", "RE: 알림 테스트 메일"
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
    Dim sender As String, subj As String
    On Error Resume Next
    sender = mail.SenderName
    subj = mail.Subject
    If SHOW_TOAST Then ShowToast "내 메일에 회신이 왔습니다", sender, subj
    If SHOW_POPUP Then
        MsgBox "내 메일에 회신이 왔습니다." & vbCrLf & vbCrLf & _
               "보낸 사람: " & sender & vbCrLf & "제목: " & subj, _
               vbInformation + vbSystemModal, "회신 알림"
    End If
End Sub

' Windows 10/11 알림 (PowerShell 사용, 별도 설치 불필요)
Private Sub ShowToast(ByVal title As String, ByVal line1 As String, ByVal line2 As String)
    Dim xml As String, ps As String
    On Error Resume Next
    xml = "<toast duration=""long""><visual><binding template=""ToastGeneric"">" & _
          "<text>" & XmlEsc(title) & "</text>" & _
          "<text>" & XmlEsc(line1) & "</text>" & _
          "<text>" & XmlEsc(line2) & "</text>" & _
          "</binding></visual><audio src=""ms-winsoundevent:Notification.Mail""/></toast>"
    ps = "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null;" & _
         "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null;" & _
         "$x = New-Object Windows.Data.Xml.Dom.XmlDocument;" & _
         "$x.LoadXml('" & xml & "');" & _
         "$t = New-Object Windows.UI.Notifications.ToastNotification $x;" & _
         "[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('" & TOAST_APP_ID & "').Show($t)"
    Shell "powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " & B64Utf16(ps), vbHide
End Sub


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
