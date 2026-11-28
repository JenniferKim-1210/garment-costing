Attribute VB_Name = "DupRemover"
'==========================================================================
'  Duplicate Mail Remover for Outlook (DupRemover)       - README.md (Korean)
'
'  When the same mail arrives twice (e.g. I am a member of two group addresses
'  such as loungepm@ and sleeppm@), keep the first copy and move the later copy
'  to Deleted Items (it can be restored from there).
'
'  "Same mail" = same Internet Message-ID, or - if that is missing -
'                same sender + subject + sent time (to the second).
'  Only copies that arrive within DUP_WINDOW_MINUTES of each other are removed.
'
'  ThisOutlookSession (see ThisOutlookSession.txt):
'    Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'        DupRemover.CheckNewMail EntryIDCollection
'        ReplyAlert.OnNewMail EntryIDCollection
'    End Sub
'
'  Macros (Alt+F8): RemoveDuplicatesNow (clean duplicates already in the Inbox),
'                   TestDuplicate (why the selected mail is / is not a duplicate)
'
'  Group mail filter: mail that came ONLY through sleeppm@ (not loungepm@, not to me
'  directly) from people listed in sleeppm_senders.txt is moved to a chosen folder.
'  Macros: MoveSleeppmNow, EditSleeppmSenders, SetSleeppmFolder, TestSleeppm
'==========================================================================
Option Explicit

'============================== SETTINGS ==============================
Private Const DUP_ENABLED As Boolean = True
Private Const DUP_WINDOW_MINUTES As Long = 30       ' remove only when the copies arrive within N minutes of each other
Private Const DUP_SEED_DAYS As Long = 1             ' remember mail received in the last N days (for copies arriving after a restart)

Private Const GROUP_MOVE_ENABLED As Boolean = True  ' move sleeppm-only mail from listed senders
Private Const MOVE_GROUP As String = "sleeppm"      ' group address (the part before @ is enough)
Private Const SKIP_GROUP As String = "loungepm"     ' if this group is also a recipient, keep the mail in the Inbox
Private Const MACRO_FOLDER As String = "D:\UserData\Documents\Outlook Macro"   ' sleeppm_senders.txt is copied from here the first time
'======================================================================

Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
Private Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F"
Private Const TABLE_CHUNK As Long = 200
Private Const SETTINGS_APP As String = "ReplyAlert"   ' ReplyAlert reads the same setting (skips the folder)

Private m_Id As Object      ' key -> EntryID of the copy that is kept
Private m_Time As Object    ' key -> received time of that copy
Private m_Seeded As Boolean
Private m_Title As String
Private m_MoveFolderName As String    ' 옮길 폴더 이름 (처음 한 번 이 이름으로 찾고, 찾으면 기억)
Private m_MoveFolder As Object
Private m_Senders As String           ' "|주소|이름|" (소문자)
Private m_SendersStamp As String
Private m_MyAddrs As String

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, _
    ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If


'==========================================================================
'  ThisOutlookSession 의 Application_NewMailEx 에서 호출 (ReplyAlert 보다 먼저)
'==========================================================================
Public Sub CheckNewMail(ByVal EntryIDCollection As String)
    Dim ids() As String, i As Long, itm As Object
    On Error Resume Next
    Init
    If DUP_ENABLED Then EnsureSeeded
    ids = Split(EntryIDCollection, ",")
    For i = LBound(ids) To UBound(ids)
        Set itm = Nothing
        Set itm = Application.Session.GetItemFromID(Trim$(ids(i)))
        If Not itm Is Nothing Then
            If Not CheckItem(itm) Then MoveGroupMail itm
        End If
    Next i
End Sub


'==========================================================================
'  수동 실행 매크로
'==========================================================================
' 받은 편지함에 이미 들어와 있는 중복 메일 정리 (먼저 온 것만 남김)
Public Sub RemoveDuplicatesNow()
    Dim v As String, days As Long, dups As Collection, eid As Variant, itm As Object, n As Long
    Init
    v = InputBox("How many days back should I check the Inbox for duplicates?" & vbCrLf & "(e.g. 3, 7, 30)", "DupRemover", "7")
    If Len(Trim$(v)) = 0 Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    days = CLng(v)
    If days < 1 Then Exit Sub

    Set dups = FindDuplicates(days)
    If dups.Count = 0 Then
        MsgBoxU "최근 " & days & "일 받은 편지함에 중복 메일이 없습니다.", vbInformation
        Exit Sub
    End If
    If MsgBoxU("최근 " & days & "일 받은 편지함에서 중복 메일 " & dups.Count & "건을 찾았습니다." & vbCrLf & vbCrLf & _
               "먼저 온 메일만 남기고, 나머지 " & dups.Count & "건을 지운 편지함으로 옮길까요?" & vbCrLf & _
               "(지운 편지함에서 되살릴 수 있습니다)", vbYesNo + vbQuestion) <> 6 Then Exit Sub   ' 6 = 예
    For Each eid In dups
        Set itm = GetItemSafe(CStr(eid))
        If Not itm Is Nothing Then
            itm.Delete
            n = n + 1
            If n Mod 20 = 0 Then DoEvents
        End If
    Next eid
    MsgBoxU "중복 메일 " & n & "건을 지운 편지함으로 옮겼습니다.", vbInformation
End Sub

' 선택한 메일이 왜 중복인지 / 아닌지 확인
Public Sub TestDuplicate()
    Dim itm As Object, k1 As String, k2 As String, msg As String, other As String
    Init
    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBoxU "메일을 하나 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If
    If itm.Class <> olMail Then
        MsgBoxU "선택한 항목이 일반 메일이 아닙니다.", vbExclamation
        Exit Sub
    End If
    EnsureSeeded
    k1 = IdKey(itm)
    k2 = FallbackKey(itm)
    If Len(k1) > 0 And m_Id.Exists(k1) Then
        If m_Id(k1) <> itm.EntryID Then other = m_Id(k1)
    End If
    If Len(other) = 0 And m_Id.Exists(k2) Then
        If m_Id(k2) <> itm.EntryID Then other = m_Id(k2)
    End If
    msg = "제목: " & itm.Subject & vbCrLf & _
          "받은 시각: " & itm.ReceivedTime & vbCrLf & _
          "Message-ID: " & IIf(Len(k1) = 0, "(없음)", Mid$(k1, 4)) & vbCrLf & vbCrLf
    If Len(other) = 0 Then
        msg = msg & "==> 최근 " & DUP_SEED_DAYS & "일 받은 메일 중 같은 메일이 없습니다."
    Else
        msg = msg & "==> 같은 메일이 하나 더 있습니다. (" & DUP_WINDOW_MINUTES & "분 안에 들어온 경우 나중 것을 지움)"
    End If
    MsgBoxU msg, vbInformation
End Sub


'==========================================================================
'  핵심 로직
'==========================================================================
' 중복이면 나중 것을 지우고 True
Private Function CheckItem(ByVal itm As Object) As Boolean
    Dim k1 As String, k2 As String, rcv As Date, eid As String, key As String
    On Error GoTo Done
    If itm.Class <> olMail Then Exit Function
    If Not DUP_ENABLED Then Exit Function
    eid = itm.EntryID
    rcv = itm.ReceivedTime
    k1 = IdKey(itm)
    k2 = FallbackKey(itm)

    If Len(k1) > 0 And m_Id.Exists(k1) Then
        key = k1
    ElseIf Len(k2) > 0 And m_Id.Exists(k2) Then
        key = k2
    End If

    If Len(key) > 0 Then
        If m_Id(key) <> eid And Abs(DateDiff("n", m_Time(key), rcv)) <= DUP_WINDOW_MINUTES Then
            If KeptCopyExists(m_Id(key)) Then
                itm.Delete                       ' 지운 편지함으로 (되살릴 수 있음)
                CheckItem = True
                Exit Function
            End If
        End If
    End If
    Remember k1, eid, rcv
    Remember k2, eid, rcv
Done:
End Function

' 남겨 둔 메일이 아직 있는지 (사용자가 이미 지웠으면 이번 메일은 남김)
Private Function KeptCopyExists(ByVal eid As String) As Boolean
    Dim itm As Object, delId As String
    On Error Resume Next
    Set itm = GetItemSafe(eid)
    If itm Is Nothing Then Exit Function
    delId = Application.Session.GetDefaultFolder(olFolderDeletedItems).EntryID
    KeptCopyExists = (itm.Parent.EntryID <> delId)
End Function

Private Sub Remember(ByVal key As String, ByVal eid As String, ByVal rcv As Date)
    If Len(key) = 0 Then Exit Sub
    m_Id(key) = eid
    m_Time(key) = rcv
End Sub

Private Function IdKey(ByVal itm As Object) As String
    Dim s As String
    s = LCase$(Trim$(GetPropStr(itm.PropertyAccessor, PR_INTERNET_MESSAGE_ID)))
    If Len(s) > 0 Then IdKey = "id:" & s
End Function

Private Function FallbackKey(ByVal itm As Object) As String
    Dim snd As String, t As Date
    On Error Resume Next
    snd = GetPropStr(itm.PropertyAccessor, PR_SENDER_SMTP_ADDRESS)
    If Len(snd) = 0 Then snd = itm.SenderEmailAddress
    t = itm.SentOn
    If Len(snd) = 0 Or t = 0 Then Exit Function
    FallbackKey = "fb:" & LCase$(snd) & "|" & itm.Subject & "|" & Format$(t, "yyyymmddhhnnss")
End Function

Private Function FallbackKeyOf(ByVal snd As String, ByVal subj As String, ByVal t As Date) As String
    If Len(snd) = 0 Or t = 0 Then Exit Function
    FallbackKeyOf = "fb:" & LCase$(snd) & "|" & subj & "|" & Format$(t, "yyyymmddhhnnss")
End Function

' Outlook 을 다시 켠 뒤에도 바로 앞에 받은 메일과 비교할 수 있도록, 최근 메일을 기억해 둠
Private Sub EnsureSeeded()
    Dim tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date, eid As String
    Dim k1 As String, k2 As String
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_SENT As Long, C_SUBJ As Long, C_SND As Long, C_MSGID As Long
    Init
    If m_Seeded Then Exit Sub
    m_Seeded = True
    On Error GoTo Done
    cutoff = Now - DUP_SEED_DAYS
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_SENT = AddCol(tbl, "SentOn")
    C_SUBJ = AddCol(tbl, "Subject")
    C_SND = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    C_MSGID = AddCol(tbl, PR_INTERNET_MESSAGE_ID)
    If C_EID < 0 Or C_RCV < 0 Then GoTo Done
    tbl.Sort "[ReceivedTime]", True          ' 최신부터 -> 나중에 읽는(더 오래된) 메일이 남길 메일로 덮어씀
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Done
            If Left$(LCase$(VStr(AV(arr, r, C_CLASS))), 8) = "ipm.note" Then
                eid = VStr(AV(arr, r, C_EID))
                k1 = VStr(AV(arr, r, C_MSGID))
                If Len(Trim$(k1)) > 0 Then Remember "id:" & LCase$(Trim$(k1)), eid, t
                k2 = FallbackKeyOf(VStr(AV(arr, r, C_SND)), VStr(AV(arr, r, C_SUBJ)), VDate(AV(arr, r, C_SENT)))
                Remember k2, eid, t
            End If
        Next r
        DoEvents
    Loop
Done:
End Sub

' 받은 편지함 최근 days 일에서 지울 중복 메일(EntryID) 목록. 같은 메일 중 가장 먼저 온 것만 남김
Private Function FindDuplicates(ByVal days As Long) As Collection
    Dim tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date, eid As String
    Dim k1 As String, k2 As String, key As String, keepId As Object, keepTime As Object, dupSet As Object, c As New Collection
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_SENT As Long, C_SUBJ As Long, C_SND As Long, C_MSGID As Long
    Dim k As Variant
    Set FindDuplicates = c
    On Error GoTo Done
    Set keepId = CreateObject("Scripting.Dictionary")
    Set keepTime = CreateObject("Scripting.Dictionary")
    Set dupSet = CreateObject("Scripting.Dictionary")
    cutoff = Now - days
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_SENT = AddCol(tbl, "SentOn")
    C_SUBJ = AddCol(tbl, "Subject")
    C_SND = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    C_MSGID = AddCol(tbl, PR_INTERNET_MESSAGE_ID)
    If C_EID < 0 Or C_RCV < 0 Then GoTo Done
    tbl.Sort "[ReceivedTime]", True          ' 최신부터 읽음
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Finished
            If Left$(LCase$(VStr(AV(arr, r, C_CLASS))), 8) = "ipm.note" Then
                eid = VStr(AV(arr, r, C_EID))
                k1 = VStr(AV(arr, r, C_MSGID))
                If Len(Trim$(k1)) > 0 Then k1 = "id:" & LCase$(Trim$(k1)) Else k1 = ""
                k2 = FallbackKeyOf(VStr(AV(arr, r, C_SND)), VStr(AV(arr, r, C_SUBJ)), VDate(AV(arr, r, C_SENT)))
                key = ""
                If Len(k1) > 0 And keepId.Exists(k1) Then
                    key = k1
                ElseIf Len(k2) > 0 And keepId.Exists(k2) Then
                    key = k2
                End If
                If Len(key) > 0 Then
                    ' 앞서 본(더 최근에 온) 같은 메일이 중복 -> 지울 목록에, 지금(더 먼저 온) 메일을 남김
                    If Abs(DateDiff("n", keepTime(key), t)) <= DUP_WINDOW_MINUTES Then dupSet(keepId(key)) = True
                End If
                If Len(k1) > 0 Then
                    keepId(k1) = eid
                    keepTime(k1) = t
                End If
                If Len(k2) > 0 Then
                    keepId(k2) = eid
                    keepTime(k2) = t
                End If
            End If
        Next r
        DoEvents
    Loop
Finished:
    For Each k In dupSet.Keys
        c.Add CStr(k)
    Next k
Done:
End Function


'==========================================================================
'  sleeppm 단독 메일 옮기기
'    받는 사람(To/CC)에 sleeppm 이 있고, loungepm 은 없고, 내 주소도 직접 없고,
'    보낸 사람이 sleeppm_senders.txt 에 있으면 -> 정해 둔 폴더로 (읽지 않음 그대로)
'==========================================================================
Private Sub MoveGroupMail(ByVal itm As Object)
    Dim fld As Object, why As String
    On Error Resume Next
    If Not GROUP_MOVE_ENABLED Then Exit Sub
    If Not ShouldMove(itm, why) Then Exit Sub
    Set fld = GetMoveFolder(False)
    If fld Is Nothing Then Exit Sub
    If itm.Parent.EntryID = fld.EntryID Then Exit Sub
    itm.Move fld
End Sub

' 옮길 메일이면 True. why = 판단 이유 (TestSleeppm 에서 보여 줌)
Private Function ShouldMove(ByVal itm As Object, ByRef why As String) As Boolean
    Dim r As Object, a As String, hasGroup As Boolean, hasSkip As Boolean, toMe As Boolean, snd As String
    On Error GoTo Fail
    If itm.Class <> olMail Then why = "일반 메일이 아님": Exit Function
    LoadSenders
    snd = SenderKey(itm)
    If Not SenderListed(snd) Then why = "보낸 사람이 목록에 없음": Exit Function
    For Each r In itm.Recipients
        If r.Type <> olBCC Then
            a = RecipKey(r)
            If IsMe(a) Then toMe = True
            a = Replace(a, " ", "")
            If InStr(1, a, MOVE_GROUP, vbBinaryCompare) > 0 Then hasGroup = True
            If InStr(1, a, SKIP_GROUP, vbBinaryCompare) > 0 Then hasSkip = True
        End If
    Next r
    ' 받는 사람/참조 표시 글자로도 확인 (그룹 이름이 "Sleep PM" 처럼 보이는 경우)
    a = LCase$(Replace(itm.To & ";" & itm.CC, " ", ""))
    If InStr(1, a, MOVE_GROUP, vbBinaryCompare) > 0 Then hasGroup = True
    If InStr(1, a, SKIP_GROUP, vbBinaryCompare) > 0 Then hasSkip = True
    If Not hasGroup Then why = "받는 사람에 " & MOVE_GROUP & " 이 없음": Exit Function
    If hasSkip Then why = SKIP_GROUP & " 도 받는 사람에 있음": Exit Function
    If toMe Then why = "내 주소가 받는 사람/참조에 직접 있음": Exit Function
    why = MOVE_GROUP & " 으로만 받은 메일 + 목록에 있는 보낸 사람"
    ShouldMove = True
    Exit Function
Fail:
    why = "확인 중 오류"
End Function

' 보낸 사람: "|주소|이름|" (소문자)
Private Function SenderKey(ByVal itm As Object) As String
    Dim s As String, eu As Object
    On Error Resume Next
    s = GetPropStr(itm.PropertyAccessor, PR_SENDER_SMTP_ADDRESS)
    If Len(s) = 0 Then
        If LCase$(itm.SenderEmailType) = "ex" Then
            Set eu = itm.Sender.GetExchangeUser
            If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
        End If
    End If
    SenderKey = "|" & LCase$(Trim$(s)) & "|" & LCase$(Trim$(itm.SenderName)) & "|" & LCase$(Trim$(itm.SenderEmailAddress)) & "|"
End Function

Private Function SenderListed(ByVal key As String) As Boolean
    Dim parts() As String, i As Long
    parts = Split(m_Senders, "|")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 2 Then
            If InStr(1, key, parts(i), vbBinaryCompare) > 0 Then
                SenderListed = True
                Exit Function
            End If
        End If
    Next i
End Function

' 받는 사람: "주소|이름" (소문자). 그룹(배포 목록)은 그룹 주소
Private Function RecipKey(ByVal r As Object) As String
    Dim s As String, ae As Object, eu As Object, dl As Object
    On Error Resume Next
    s = GetPropStr(r.PropertyAccessor, PR_SMTP_ADDRESS)
    If Len(s) = 0 Then
        Set ae = r.AddressEntry
        If Not ae Is Nothing Then
            Set eu = ae.GetExchangeUser
            If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
            If Len(s) = 0 Then
                Set dl = ae.GetExchangeDistributionList
                If Not dl Is Nothing Then s = dl.PrimarySmtpAddress
            End If
        End If
    End If
    If Len(s) = 0 Then s = r.Address
    RecipKey = LCase$(Trim$(s)) & "|" & LCase$(Trim$(r.Name)) & "|" & LCase$(Trim$(r.Address))
End Function

Private Function IsMe(ByVal rk As String) As Boolean
    Dim addr As String, p As Long
    p = InStr(1, rk, "|")
    If p > 0 Then addr = Left$(rk, p - 1) Else addr = rk
    If Len(addr) = 0 Or InStr(1, addr, "@") = 0 Then Exit Function
    IsMe = (InStr(1, MyAddrs(), ";" & addr & ";", vbTextCompare) > 0)
End Function

Private Function MyAddrs() As String
    Dim s As String, eu As Object, acc As Object
    If Len(m_MyAddrs) > 0 Then MyAddrs = m_MyAddrs: Exit Function
    On Error Resume Next
    Set eu = Application.Session.CurrentUser.AddressEntry.GetExchangeUser
    If Not eu Is Nothing Then s = eu.PrimarySmtpAddress
    For Each acc In Application.Session.Accounts
        s = s & ";" & acc.SmtpAddress
    Next acc
    m_MyAddrs = ";" & LCase$(Replace(s, " ", "")) & ";"
    MyAddrs = m_MyAddrs
End Function

' --- 보낸 사람 목록 파일 ---
Private Function SendersFilePath() As String
    SendersFilePath = Environ$("APPDATA") & "\ReplyAlert\sleeppm_senders.txt"
End Function

Private Sub LoadSenders()
    Dim path As String, src As String, stamp As String, fso As Object
    On Error Resume Next
    path = SendersFilePath()
    src = MACRO_FOLDER & "\sleeppm_senders.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then fso.CreateFolder fso.GetParentFolderName(path)
    ' 목록이 없거나 비어 있으면: 매크로 저장 폴더의 sleeppm_senders.txt 를 복사 (없으면 빈 양식)
    If Len(Dir$(path)) = 0 Or ParseSenders(ReadUtf8(path)) = "|" Then
        If Len(Dir$(src)) > 0 Then
            fso.CopyFile src, path, True
            m_SendersStamp = ""
        ElseIf Len(Dir$(path)) = 0 Then
            WriteUtf8 path, "# " & MOVE_GROUP & " 으로만 받은 메일 중, 아래 사람이 보낸 메일은 '" & m_MoveFolderName & "' 폴더로 옮깁니다." & vbCrLf & _
                            "# 한 줄에 한 사람 (메일 주소 또는 Outlook 에 보이는 이름). # 으로 시작하는 줄은 무시" & vbCrLf
        End If
    End If
    stamp = CStr(FileDateTime(path))
    If stamp = m_SendersStamp And Len(m_SendersStamp) > 0 Then Exit Sub
    m_SendersStamp = stamp
    m_Senders = ParseSenders(ReadUtf8(path))
End Sub

' 파일 내용 -> "|주소|이름|" (소문자). "이름 <주소>" 형식이면 주소만
Private Function ParseSenders(ByVal txt As String) As String
    Dim lines() As String, i As Long, ln As String, lst As String, a As Long, b As Long
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    lst = "|"
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(Replace(Replace(lines(i), ChrW(65279), ""), vbTab, " "))
        If Left$(ln, 1) <> "#" And Len(ln) > 2 Then
            a = InStr(1, ln, "<")
            b = InStr(1, ln, ">")
            If a > 0 And b > a Then
                lst = AddSender(lst, Left$(ln, a - 1))                ' 이름
                lst = AddSender(lst, Mid$(ln, a + 1, b - a - 1))      ' 주소
            Else
                lst = AddSender(lst, ln)
            End If
        End If
    Next i
    ParseSenders = lst
End Function

' 목록에 추가. 메일 주소면 @ 앞부분도 (Exchange 내부 주소로 들어온 메일도 찾도록)
Private Function AddSender(ByVal lst As String, ByVal v As String) As String
    Dim p As Long
    v = LCase$(Trim$(Replace(v, "|", "")))
    If Len(v) > 2 Then lst = lst & v & "|"
    p = InStr(1, v, "@")
    If p > 4 Then lst = lst & Left$(v, p - 1) & "|"
    AddSender = lst
End Function

' --- 옮길 폴더 ---
Private Function GetMoveFolder(ByVal allowPick As Boolean) As Object
    Dim id As String, fld As Object
    On Error Resume Next
    If Not m_MoveFolder Is Nothing Then
        id = ""
        id = m_MoveFolder.EntryID
        If Len(id) > 0 Then Set GetMoveFolder = m_MoveFolder: Exit Function
        Set m_MoveFolder = Nothing
    End If
    id = GetSetting(SETTINGS_APP, "SleepFilter", "FolderId", "")
    If Len(id) > 0 Then Set fld = Application.Session.GetFolderFromID(id)
    If fld Is Nothing Then
        ' 이름으로 찾기: 받은 편지함 아래 -> 메일함 전체
        Set fld = FindFolderByName(Application.Session.GetDefaultFolder(olFolderInbox), m_MoveFolderName, 0)
        If fld Is Nothing Then Set fld = FindFolderByName(Application.Session.GetDefaultFolder(olFolderInbox).Parent, m_MoveFolderName, 0)
        If Not fld Is Nothing Then SaveSetting SETTINGS_APP, "SleepFilter", "FolderId", fld.EntryID
    End If
    If fld Is Nothing And allowPick Then Set fld = PickMoveFolder()
    Set m_MoveFolder = fld
    Set GetMoveFolder = fld
End Function

Private Function FindFolderByName(ByVal pf As Object, ByVal nm As String, ByVal depth As Long) As Object
    Dim f As Object, hit As Object
    On Error Resume Next
    If depth > 4 Then Exit Function
    For Each f In pf.Folders
        If f.DefaultItemType = olMailItem Then
            If StrComp(Trim$(f.Name), nm, vbTextCompare) = 0 Then
                Set FindFolderByName = f
                Exit Function
            End If
        End If
    Next f
    For Each f In pf.Folders
        If f.DefaultItemType = olMailItem Then
            Set hit = FindFolderByName(f, nm, depth + 1)
            If Not hit Is Nothing Then
                Set FindFolderByName = hit
                Exit Function
            End If
        End If
    Next f
End Function

Private Function PickMoveFolder() As Object
    Dim fld As Object
    On Error Resume Next
    MsgBoxU "sleeppm 메일을 옮길 폴더를 골라 주세요. (예: " & m_MoveFolderName & ")", vbInformation
    Set fld = Application.Session.PickFolder
    If fld Is Nothing Then Exit Function
    If fld.DefaultItemType <> olMailItem Then
        MsgBoxU "메일 폴더를 골라 주세요.", vbExclamation
        Exit Function
    End If
    SaveSetting SETTINGS_APP, "SleepFilter", "FolderId", fld.EntryID
    Set m_MoveFolder = fld
    Set PickMoveFolder = fld
End Function

' --- 수동 실행 매크로 ---
' 받은 편지함에 이미 들어와 있는 sleeppm 단독 메일을 한 번에 옮김
Public Sub MoveSleeppmNow()
    Dim v As String, days As Long, fld As Object, tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_NAME As Long, C_SND As Long, C_SADDR As Long
    Dim cand As New Collection, eid As Variant, itm As Object, why As String, toMove As New Collection, n As Long
    Dim reasons As Object, k As Variant, msg As String
    Init
    LoadSenders
    If Len(m_Senders) <= 1 Then
        MsgBoxU "보낸 사람 목록이 비어 있습니다. EditSleeppmSenders 로 먼저 적어 주세요.", vbExclamation
        Exit Sub
    End If
    Set fld = GetMoveFolder(True)
    If fld Is Nothing Then Exit Sub
    v = InputBox("How many days back should I check the Inbox?" & vbCrLf & "(e.g. 30, 90, 365)", "DupRemover", "90")
    If Len(Trim$(v)) = 0 Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    days = CLng(v)
    If days < 1 Then Exit Sub
    cutoff = Now - days

    On Error Resume Next
    ' 1) 보낸 사람으로 먼저 거르기 (빠름)
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_NAME = AddCol(tbl, "SenderName")
    C_SND = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    C_SADDR = AddCol(tbl, "SenderEmailAddress")
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Scanned
            If Left$(LCase$(VStr(AV(arr, r, C_CLASS))), 8) = "ipm.note" Then
                If SenderListed("|" & LCase$(VStr(AV(arr, r, C_SND))) & "|" & LCase$(VStr(AV(arr, r, C_NAME))) & "|" & _
                                LCase$(VStr(AV(arr, r, C_SADDR))) & "|") Then
                    cand.Add VStr(AV(arr, r, C_EID))
                End If
            End If
        Next r
        DoEvents
    Loop
Scanned:
    ' 2) 받는 사람 확인
    Set reasons = CreateObject("Scripting.Dictionary")
    For Each eid In cand
        Set itm = Nothing
        Set itm = GetItemSafe(CStr(eid))
        If Not itm Is Nothing Then
            If ShouldMove(itm, why) Then
                toMove.Add CStr(eid)
            Else
                reasons(why) = reasons(why) + 1
            End If
        End If
        n = n + 1
        If n Mod 20 = 0 Then DoEvents
    Next eid
    On Error GoTo 0

    If toMove.Count = 0 Then
        msg = "최근 " & days & "일 받은 편지함에 옮길 " & MOVE_GROUP & " 메일이 없습니다." & vbCrLf & vbCrLf & _
              "목록에 있는 사람이 보낸 메일: " & cand.Count & "건"
        For Each k In reasons.Keys
            msg = msg & vbCrLf & "  - " & k & ": " & reasons(k) & "건"
        Next k
        If cand.Count = 0 Then msg = msg & vbCrLf & vbCrLf & "(보낸 사람을 못 찾았습니다. 메일 하나를 선택하고 TestSleeppm 을 실행해 결과를 알려 주세요)"
        MsgBoxU msg, vbInformation
        Exit Sub
    End If
    If MsgBoxU("최근 " & days & "일 받은 편지함에서 " & MOVE_GROUP & " 으로만 받은 메일 " & toMove.Count & "건을 찾았습니다." & vbCrLf & vbCrLf & _
               "'" & fld.Name & "' 폴더로 옮길까요?", vbYesNo + vbQuestion) <> 6 Then Exit Sub
    n = 0
    On Error Resume Next
    For Each eid In toMove
        Set itm = Nothing
        Set itm = GetItemSafe(CStr(eid))
        If Not itm Is Nothing Then
            itm.Move fld
            If Err.Number = 0 Then n = n + 1
            Err.Clear
        End If
        If n Mod 20 = 0 Then DoEvents
    Next eid
    On Error GoTo 0
    MsgBoxU n & "건을 '" & fld.Name & "' 폴더로 옮겼습니다.", vbInformation
End Sub

' 보낸 사람 목록 파일을 메모장으로 열기 (저장하면 바로 적용)
Public Sub EditSleeppmSenders()
    Init
    LoadSenders
    Shell "notepad.exe " & Chr$(34) & SendersFilePath() & Chr$(34), vbNormalFocus
End Sub

' 옮길 폴더 다시 고르기
Public Sub SetSleeppmFolder()
    Dim fld As Object
    Init
    Set fld = PickMoveFolder()
    If Not fld Is Nothing Then MsgBoxU "앞으로 '" & fld.Name & "' 폴더로 옮깁니다.", vbInformation
End Sub

' 선택한 메일이 옮겨질 메일인지 확인
Public Sub TestSleeppm()
    Dim itm As Object, why As String, ok As Boolean, fld As Object, nm As String
    Init
    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBoxU "메일을 하나 선택한 뒤 실행하세요.", vbExclamation
        Exit Sub
    End If
    ok = ShouldMove(itm, why)
    Set fld = GetMoveFolder(False)
    If fld Is Nothing Then nm = "(폴더를 못 찾음 - SetSleeppmFolder 로 골라 주세요)" Else nm = fld.Name
    MsgBoxU "제목: " & itm.Subject & vbCrLf & _
            "보낸 사람: " & itm.SenderName & vbCrLf & _
            "받는 사람: " & itm.To & vbCrLf & _
            "참조: " & itm.CC & vbCrLf & vbCrLf & _
            IIf(ok, "==> 옮길 메일입니다: ", "==> 옮기지 않습니다: ") & why & vbCrLf & _
            "옮길 폴더: " & nm, vbInformation
End Sub


'==========================================================================
'  유틸리티
'==========================================================================
Private Sub Init()
    If Len(m_Title) > 0 Then Exit Sub
    m_Title = "중복 메일 정리"
    m_MoveFolderName = "지워도 되는"
    Set m_Id = CreateObject("Scripting.Dictionary")
    Set m_Time = CreateObject("Scripting.Dictionary")
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
    Dim st As Object
    On Error GoTo Done
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
Done:
End Sub

Private Function GetItemSafe(ByVal entryId As String) As Object
    On Error Resume Next
    If Len(entryId) = 0 Then Exit Function
    Set GetItemSafe = Application.Session.GetItemFromID(entryId)
End Function

Private Function GetPropStr(ByVal pa As Object, ByVal tag As String) As String
    On Error Resume Next
    GetPropStr = CStr(pa.GetProperty(tag))
End Function

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

' "김" 같은 표기를 실제 글자로 (빌드 때 한글을 이렇게 바꿔 둠)
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
    If Len(title) = 0 Then title = m_Title
    MsgBoxU = MessageBoxW(0, StrPtr(prompt), StrPtr(title), buttons Or &H10000)
End Function
