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
Private m_MoveFolderName As String
Private m_MoveFolder As Object
Private m_Senders As String
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
' ThisOutlookSession Application_NewMailEx (ReplyAlert )
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
'==========================================================================
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
        MsgBoxU U("\uCD5C\uADFC ") & days & U("\uC77C \uBC1B\uC740 \uD3B8\uC9C0\uD568\uC5D0 \uC911\uBCF5 \uBA54\uC77C\uC774 \uC5C6\uC2B5\uB2C8\uB2E4."), vbInformation
        Exit Sub
    End If
    If MsgBoxU(U("\uCD5C\uADFC ") & days & U("\uC77C \uBC1B\uC740 \uD3B8\uC9C0\uD568\uC5D0\uC11C \uC911\uBCF5 \uBA54\uC77C ") & dups.Count & U("\uAC74\uC744 \uCC3E\uC558\uC2B5\uB2C8\uB2E4.") & vbCrLf & vbCrLf & _
               U("\uBA3C\uC800 \uC628 \uBA54\uC77C\uB9CC \uB0A8\uAE30\uACE0, \uB098\uBA38\uC9C0 ") & dups.Count & U("\uAC74\uC744 \uC9C0\uC6B4 \uD3B8\uC9C0\uD568\uC73C\uB85C \uC62E\uAE38\uAE4C\uC694?") & vbCrLf & _
               U("(\uC9C0\uC6B4 \uD3B8\uC9C0\uD568\uC5D0\uC11C \uB418\uC0B4\uB9B4 \uC218 \uC788\uC2B5\uB2C8\uB2E4)"), vbYesNo + vbQuestion) <> 6 Then Exit Sub   ' 6 =
    For Each eid In dups
        Set itm = GetItemSafe(CStr(eid))
        If Not itm Is Nothing Then
            itm.Delete
            n = n + 1
            If n Mod 20 = 0 Then DoEvents
        End If
    Next eid
    MsgBoxU U("\uC911\uBCF5 \uBA54\uC77C ") & n & U("\uAC74\uC744 \uC9C0\uC6B4 \uD3B8\uC9C0\uD568\uC73C\uB85C \uC62E\uACBC\uC2B5\uB2C8\uB2E4."), vbInformation
End Sub

Public Sub TestDuplicate()
    Dim itm As Object, k1 As String, k2 As String, msg As String, other As String
    Init
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
    EnsureSeeded
    k1 = IdKey(itm)
    k2 = FallbackKey(itm)
    If Len(k1) > 0 And m_Id.Exists(k1) Then
        If m_Id(k1) <> itm.EntryID Then other = m_Id(k1)
    End If
    If Len(other) = 0 And m_Id.Exists(k2) Then
        If m_Id(k2) <> itm.EntryID Then other = m_Id(k2)
    End If
    msg = U("\uC81C\uBAA9: ") & itm.Subject & vbCrLf & _
          U("\uBC1B\uC740 \uC2DC\uAC01: ") & itm.ReceivedTime & vbCrLf & _
          "Message-ID: " & IIf(Len(k1) = 0, U("(\uC5C6\uC74C)"), Mid$(k1, 4)) & vbCrLf & vbCrLf
    If Len(other) = 0 Then
        msg = msg & U("==> \uCD5C\uADFC ") & DUP_SEED_DAYS & U("\uC77C \uBC1B\uC740 \uBA54\uC77C \uC911 \uAC19\uC740 \uBA54\uC77C\uC774 \uC5C6\uC2B5\uB2C8\uB2E4.")
    Else
        msg = msg & U("==> \uAC19\uC740 \uBA54\uC77C\uC774 \uD558\uB098 \uB354 \uC788\uC2B5\uB2C8\uB2E4. (") & DUP_WINDOW_MINUTES & U("\uBD84 \uC548\uC5D0 \uB4E4\uC5B4\uC628 \uACBD\uC6B0 \uB098\uC911 \uAC83\uC744 \uC9C0\uC6C0)")
    End If
    MsgBoxU msg, vbInformation
End Sub


'==========================================================================
'==========================================================================
' True
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
                itm.Delete
                CheckItem = True
                Exit Function
            End If
        End If
    End If
    Remember k1, eid, rcv
    Remember k2, eid, rcv
Done:
End Function

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

' Outlook ,
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
    tbl.Sort "[ReceivedTime]", True
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

' days (EntryID) .
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
    tbl.Sort "[ReceivedTime]", True
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
' sleeppm
' (To/CC) sleeppm , loungepm , ,
' sleeppm_senders.txt -> ( )
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

' True. why = (TestSleeppm )
Private Function ShouldMove(ByVal itm As Object, ByRef why As String) As Boolean
    Dim r As Object, a As String, hasGroup As Boolean, hasSkip As Boolean, toMe As Boolean, snd As String
    On Error GoTo Fail
    If itm.Class <> olMail Then why = U("\uC77C\uBC18 \uBA54\uC77C\uC774 \uC544\uB2D8"): Exit Function
    LoadSenders
    snd = SenderKey(itm)
    If Not SenderListed(snd) Then why = U("\uBCF4\uB0B8 \uC0AC\uB78C\uC774 \uBAA9\uB85D\uC5D0 \uC5C6\uC74C"): Exit Function
    For Each r In itm.Recipients
        If r.Type <> olBCC Then
            a = RecipKey(r)
            If InStr(1, a, MOVE_GROUP, vbBinaryCompare) > 0 Then hasGroup = True
            If InStr(1, a, SKIP_GROUP, vbBinaryCompare) > 0 Then hasSkip = True
            If IsMe(a) Then toMe = True
        End If
    Next r
    If Not hasGroup Then why = U("\uBC1B\uB294 \uC0AC\uB78C\uC5D0 ") & MOVE_GROUP & U(" \uC774 \uC5C6\uC74C"): Exit Function
    If hasSkip Then why = SKIP_GROUP & U(" \uB3C4 \uBC1B\uB294 \uC0AC\uB78C\uC5D0 \uC788\uC74C"): Exit Function
    If toMe Then why = U("\uB0B4 \uC8FC\uC18C\uAC00 \uBC1B\uB294 \uC0AC\uB78C/\uCC38\uC870\uC5D0 \uC9C1\uC811 \uC788\uC74C"): Exit Function
    why = MOVE_GROUP & U(" \uC73C\uB85C\uB9CC \uBC1B\uC740 \uBA54\uC77C + \uBAA9\uB85D\uC5D0 \uC788\uB294 \uBCF4\uB0B8 \uC0AC\uB78C")
    ShouldMove = True
    Exit Function
Fail:
    why = U("\uD655\uC778 \uC911 \uC624\uB958")
End Function

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
    If Len(s) = 0 Then s = itm.SenderEmailAddress
    SenderKey = "|" & LCase$(Trim$(s)) & "|" & LCase$(Trim$(itm.SenderName)) & "|"
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
    RecipKey = LCase$(Trim$(s)) & "|" & LCase$(Trim$(r.Name))
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
    ' : sleeppm_senders.txt ( )
    If Len(Dir$(path)) = 0 Or ParseSenders(ReadUtf8(path)) = "|" Then
        If Len(Dir$(src)) > 0 Then
            fso.CopyFile src, path, True
            m_SendersStamp = ""
        ElseIf Len(Dir$(path)) = 0 Then
            WriteUtf8 path, "# " & MOVE_GROUP & U(" \uC73C\uB85C\uB9CC \uBC1B\uC740 \uBA54\uC77C \uC911, \uC544\uB798 \uC0AC\uB78C\uC774 \uBCF4\uB0B8 \uBA54\uC77C\uC740 '") & m_MoveFolderName & U("' \uD3F4\uB354\uB85C \uC62E\uAE41\uB2C8\uB2E4.") & vbCrLf & _
                            U("# \uD55C \uC904\uC5D0 \uD55C \uC0AC\uB78C (\uBA54\uC77C \uC8FC\uC18C \uB610\uB294 Outlook \uC5D0 \uBCF4\uC774\uB294 \uC774\uB984). # \uC73C\uB85C \uC2DC\uC791\uD558\uB294 \uC904\uC740 \uBB34\uC2DC") & vbCrLf
        End If
    End If
    stamp = CStr(FileDateTime(path))
    If stamp = m_SendersStamp And Len(m_SendersStamp) > 0 Then Exit Sub
    m_SendersStamp = stamp
    m_Senders = ParseSenders(ReadUtf8(path))
End Sub

Private Function ParseSenders(ByVal txt As String) As String
    Dim lines() As String, i As Long, ln As String, lst As String, a As Long, b As Long
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    lst = "|"
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(Replace(Replace(lines(i), ChrW(65279), ""), vbTab, " "))
        a = InStr(1, ln, "<")
        b = InStr(1, ln, ">")
        If a > 0 And b > a Then ln = Mid$(ln, a + 1, b - a - 1)
        ln = LCase$(Trim$(ln))
        If Len(ln) > 2 And Left$(ln, 1) <> "#" Then lst = lst & Replace(ln, "|", "") & "|"
    Next i
    ParseSenders = lst
End Function

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
    MsgBoxU U("sleeppm \uBA54\uC77C\uC744 \uC62E\uAE38 \uD3F4\uB354\uB97C \uACE8\uB77C \uC8FC\uC138\uC694. (\uC608: ") & m_MoveFolderName & ")", vbInformation
    Set fld = Application.Session.PickFolder
    If fld Is Nothing Then Exit Function
    If fld.DefaultItemType <> olMailItem Then
        MsgBoxU U("\uBA54\uC77C \uD3F4\uB354\uB97C \uACE8\uB77C \uC8FC\uC138\uC694."), vbExclamation
        Exit Function
    End If
    SaveSetting SETTINGS_APP, "SleepFilter", "FolderId", fld.EntryID
    Set m_MoveFolder = fld
    Set PickMoveFolder = fld
End Function

' sleeppm
Public Sub MoveSleeppmNow()
    Dim v As String, days As Long, fld As Object, tbl As Object, arr As Variant, r As Long, t As Date, cutoff As Date
    Dim C_EID As Long, C_CLASS As Long, C_RCV As Long, C_NAME As Long, C_SND As Long
    Dim cand As New Collection, eid As Variant, itm As Object, why As String, toMove As New Collection, n As Long
    Init
    LoadSenders
    If Len(m_Senders) <= 1 Then
        MsgBoxU U("\uBCF4\uB0B8 \uC0AC\uB78C \uBAA9\uB85D\uC774 \uBE44\uC5B4 \uC788\uC2B5\uB2C8\uB2E4. EditSleeppmSenders \uB85C \uBA3C\uC800 \uC801\uC5B4 \uC8FC\uC138\uC694."), vbExclamation
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
    ' 1) ()
    Set tbl = Application.Session.GetDefaultFolder(olFolderInbox).GetTable
    tbl.Columns.RemoveAll
    C_EID = AddCol(tbl, "EntryID")
    C_CLASS = AddCol(tbl, "MessageClass")
    C_RCV = AddCol(tbl, "ReceivedTime")
    C_NAME = AddCol(tbl, "SenderName")
    C_SND = AddCol(tbl, PR_SENDER_SMTP_ADDRESS)
    tbl.Sort "[ReceivedTime]", True
    Do Until tbl.EndOfTable
        arr = tbl.GetArray(TABLE_CHUNK)
        If Not IsArray(arr) Then Exit Do
        For r = LBound(arr, 1) To UBound(arr, 1)
            t = VDate(AV(arr, r, C_RCV))
            If t <> 0 And t < cutoff Then GoTo Scanned
            If Left$(LCase$(VStr(AV(arr, r, C_CLASS))), 8) = "ipm.note" Then
                If SenderListed("|" & LCase$(VStr(AV(arr, r, C_SND))) & "|" & LCase$(VStr(AV(arr, r, C_NAME))) & "|") Then
                    cand.Add VStr(AV(arr, r, C_EID))
                End If
            End If
        Next r
        DoEvents
    Loop
Scanned:
    ' 2)
    For Each eid In cand
        Set itm = Nothing
        Set itm = GetItemSafe(CStr(eid))
        If Not itm Is Nothing Then
            If ShouldMove(itm, why) Then toMove.Add CStr(eid)
        End If
        n = n + 1
        If n Mod 20 = 0 Then DoEvents
    Next eid
    On Error GoTo 0

    If toMove.Count = 0 Then
        MsgBoxU U("\uCD5C\uADFC ") & days & U("\uC77C \uBC1B\uC740 \uD3B8\uC9C0\uD568\uC5D0 \uC62E\uAE38 ") & MOVE_GROUP & U(" \uBA54\uC77C\uC774 \uC5C6\uC2B5\uB2C8\uB2E4."), vbInformation
        Exit Sub
    End If
    If MsgBoxU(U("\uCD5C\uADFC ") & days & U("\uC77C \uBC1B\uC740 \uD3B8\uC9C0\uD568\uC5D0\uC11C ") & MOVE_GROUP & U(" \uC73C\uB85C\uB9CC \uBC1B\uC740 \uBA54\uC77C ") & toMove.Count & U("\uAC74\uC744 \uCC3E\uC558\uC2B5\uB2C8\uB2E4.") & vbCrLf & vbCrLf & _
               "'" & fld.Name & U("' \uD3F4\uB354\uB85C \uC62E\uAE38\uAE4C\uC694?"), vbYesNo + vbQuestion) <> 6 Then Exit Sub
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
    MsgBoxU n & U("\uAC74\uC744 '") & fld.Name & U("' \uD3F4\uB354\uB85C \uC62E\uACBC\uC2B5\uB2C8\uB2E4."), vbInformation
End Sub

Public Sub EditSleeppmSenders()
    Init
    LoadSenders
    Shell "notepad.exe " & Chr$(34) & SendersFilePath() & Chr$(34), vbNormalFocus
End Sub

Public Sub SetSleeppmFolder()
    Dim fld As Object
    Init
    Set fld = PickMoveFolder()
    If Not fld Is Nothing Then MsgBoxU U("\uC55E\uC73C\uB85C '") & fld.Name & U("' \uD3F4\uB354\uB85C \uC62E\uAE41\uB2C8\uB2E4."), vbInformation
End Sub

Public Sub TestSleeppm()
    Dim itm As Object, why As String, ok As Boolean, fld As Object, nm As String
    Init
    On Error Resume Next
    Set itm = Application.ActiveExplorer.Selection.Item(1)
    On Error GoTo 0
    If itm Is Nothing Then
        MsgBoxU U("\uBA54\uC77C\uC744 \uD558\uB098 \uC120\uD0DD\uD55C \uB4A4 \uC2E4\uD589\uD558\uC138\uC694."), vbExclamation
        Exit Sub
    End If
    ok = ShouldMove(itm, why)
    Set fld = GetMoveFolder(False)
    If fld Is Nothing Then nm = U("(\uD3F4\uB354\uB97C \uBABB \uCC3E\uC74C - SetSleeppmFolder \uB85C \uACE8\uB77C \uC8FC\uC138\uC694)") Else nm = fld.Name
    MsgBoxU U("\uC81C\uBAA9: ") & itm.Subject & vbCrLf & _
            U("\uBCF4\uB0B8 \uC0AC\uB78C: ") & itm.SenderName & vbCrLf & _
            U("\uBC1B\uB294 \uC0AC\uB78C: ") & itm.To & vbCrLf & _
            U("\uCC38\uC870: ") & itm.CC & vbCrLf & vbCrLf & _
            IIf(ok, U("==> \uC62E\uAE38 \uBA54\uC77C\uC785\uB2C8\uB2E4: "), U("==> \uC62E\uAE30\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4: ")) & why & vbCrLf & _
            U("\uC62E\uAE38 \uD3F4\uB354: ") & nm, vbInformation
End Sub


'==========================================================================
'==========================================================================
Private Sub Init()
    If Len(m_Title) > 0 Then Exit Sub
    m_Title = U("\uC911\uBCF5 \uBA54\uC77C \uC815\uB9AC")
    m_MoveFolderName = U("\uC9C0\uC6CC\uB3C4 \uB418\uB294")
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
    If Len(title) = 0 Then title = m_Title
    MsgBoxU = MessageBoxW(0, StrPtr(prompt), StrPtr(title), buttons Or &H10000)
End Function
