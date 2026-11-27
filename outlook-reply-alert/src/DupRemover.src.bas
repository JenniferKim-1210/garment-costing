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
'==========================================================================
Option Explicit

'============================== SETTINGS ==============================
Private Const DUP_ENABLED As Boolean = True
Private Const DUP_WINDOW_MINUTES As Long = 30       ' remove only when the copies arrive within N minutes of each other
Private Const DUP_SEED_DAYS As Long = 1             ' remember mail received in the last N days (for copies arriving after a restart)
'======================================================================

Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const PR_SENDER_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"
Private Const TABLE_CHUNK As Long = 200

Private m_Id As Object      ' key -> EntryID of the copy that is kept
Private m_Time As Object    ' key -> received time of that copy
Private m_Seeded As Boolean
Private m_Title As String

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
    If Not DUP_ENABLED Then Exit Sub
    EnsureSeeded
    ids = Split(EntryIDCollection, ",")
    For i = LBound(ids) To UBound(ids)
        Set itm = Nothing
        Set itm = Application.Session.GetItemFromID(Trim$(ids(i)))
        If Not itm Is Nothing Then CheckItem itm
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
'  유틸리티
'==========================================================================
Private Sub Init()
    If Len(m_Title) > 0 Then Exit Sub
    m_Title = "중복 메일 정리"
    Set m_Id = CreateObject("Scripting.Dictionary")
    Set m_Time = CreateObject("Scripting.Dictionary")
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
