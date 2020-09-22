Attribute VB_Name = "mPolyRec"
Option Explicit

'+------------------+----------------------------------+
'| mPolyRec.bas     | developed in Visual Basic 6.0    |
'+---------+--------+----------------------------------+
'| Release | Development - June 20 2006 - 60620        |
'+---------+-------+-----------------------------------+
'| Original author | dafhi                             |
'+-------------+---+-----------------------------------+
'| Description | Musical Event Recorder                |
'+-------------+---------------------------------------+
'| - Dependencies -                                    |
'| mMySynth.bas                                        |
'| mGeneral.bas                                        |
'| mAudioBase.bas                                      |
'+-------------+---------------------------------------+

'Example:

'Dim MySong As Song
'
'Private Sub Form_Load()
'
' Call Song_NewTrack(MySong, [MaxPoly], [MinPoly]) 'Arm the recorder
'
' A. MaxPoly is the number of elements that are allowed simultaneously.
' B. MaxInputVal is the maximum value given by your input system
'   Example 1 - MIDI (127) or QWERTY keyboard (255)
'   Example 2 - Musical Track with 460 note events
' C. MinPoly when set lower than MaxPoly can be used to ignore
'   the most distant notes, rather than remove them from memory
'   Example:  SetPoly (poly 16), , (minpoly 2)
'   1. hold in this order -> CDEF .. only EF will play
'   2. release F .. DE will play
'   3. release E .. CD will play
'
'End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeySpace Then
'        Record_EVENT mSong, 0, TIME_MARK
'    Else
'        If Record_QWERTY(mSong, KeyCode, NOTE_ON) Then
'            'it's recording
'        End If
'    End If
'End Sub

'Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    Record_QWERTY mSong, KeyCode
'End Sub

'You wont hear anything with this example.

'The following two subs show how to access
'the recorded data

'------------------------------------------------------'

'Private Sub z_LoopThroughTracks(pSong As Song)
'Dim L1 As Long
'    For L1 = 1 To pSong.Proc.cTrack
'        z_LoopThroughSequence pSong.Track(L1).OnOff.Seq
'    Next
'End Sub
'Private Sub z_LoopThroughSequence(pSeq As SequencerElement)
'Dim L1 As Long, lElem As Long, sTime As Single
'    lElem = pSeq.Events.LinkHead
'    For L1 = 1 To pSeq.Events.StackPtr
'        sTime = pSeq.NINFO(lElem).abs_time 'there are other members to (NoteEvent) type
'        lElem = pSeq.Events.Link(lElem).Next 'linked list
'    Next L1
'End Sub

'------------------------------------------------------'

'Sequencer:
' A. Sub Sequencer() for event architecture
' B. mMySynth.bas for synthesizer algorithm

Public Const gMAX_TRACKS As Byte = 32
Public Const gMAX_EVENTS As Integer = 32000

Public gNoteIndex As Long

Dim mEventType As Long
Dim mNoteVal   As Long
Dim mCodeRmv   As Long
Dim mElemRmv   As Long
Dim mFreq      As Double
Dim mExpliciTime As Single
Public gTimeNow  As Single

' ===== Linked list

Private Type mSwapStor  'salvage as much as possible
    Elem1    As Integer 'if polyphony is lowered
    Elem2    As Integer
End Type '.. keeps track of swapped list elements
         'for user defined arrays that associate
         'with the linked list

Private Type mSwapInfo
    Swap()   As mSwapStor
    Size     As Long
End Type
                
Private Type mLink
    Next  As Integer
    Prev  As Integer
    Code  As Long    'Example:
End Type 'PolyAppend MyController, MidiNote60 'note on
         'PolyRemove MyController, MidiNote60 'note off

Public Type LinkSys
    ArySize  As Long
    MinPoly  As Long
    MaxInput As Long  '0 to 255 etc (say, MidiNoteVal)
    StackPtr As Long
    LinkHead As Long
    LinkTail As Long
    Link()   As mLink '1 to ArySize
    ReMv()   As Long  '1 to MaxInput .. for link removal.
    SwapInfo As mSwapInfo 'preserve much during polyphony
End Type                  'change while notes playing


' ===== Note System

Public Type NoteEvent
    dFreq    As Double
    abs_time As Single
    sNote    As Single
    Event    As Integer
    Ref      As Integer
                          '1. Key 65 is down (not same as unhappy)
                          '2. link list element 104 is available
                          '3. Mov Ref(65), 104
                          '4. Key 65 is up
                          '5. list element 104 becomes available again
End Type

Public Type SequencerElement
    Events   As LinkSys
    NINFO()  As NoteEvent
    RecElem  As Long
    PlayElem As Long
    LngTimeD As Long
    StkPtrP  As Long
End Type

Public Type ChanInf
    dPos     As Double 'add dPos2 etc. for more FM operators.  also see NoteInfo type below
    sVol     As Single
End Type
                         
Public Type NoteInfo
    InfoL    As ChanInf
    InfoR    As ChanInf
    dIncr    As Double 'add dIncr2 etc. for more FM operators ..
                       'confused?  value = Sin(ChanInf.dPos + Sin(ChanInf.dPos2))
                       'Add ChanInf.dPos, dIncr: Add ChanInf.dPos2, dIncr2
    iVol     As Single
    rlsTime  As Single
End Type

Public Type PolyphonyCoupling
    Links    As LinkSys
    NI()     As NoteInfo
End Type

Public Type SeqAndPoly
    Poly     As PolyphonyCoupling
    Seq      As SequencerElement
End Type

Public Type TrackQ
    OnOff    As SeqAndPoly
    tMute    As Boolean
End Type

Dim I      As Long
Dim mElem  As Long
Dim mLink  As mLink

Public gWriteLen As Long

Private Type SongProc
    NextDIM    As Integer
    cTrack     As Integer
    Ptr        As Integer
    PtrP       As Integer
    song_len   As Single
    song_lenP  As Single
    KeyDwn()   As Long
End Type

Public Type Song
    Track()  As TrackQ
    Proc     As SongProc
End Type

Public gIsPaused     As Boolean

'fast forward / rewind
Dim mTimePrv   As Single
Dim mTimeNow   As Single

Dim mLngStart  As Currency
Dim timeNowL   As Currency
Dim timeFreq   As Currency

Dim mKeyDown(255) As Integer
Dim mSoftKeyUP    As Boolean

''''''''''''''''''''''''''''''
' Event Recording            '
''''''''''''''''''''''''''''''

'First, call this sub
Public Function Song_NewTrack(pSong As Song, Optional ByVal MaxPoly As Integer = 255, Optional ByVal MinPoly As Long = -1) As Byte
Dim NewTrkCount As Byte, DimMOD As Byte
    
    If pSong.Proc.cTrack < gMAX_TRACKS Then
        NewTrkCount = pSong.Proc.cTrack + 1
        pSong.Proc.cTrack = NewTrkCount
        If NewTrkCount > pSong.Proc.NextDIM Then
            DimMOD = pSong.Proc.NextDIM + 16
            pSong.Proc.NextDIM = DimMOD
            ReDim Preserve pSong.Track(1 To DimMOD)
            For DimMOD = NewTrkCount To DimMOD
                z_SetPoly pSong.Track(DimMOD), MaxPoly, MinPoly
            Next
        End If
        pSong.Proc.PtrP = pSong.Proc.Ptr
        pSong.Proc.Ptr = NewTrkCount
        If NewTrkCount = 1 Then
            gIsPaused = True
            QueryPerformanceFrequency timeFreq
            Play_Restart pSong
        Else
            Song_ClearData pSong, NewTrkCount
        End If
    End If
    Song_NewTrack = NewTrkCount

End Function

'Then pick one of these two
Public Function Record_EVENT(pSong As Song, ByVal pNoteVal As Integer, Optional ByVal pEventType As Long = NOTE_OFF, Optional ByVal pFreq As Double, Optional ByVal pTime As Single = -1) As Long

    Record_EVENT = 0
    
    If gIsPaused Then Exit Function
    
    If pSong.Proc.Ptr <= pSong.Proc.cTrack And pSong.Proc.cTrack > 0 Then
        z_RecEVENT pSong, pSong.Track(pSong.Proc.Ptr), pSong.Track(pSong.Proc.Ptr).OnOff, Record_EVENT, pNoteVal, pEventType, pFreq, pTime
    End If

End Function
Public Function Record_QWERTY(pSong As Song, ByVal KeyCode As Integer, Optional ByVal pEventType As Long = NOTE_OFF, Optional ByVal pTime As Single = -1) As Long
Dim lNoteIncr As Double

    Record_QWERTY = False
        
    gNoteIndex = NoteIndexFromQwerty(KeyCode)
    If SampleFreq <> 0 Then lNoteIncr = GetFreq(gNoteIndex)
    Record_QWERTY = Record_EVENT(pSong, gNoteIndex, pEventType, lNoteIncr, pTime)
    
End Function

'the low level stuff
Private Sub z_RecEVENT(pSong As Song, pPolyS As TrackQ, pSeqP As SeqAndPoly, pRetElem As Long, ByVal pNoteVal As Integer, Optional ByVal pEventType As Long = NOTE_OFF, Optional ByVal pFreq As Double, Optional ByVal pTime As Single = -1)
    
    mExpliciTime = pTime
    QueryPerformanceCounter timeNowL
'    timeNowL = timeGetTime
    
    If pNoteVal > 0 And pSeqP.Poly.Links.ArySize > 0 Then
    
        mNoteVal = pNoteVal 'minimize parameter passing
        mEventType = pEventType
        mFreq = pFreq
        
        If mKeyDown(pNoteVal) = 0 _
          And pEventType = NOTE_ON Then
            
            mCodeRmv = mNoteVal * 2 - 1
            z_RecEVENT2 pSong, pPolyS, pSeqP, pRetElem
            
            If pRetElem = 0 Then Exit Sub
            
            'L1 is the array element for this event
            'in the sequencer's linked list

            mKeyDown(pNoteVal) = pRetElem
        
        ElseIf pEventType = NOTE_OFF Then
        
            If mKeyDown(pNoteVal) > 0 Then
            
                mCodeRmv = mNoteVal * 2
                z_RecEVENT2 pSong, pPolyS, pSeqP, pRetElem
                
                If pRetElem = 0 Then Exit Sub
                
                ' .Ref = stored NOTE_ON input code,
                'which sequencer hands to polyphony
                'to remove the note
                
                pSeqP.Seq.NINFO(pRetElem).Ref = mKeyDown(pNoteVal)
                '# z_NoteOnOff() shows how this is used #
    
                'creating reference for corresponding note on
                pSeqP.Seq.NINFO(mKeyDown(pNoteVal)).Ref = pRetElem
                
                'Reset "hardware"
                If mSoftKeyUP = True Then
                    mSoftKeyUP = False 'z_CheckNoteOffEXIST() creates True condition
                    mKeyDown(pNoteVal) = -1
                Else
                    mKeyDown(pNoteVal) = 0
                End If
                
            ElseIf mKeyDown(pNoteVal) = -1 Then
                mKeyDown(pNoteVal) = 0
            End If
            
        ElseIf pEventType = TIME_MARK Then
            z_RecEVENT2 pSong, pPolyS, pSeqP, pRetElem
        End If
        
    ElseIf pEventType = TIME_MARK Then
        z_RecEVENT2 pSong, pPolyS, pSeqP, pRetElem
    End If

End Sub
Private Sub z_RecEVENT2(pSong As Song, pPolyS As TrackQ, pSeqP As SeqAndPoly, pL1 As Long)
Dim ldeltaSplice As Single
    
    If pSeqP.Seq.Events.StackPtr = 0 Then
        If mExpliciTime < 0 Then
            z_SngLngPosD gTimeNow, (timeNowL - mLngStart)
            pSeqP.Seq.LngTimeD = 0
        Else
            gTimeNow = mExpliciTime
            z_LngPosSngD pSeqP.Seq.LngTimeD, gTimeNow
        End If
        z_Search pL1, pSeqP, pSeqP.Seq.Events.Link
    Else
        If mExpliciTime < 0 Then
            QueryPerformanceCounter timeNowL
            gTimeNow = (timeNowL - mLngStart) / CSng(timeFreq)
            pSeqP.Seq.LngTimeD = 0
        Else
            gTimeNow = mExpliciTime
        End If
        z_Search pL1, pSeqP, pSeqP.Seq.Events.Link
    End If
    
    If pL1 = 0 Then Exit Sub
    
    pSeqP.Seq.RecElem = pL1 'this just makes search a little less cpu intensive next time around
    
    pSeqP.Seq.NINFO(pL1).abs_time = gTimeNow
    pSeqP.Seq.NINFO(pL1).Event = mEventType
    pSeqP.Seq.NINFO(pL1).dFreq = mFreq
    pSeqP.Seq.NINFO(pL1).sNote = mNoteVal
    
    If gTimeNow > pSong.Proc.song_len Then pSong.Proc.song_len = gTimeNow
    
    'will try to figure out the holes some time ..
    'if playback element is farther down the list than
    'newest element, (rare - it "shouldn't" happen after
    'z_Search but does) this loop will rewind
    'the playback element
    
    mElem = pL1
    Do While True
        If mElem = pSeqP.Seq.PlayElem Then
            pSeqP.Seq.PlayElem = pSeqP.Seq.Events.Link(pL1).Prev
            If mExpliciTime >= 0 Then
                z_LngPosSngD pSeqP.Seq.LngTimeD, mExpliciTime - pSeqP.Seq.NINFO(pSeqP.Seq.PlayElem).abs_time
            End If
            Exit Do
        ElseIf mElem = pSeqP.Seq.Events.LinkTail Then
            Exit Do
        End If
        mElem = pSeqP.Seq.Events.Link(mElem).Next
    Loop
    
End Sub


''''''''''''''''''''''''''''''
' Event Recording - Search   '
''''''''''''''''''''''''''''''


Private Sub z_Search(pElem As Long, pSeqP As SeqAndPoly, pLink() As mLink)
Dim lTime As Single, lFound As Long, lTime2 As Single, ElemP As Long
        
    'z_Search adjusts pElem to equal the link
    'array element that timeNow must precede.
    
    'pElem is then passed to PolyInsert which
    'returns newest link element
    
    pElem = pSeqP.Seq.RecElem
    
    lTime2 = pSeqP.Seq.NINFO(pElem).abs_time
    
    'error trap - note off occurs before note on ..
    '.. happens in stress test
    mElem = mKeyDown(mNoteVal)
'    mElem = pSeqP.Poly.KeyDown(mNoteVal)
    If mElem > 0 Then
        lTime = pSeqP.Seq.NINFO(mElem).abs_time
        If lTime2 < lTime Then
            mElem = pSeqP.Seq.Events.Link(mElem).Next
            pElem = z_PolyInsert(pSeqP.Seq.Events, mElem, mCodeRmv, False)
            gTimeNow = lTime
            Exit Sub
        End If
    End If
    
    If lTime2 < gTimeNow Then
        Do While True 'infinite
            lTime = pSeqP.Seq.NINFO(pElem).abs_time
            If gTimeNow < lTime Then
                Exit Do
            ElseIf gTimeNow = lTime Then
                z_Search_TimeEqual pElem, pSeqP
                Exit Do
            ElseIf pElem = pSeqP.Seq.Events.LinkTail Then
                pElem = pSeqP.Seq.Events.Link(pElem).Next
                Exit Do
            End If
            pElem = pSeqP.Seq.Events.Link(pElem).Next
        Loop
    ElseIf gTimeNow < lTime2 Then
        Do While True
            pElem = pSeqP.Seq.Events.Link(pElem).Prev
            lTime = pSeqP.Seq.NINFO(pElem).abs_time
            If lTime < gTimeNow Then
                pElem = pSeqP.Seq.Events.Link(pElem).Next
                Exit Do
            ElseIf gTimeNow = lTime Then
                z_Search_TimeEqual pElem, pSeqP
                Exit Do
            ElseIf pElem = pSeqP.Seq.Events.LinkHead Then
                Exit Do
            End If
        Loop
    Else 'lTime2 = timeNow
        z_Search_TimeEqual pElem, pSeqP
    End If
        
    pElem = z_PolyInsert(pSeqP.Seq.Events, pElem, mCodeRmv, False)
    
    'Returned value = array element for new link
    'that immediately precedes link value given
    
    'pElem is now the newest link value
    
End Sub
Private Sub z_Search_TimeEqual(pElem As Long, pSeqP As SeqAndPoly)

    'What's done in here:
    
    ' 1. if a particular note has a length
    'of zero, NOTE_GONE link element should
    'immediately follow the NOTE_ON
    
    ' 2a. otherwise All NOTE_GONEs go before
    'NOTE_ONs for best polyphony handling
    
    ' 2b. The latest event goes last in line
    'within the NOTE_GONE or NOTE_ON group
    
    If mEventType = NOTE_GONE Then
        z_NOTE_GONE_TimeE pSeqP, pElem
    Else 'latest event assumed as NOTE_ON for now
        Do While pElem <> pSeqP.Seq.Events.LinkTail
            pElem = pSeqP.Seq.Events.Link(pElem).Next
            If gTimeNow < pSeqP.Seq.NINFO(pElem).abs_time Then
                Exit Sub
            End If
        Loop
        pElem = pSeqP.Seq.Events.Link(pElem).Next
    End If
    
End Sub
Private Function z_NOTE_GONE_TimeE(pSeqP As SeqAndPoly, pElem As Long) As Long
Dim lElemSave As Long, lFirstOn As Long
Dim lFoundSomething As Long

    lElemSave = pElem
    
    'search backward
    Do While True
        If pSeqP.Seq.NINFO(pElem).abs_time < gTimeNow Then
            pElem = pSeqP.Seq.Events.Link(pElem).Next
            Exit Do
        ElseIf mKeyDown(mNoteVal) = pElem Then
'        ElseIf pSeqP.Poly.KeyDwn(mNoteVal) = pElem Then
            lFirstOn = pElem
            pElem = pSeqP.Seq.Events.Link(pElem).Next
            lFoundSomething = pElem
            Exit Do
        ElseIf pSeqP.Seq.NINFO(pElem).Event = NOTE_ON Then
            lFirstOn = pElem 'track first NOTE_ON as best
            'starting place if no zero-length note is found
        ElseIf pElem = pSeqP.Seq.Events.LinkHead Then
            Exit Do
        End If
        pElem = pSeqP.Seq.Events.Link(pElem).Prev
    Loop
    
    If lFoundSomething = 0 Then 'search forward
        If lFirstOn > 0 Then lElemSave = lFirstOn
        pElem = lElemSave
        Do While True
            If mKeyDown(mNoteVal) = pElem Then
'            If pSeqP.Poly.KeyDwn(mNoteVal) = pElem Then
                pElem = pSeqP.Seq.Events.Link(pElem).Next
                lFoundSomething = pElem 'note with length zero
                Exit Do
            ElseIf gTimeNow < pSeqP.Seq.NINFO(pElem).abs_time Then
                lFoundSomething = pElem
                Exit Do
            ElseIf pElem = pSeqP.Seq.Events.LinkTail Then
                If pSeqP.Seq.NINFO(pElem).Event = NOTE_GONE Then
                    pElem = pSeqP.Seq.Events.Link(pElem).Next
                End If
                lFoundSomething = pElem
                Exit Do
            ElseIf pSeqP.Seq.NINFO(pElem).Event = NOTE_ON Then
                If lFirstOn = 0 Then lFirstOn = pElem
            End If
            pElem = pSeqP.Seq.Events.Link(pElem).Next
        Loop
    End If
    
    If lFoundSomething = 0 Then
        If lFirstOn > 0 Then lElemSave = lFirstOn
        pElem = lElemSave
        Do While pElem <> pSeqP.Seq.Events.LinkHead
            If pSeqP.Seq.NINFO(pElem).Event = NOTE_GONE Then
                pElem = pSeqP.Seq.Events.Link(pElem).Next
                lFoundSomething = pElem
                Exit Do
            End If
            pElem = pSeqP.Seq.Events.Link(pElem).Prev
        Loop
    End If
    
End Function
Private Sub z_LngPosSngD(pRetTimeD As Long, ByVal pSngTimeD As Single)
    'global PCM from module mAudioBASE
    pRetTimeD = Int(pSngTimeD * PCM.lSamplesPerSec + 0.5)
End Sub
Private Sub z_SngLngPosD(pRetSng As Single, ByVal pLngPosD As Long)
    pRetSng = pLngPosD / timeFreq
End Sub


Public Sub Play_Restart(pSong As Song)
Dim L1 As Long, lPtr As Long, lIsPaused As Boolean

    If pSong.Proc.cTrack < 1 Then Exit Sub

    lIsPaused = gIsPaused
    gIsPaused = False
    lPtr = pSong.Proc.Ptr
    
    Song_CheckHeldEvents pSong
    
    For L1 = 1 To pSong.Proc.cTrack
        z_PlayRestart pSong, pSong.Track(L1).OnOff
    Next L1
    gIsPaused = lIsPaused
    pSong.Proc.Ptr = lPtr
    
    pSong.Proc.song_lenP = pSong.Proc.song_len
'    mLngStart = timeGetTime
    QueryPerformanceCounter mLngStart
    timeNowL = mLngStart
    gTimeNow = 0
    
End Sub
Public Sub Song_CheckHeldEvents(pSong As Song)
    I = pSong.Proc.Ptr
    If I > pSong.Proc.cTrack Or I < 1 Then Exit Sub
    z_CheckNoteOffEXIST pSong, pSong.Track(I).OnOff.Seq
End Sub
Private Sub z_CheckNoteOffEXIST(pSong As Song, pSeq As SequencerElement)
Dim L1 As Long

    If pSeq.StkPtrP < 1 Then pSeq.StkPtrP = 1
    
    'last-recorded events are linear linked list references ..
    'no .LinkHead and Link().Next to deal with here
    
    For L1 = pSeq.StkPtrP To pSeq.Events.StackPtr
        If pSeq.NINFO(L1).Ref = 0 Then
            If pSeq.NINFO(L1).Event = NOTE_ON Then
                mSoftKeyUP = True
                Record_EVENT pSong, pSeq.NINFO(L1).sNote
                mSoftKeyUP = False
            End If
        End If
    Next
    
    'LinkHead and .Next are used for musical event time

End Sub
Private Sub z_PlayRestart(pSong As Song, pSeqP As SeqAndPoly)
    
    'StkPtrP used for 'rare' event
    
    '1. user has notes held
    '2. Play_Restart request before NOTE_OFF have occurred
    pSeqP.Seq.StkPtrP = pSeqP.Seq.Events.StackPtr + 1
    
    'NOTE_OFF events will be created by z_CheckNoteOffEXIST() ..
    'for such notes, and it is best to only search newest
    'events in the sequencer ..
    
    'Imagine CDE held down.  At this point,
    'StackPtr = 3, StkPtrP = 0 (changed to 1 if less than 1)
    'Play_Restart requested
    ' z_CheckNoteOffEXIST() (inside play_restart)
    'For L1 = {1 To 3}
    ' finds 3 NOTE_ON with no NOTE_OFF .. (fixes problem)
    'Next
    
    'StackPtr is now 6
    
    'StkPtrP = StackPtr + 1
    
    'User holds down 3 more notes
    'Play_Restart
    ' z_CheckNoteOffEXIST()
    'For L1 = {7 To 9} ..
    'Next
    
    'StkPtrP has now left the building
    
    pSeqP.Seq.PlayElem = pSeqP.Seq.Events.Link(pSeqP.Seq.Events.LinkHead).Prev
    mElem = pSeqP.Seq.Events.LinkHead
    
    z_LngPosSngD pSeqP.Seq.LngTimeD, pSeqP.Seq.NINFO(mElem).abs_time
    z_SilenceTrack pSeqP.Poly.Links

End Sub
Public Sub SilenceTracks(pSong As Song)
Dim L1 As Long
    For L1 = 1 To pSong.Proc.cTrack
        z_SilenceTrack pSong.Track(L1).OnOff.Poly.Links
    Next
End Sub
Private Sub z_SilenceTrack(pLinkS As LinkSys)
    mElem = pLinkS.LinkHead
    For I = 1 To pLinkS.StackPtr
        z_PolyRemove pLinkS, pLinkS.Link(mElem).Code
        mElem = pLinkS.Link(mElem).Next
    Next
End Sub
Public Sub Song_ClearData(pSong As Song, Optional ByVal pTrackNum As Integer = -1)
Dim L1 As Long, lLen As Single
    
    pSong.Proc.song_len = 0
    
    If pTrackNum < 0 Then
        If pSong.Proc.cTrack > 0 Then
            pSong.Proc.song_len = 0
            For L1 = 1 To pSong.Proc.cTrack
                z_EraseTrack pSong, pSong.Track(L1)
            Next
            pSong.Proc.cTrack = 1
            pSong.Proc.Ptr = 1
            Play_Restart pSong
            If Not gIsPaused Then
                Song_TogglePause
            End If
        End If
    Else
        pSong.Proc.song_len = 0
        z_EraseTrack pSong, pSong.Track(pTrackNum)
        For L1 = 1 To pSong.Proc.cTrack
            lLen = GetTrackLength(pSong.Track(L1))
            If lLen > pSong.Proc.song_len Then pSong.Proc.song_len = lLen
        Next
    
    End If

End Sub
Private Sub z_EraseTrack(pSong As Song, pTrack As TrackQ)

    mElem = pTrack.OnOff.Seq.Events.LinkHead
    Do While True
        pTrack.OnOff.Seq.NINFO(mElem).Ref = 0
        If mElem = pTrack.OnOff.Seq.Events.LinkTail Then
            pTrack.OnOff.Seq.NINFO(mElem).abs_time = 0
            Exit Do
        End If
        mElem = pTrack.OnOff.Seq.Events.Link(mElem).Next
    Loop
    pTrack.OnOff.Seq.Events.StackPtr = 0
    pTrack.OnOff.Seq.RecElem = pTrack.OnOff.Seq.Events.LinkHead
    
End Sub
Public Sub Song_TogglePause(Optional pRetPause As Boolean)
    gIsPaused = Not gIsPaused
    If gIsPaused Then
'        timeNowL = timeGetTime
        QueryPerformanceCounter timeNowL
    End If
    pRetPause = gIsPaused
End Sub

Public Function GetTrackLength(pTrk As TrackQ) As Single
Dim lLen As Single
    If pTrk.OnOff.Seq.Events.ArySize > 0 Then
        lLen = pTrk.OnOff.Seq.NINFO(pTrk.OnOff.Seq.Events.LinkTail).abs_time
    End If
    GetTrackLength = lLen
End Function


''''''''''''''''''''''''''''''
' Playback                   '
''''''''''''''''''''''''''''''

Public Sub Sequencer(pSong As Song, sngRender() As Single, ByVal pSamPos1 As Long, ByVal pSamPos2 As Long)
Dim L1 As Long
    
    GetTimeNow
    
    For L1 = 1 To pSong.Proc.cTrack
        z_Sequencer pSong.Track(L1), sngRender, pSamPos1, pSamPos2
    Next
    
End Sub
Private Sub z_Sequencer(pPoly As TrackQ, sngRender() As Single, ByVal pSamPos1 As Long, ByVal pSamPos2 As Long)
Dim lForw   As Long
    
    Do While pSamPos1 <= pSamPos2
        If pPoly.OnOff.Seq.LngTimeD = 0 Then
            If pPoly.OnOff.Seq.PlayElem <> pPoly.OnOff.Seq.Events.LinkTail And _
               pPoly.OnOff.Seq.Events.StackPtr > 0 And _
               gIsPaused = False Then
                'sample position has caught up to an event
                z_TmeDelta_Zero pPoly, pPoly.OnOff, pPoly.OnOff.Seq.PlayElem
            Else
                'run out of events, so play whatever's in polyphony
                '(NOTE_GONEs may not have been recorded :)
                z_PlayPoly pPoly, sngRender, pSamPos1, pSamPos2
                pSamPos1 = pSamPos2 + PCM.nChannels
            End If
            
        Else 'sample position catching up to an event
            lForw = pSamPos1 + pPoly.OnOff.Seq.LngTimeD * PCM.nChannels - PCM.nChannels
            If lForw > pSamPos2 Then 'stream write length < next event
                z_PlayPoly pPoly, sngRender, pSamPos1, pSamPos2
                pSamPos1 = pSamPos2 + PCM.nChannels
            Else
                z_PlayPoly pPoly, sngRender, pSamPos1, lForw
                pSamPos1 = lForw + PCM.nChannels
            End If
        End If 'LngTimeD = 0

    Loop
        
End Sub
Private Sub z_TmeDelta_Zero(pPolyS As TrackQ, pSeqP As SeqAndPoly, pElem As Long)

    mElem = pSeqP.Seq.Events.Link(pElem).Next
    mTimeNow = pSeqP.Seq.NINFO(mElem).abs_time

    Do While True
        pElem = mElem
        mTimePrv = mTimeNow
        z_NoteOnOff pSeqP.Poly, pElem, pSeqP.Seq.NINFO(pElem)
        mElem = pSeqP.Seq.Events.Link(pElem).Next
        mTimeNow = pSeqP.Seq.NINFO(mElem).abs_time
        If pElem = pSeqP.Seq.Events.LinkTail Then
            Exit Do
        Else
            z_LngPosSngD pSeqP.Seq.LngTimeD, mTimeNow - mTimePrv
            Exit Do
        End If
    Loop
    
End Sub
Private Sub z_TmeDelta_Zero2T(pPolyS As TrackQ, pSeqP As SeqAndPoly, pElem As Long)
Dim lTimeD As Single, lTime As Long

    If pElem = pSeqP.Seq.Events.LinkTail Then
        mTimePrv = pSeqP.Seq.NINFO(pElem).abs_time
    Else
        mTimePrv = 0
    End If
    
    pElem = pSeqP.Seq.Events.Link(pElem).Next
    Do While True
        z_NoteOnOff pSeqP.Poly, pElem, pSeqP.Seq.NINFO(pElem)
        mTimeNow = pSeqP.Seq.NINFO(pElem).abs_time
        mTimePrv = mTimeNow
        If pElem = pSeqP.Seq.Events.LinkTail Then
            Exit Do
        Else
            pElem = pSeqP.Seq.Events.Link(pElem).Next
            mTimeNow = pSeqP.Seq.NINFO(pElem).abs_time
            z_LngPosSngD pSeqP.Seq.LngTimeD, mTimeNow - mTimePrv
            If pSeqP.Seq.LngTimeD > 0 Then Exit Do
        End If
    Loop
    
End Sub
Private Sub z_NoteOnOff(pPoly As PolyphonyCoupling, pElem As Long, pEvent As NoteEvent)
Dim L1 As Long, lInput As Long
        
    If pEvent.Event = NOTE_ON Then
        If z_PolyAppend(pPoly.Links, pElem) > 0 Then
            With pPoly.NI(pPoly.Links.LinkTail)
                .dIncr = pEvent.dFreq
                .InfoL.dPos = 0
                .InfoL.sVol = 1
                .InfoR = .InfoL
            End With
        End If
    ElseIf pEvent.Event = NOTE_OFF Then
        z_PolyRemove pPoly.Links, pEvent.Ref
    End If
    
End Sub


Private Sub z_PlayPoly(pPoly As TrackQ, sngRender() As Single, pSamPos1 As Long, ByVal pSamPos2 As Long)

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Polyphony handled by linked list .. ordinary   '
    ' arrays in the TrackQ                           '
    '                                                '
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    
    gWriteLen = (pSamPos2 - pSamPos1 + PCM.nChannels) / PCM.nChannels
    
    If PCM.nChannels < 1 Then Exit Sub
    
    z_PlayPoly2 pPoly.OnOff, sngRender, pSamPos1, pSamPos2
    
End Sub
Private Sub z_PlayPoly2(pSeqP As SeqAndPoly, sngRender() As Single, pSamPos1 As Long, ByVal pSamPos2 As Long)
Dim I1 As Long, lElem As Long

    If Not gIsPaused Then
    If pSeqP.Seq.Events.StackPtr > 0 Then
        If pSeqP.Seq.PlayElem <> pSeqP.Seq.Events.LinkTail Then
            Add pSeqP.Seq.LngTimeD, -gWriteLen
        End If
    End If
    End If
    
    If pSeqP.Poly.Links.MinPoly < 1 Then
        
        'polyphonic array element dictated by the link system
        lElem = pSeqP.Poly.Links.LinkHead
        
        For I1 = 1 To pSeqP.Poly.Links.StackPtr
            'mMySynth.bas ..
            StereoArchitecture pSeqP, pSeqP.Poly.NI(lElem), sngRender, pSamPos1, pSamPos2
            lElem = pSeqP.Poly.Links.Link(lElem).Next
        Next
    
    Else 'pPoly.polyex.MinPoly > 0
    
        lElem = pSeqP.Poly.Links.LinkTail
        
        For I1 = 1 To LMin(pSeqP.Poly.Links.StackPtr, pSeqP.Poly.Links.MinPoly)
            StereoArchitecture pSeqP, pSeqP.Poly.NI(lElem), sngRender, pSamPos1, pSamPos2
            lElem = pSeqP.Poly.Links.Link(lElem).Prev
        Next
    
    End If 'pPoly.polyex.MinPoly < 1
    
End Sub


''''''''''''''''''''''''''''''
' Initialization / Handling  '
''''''''''''''''''''''''''''''

Private Sub z_SetPoly(pController As TrackQ, Optional ByVal MaxPoly As Integer = 255, Optional ByVal MinPoly As Long = -1)

    If MinPoly > MaxPoly Then MinPoly = MaxPoly
    
    z_SetLinks pController.OnOff.Seq.Events, gMAX_EVENTS
    Erase pController.OnOff.Seq.NINFO
    ReDim pController.OnOff.Seq.NINFO(1 To gMAX_EVENTS)
    
    z_SetPoly2 pController.OnOff.Poly, MaxPoly, MinPoly, gMAX_EVENTS
    
    pController.OnOff.Seq.RecElem = pController.OnOff.Seq.Events.LinkHead
    pController.OnOff.Seq.PlayElem = pController.OnOff.Seq.Events.Link(pController.OnOff.Seq.Events.LinkHead).Prev
    
End Sub
Private Sub z_SetPoly2(pPoly As PolyphonyCoupling, MaxPoly As Integer, MinPoly As Long, Optional ByVal MaxInputVal As Long = gMAX_EVENTS)
    z_SetLinks pPoly.Links, MaxPoly, MaxInputVal
    pPoly.Links.MinPoly = MinPoly
    Erase pPoly.NI
    ReDim pPoly.NI(1 To MaxPoly)
'    Erase pPoly.KeyDwn
'    ReDim pPoly.KeyDwn(0 To 255)
End Sub

Private Sub z_SetLinks(pLink As LinkSys, Optional ByVal ArySize As Integer = 255, Optional ByVal MaxInputVal As Integer = 255, Optional bPreserve As Boolean = False)
Dim Begin_    As Long
Dim Elem_     As Long
Dim StorP     As Long
Dim StorN     As Long
Dim lNumOut   As Long
Dim StorOut() As Integer
    
    If ArySize < 1 Then 'garbage in, no error
        ArySize = 1
        pLink.StackPtr = 0
    End If
    If MaxInputVal < 0 Then MaxInputVal = 255
    
    If ArySize = pLink.ArySize Then GoTo DEUX
    
    If bPreserve And pLink.ArySize > 0 Then
    
        lNumOut = pLink.StackPtr - ArySize
            
        If ArySize > 255 Then 'assuming sequencer data
        
            'Example: C D E F G are recorded, but array
            'size is changed to 4, remove G
        
            Elem_ = pLink.LinkTail
            For I = 1 To lNumOut
                z_PolyRemove pLink, pLink.ReMv(pLink.Link(Elem_).Code)
                Elem_ = pLink.Link(Elem_).Prev
            Next
        
        Else 'assuming hardware controller
    
            'Example: in order, C D E are playing, and
            'polyphony is changed to 2, remove C
        
            Elem_ = pLink.LinkHead
            For I = 1 To lNumOut
                z_PolyRemove pLink, pLink.ReMv(pLink.Link(Elem_).Code)
                Elem_ = pLink.Link(Elem_).Next
            Next
        
        End If 'ArySize > 255
        
        'Rewire referenced array indices > new ArySize
        
        '1. Store ouf-of-bound references
        
        ReDim StorOut(1 To lNumOut)
        Elem_ = pLink.LinkHead
        For I = 1 To ArySize
            If Elem_ > ArySize Then
                Add StorN, 1
                StorOut(StorN) = Elem_
            End If
            Elem_ = pLink.Link(Elem_).Next
        Next
        
        Erase pLink.SwapInfo.Swap
        ReDim pLink.SwapInfo.Swap(1 To StorN)
        pLink.SwapInfo.Size = StorN
        
        '2. Find unused links (past .LinkTail)
        'of array index <= new ArySize ..
        
        I = 1
        Elem_ = pLink.Link(pLink.LinkTail).Next
        Do While I <= StorN
            If Elem_ <= ArySize Then 'found in-bounds
                pLink.SwapInfo.Swap(I).Elem1 = Elem_
                pLink.SwapInfo.Swap(I).Elem2 = StorOut(I)
                
                '3. Swap used (at or before link tail)
                'out-of-bound (elem ref > ArySize)
                'with unused in-bounds, reconnect
                'neighbors,
                z_PolySwapOut pLink, StorOut(I), Elem_
                                     
                Add I, 1
            End If
            Elem_ = pLink.Link(Elem_).Next
        Loop
        
        ReDim Preserve pLink.Link(1 To ArySize)
        pLink.ArySize = ArySize
        
        Begin_ = pLink.ArySize
        If Begin_ < 1 Then Begin_ = 1
        
    Else
        
        Erase pLink.Link
        ReDim pLink.Link(1 To ArySize)
        
        pLink.StackPtr = 0
        pLink.LinkHead = 1
        pLink.LinkTail = 1
        
        Begin_ = 1
    
    End If
    
    For I = Begin_ To ArySize - 1
        'To help yourself visualize what is going on here,
        'imagine four relay members on a team.
        
        'They have numbers on their shirt corresponding to
        'position.
        
        'Person 2 is carrying a baton that has a red "1"
        '(meaning previous person was the first runner)
        'and a green "3" (number of the next person).
        
        'Likewise, person 3 has a baton with a red 2 and
        'a green 4.
    
        pLink.Link(I).Next = I + 1
        pLink.Link(I + 1).Prev = I
    Next
    
    'connect the ends
    pLink.Link(1).Prev = ArySize
    pLink.Link(I).Next = 1
    pLink.ArySize = ArySize
    
DEUX:
    
    Erase pLink.ReMv
    ReDim pLink.ReMv(0 To MaxInputVal)
        
    pLink.MaxInput = MaxInputVal
        
End Sub
Private Function z_PolySwapOut(pLinkS As LinkSys, ByVal pElemRemv As Integer, ByVal pElemAvail As Integer) As Long
    z_PolySwapOut2 pLinkS, pLinkS.Link(pElemRemv), pLinkS.Link(pElemAvail), pElemRemv, pElemAvail
End Function
Private Function z_PolySwapOut2(pLinkS As LinkSys, pLinkRemv As mLink, pLinkAvail As mLink, pElemR As Integer, pElemA As Integer) As Long
    mLink = pLinkRemv
    pLinkRemv = pLinkAvail
    pLinkAvail = mLink
    z_PolyRehook pLinkS.Link(pLinkRemv.Prev), _
                 pLinkS.Link(pLinkRemv.Next), pElemA
    z_PolyRehook pLinkS.Link(pLinkAvail.Prev), _
                 pLinkS.Link(pLinkAvail.Next), pElemR
    pLinkS.ReMv(pLinkRemv.Code) = pElemA
    pLinkRemv.Code = 0
    If pLinkS.LinkHead = pElemR Then pLinkS.LinkHead = pElemA
    If pLinkS.LinkTail = pElemR Then pLinkS.LinkTail = pElemA
End Function
Private Sub z_PolyRehook(pLinkP As mLink, pLinkN As mLink, pElem As Integer)
    pLinkP.Next = pElem
    pLinkN.Prev = pElem
End Sub

Private Function z_PolyAppend(pLink As LinkSys, ByVal InputCode As Long, Optional RemoveFirstIfFull As Boolean = True) As Long

    If pLink.ArySize > 0 And InputCode <= pLink.MaxInput And InputCode > 0 Then
        If pLink.ReMv(InputCode) = 0 Then
            If pLink.StackPtr < pLink.ArySize Then
                If pLink.StackPtr = 0 Then
                    pLink.LinkHead = 1
                    pLink.LinkTail = 1
                Else
                    pLink.LinkTail = pLink.Link(pLink.LinkTail).Next
                End If
                pLink.StackPtr = pLink.StackPtr + 1
                z_PolyAppend = pLink.LinkTail
            ElseIf RemoveFirstIfFull Then
                pLink.ReMv(pLink.Link(pLink.LinkHead).Code) = 0
                pLink.LinkHead = pLink.Link(pLink.LinkHead).Next
                pLink.LinkTail = pLink.Link(pLink.LinkTail).Next
                z_PolyAppend = pLink.LinkTail
            Else
                z_PolyAppend = 0
            End If
        Else
            z_PolyAppend = 0
        End If
    Else
        z_PolyAppend = 0
    End If
    
    If z_PolyAppend > 0 Then
        pLink.Link(pLink.LinkTail).Code = InputCode
        pLink.ReMv(InputCode) = pLink.LinkTail
    End If

End Function
Private Function z_PolyRemove(pLink As LinkSys, ByVal InputCode As Long) As Long
Dim lLem As Long
Dim ElemN  As Long
Dim ElemP  As Long

    If pLink.StackPtr > 0 And InputCode <= pLink.MaxInput Then
        lLem = pLink.ReMv(InputCode)
        If lLem <> 0 Then
            If lLem = pLink.LinkTail Then
                If pLink.LinkTail <> pLink.LinkHead Then
                    pLink.LinkTail = pLink.Link(pLink.LinkTail).Prev
                End If
            ElseIf lLem = pLink.LinkHead Then
                pLink.LinkHead = pLink.Link(pLink.LinkHead).Next
            Else
                ElemN = pLink.Link(lLem).Next
                ElemP = pLink.Link(lLem).Prev
                pLink.Link(ElemP).Next = ElemN
                pLink.Link(ElemN).Prev = ElemP
                ElemP = pLink.LinkTail
                ElemN = pLink.Link(ElemP).Next
                pLink.Link(lLem).Prev = ElemP
                pLink.Link(lLem).Next = ElemN
                pLink.Link(ElemP).Next = lLem
                pLink.Link(ElemN).Prev = lLem
            End If
            pLink.ReMv(InputCode) = 0
            pLink.StackPtr = pLink.StackPtr - 1
            z_PolyRemove = 1
        Else
            z_PolyRemove = 0
        End If
    Else
        z_PolyRemove = 0
    End If

End Function
Private Function z_PolyInsert(pLink As LinkSys, ByVal PrecedeWhichLink As Long, ByVal InputCode As Long, Optional RemoveFirstIfFull As Boolean = True) As Long
Dim ElemN  As Long
Dim ElemP As Long

    If pLink.ArySize > 0 And InputCode <= pLink.MaxInput Then
        
        If pLink.StackPtr = 0 Then
            pLink.StackPtr = 1
            z_PolyInsert = pLink.LinkHead
            pLink.LinkTail = z_PolyInsert
            pLink.Link(z_PolyInsert).Code = InputCode
            pLink.ReMv(InputCode) = z_PolyInsert
            Exit Function
        End If
        If pLink.StackPtr < pLink.ArySize Then
            pLink.StackPtr = pLink.StackPtr + 1
            z_PolyInsert = pLink.Link(pLink.LinkTail).Next
        ElseIf RemoveFirstIfFull Then
            pLink.ReMv(pLink.Link(pLink.LinkHead).Code) = 0
            z_PolyInsert = pLink.LinkHead
            pLink.LinkHead = pLink.Link(pLink.LinkHead).Next
        Else
            z_PolyInsert = 0
        End If
    Else
        z_PolyInsert = 0
    End If

    If z_PolyInsert > 0 Then
        pLink.Link(z_PolyInsert).Code = InputCode
        pLink.ReMv(InputCode) = z_PolyInsert
        If pLink.Link(pLink.LinkTail).Next = PrecedeWhichLink Then
            pLink.LinkTail = PrecedeWhichLink
            Exit Function
        End If
        'unhook a link that follows the list tail
        ElemN = pLink.Link(z_PolyInsert).Next
        ElemP = pLink.Link(z_PolyInsert).Prev
        pLink.Link(ElemP).Next = ElemN
        pLink.Link(ElemN).Prev = ElemP
        'snap into place just before 'PrecedeWhichLink'
        ElemP = pLink.Link(PrecedeWhichLink).Prev
        pLink.Link(z_PolyInsert).Prev = ElemP
        pLink.Link(z_PolyInsert).Next = PrecedeWhichLink
        pLink.Link(ElemP).Next = z_PolyInsert
        pLink.Link(PrecedeWhichLink).Prev = z_PolyInsert
        If PrecedeWhichLink = pLink.LinkHead Then
            pLink.LinkHead = z_PolyInsert
        End If
    End If

End Function

Public Sub GetTimeNow(Optional pSongPosition As Boolean)
Dim L1 As Currency

    If gIsPaused Or pSongPosition Then
'        L1 = timeGetTime
        QueryPerformanceCounter L1
'        z_SignedTimeDelta L1, timeNowL
        Add L1, -timeNowL
        
        'continuously bringing timeNowL and mLngStart forward ..
        
'        z_AddTimeDelta timeNowL, L1
        Add timeNowL, L1
'        z_AddTimeDelta mLngStart, L1
        Add mLngStart, L1
    Else
'        timeNowL = timeGetTime
        QueryPerformanceCounter timeNowL
'        z_SignedTimeDelta timeNowL, mLngStart
        Add timeNowL, -mLngStart
        gTimeNow = (timeNowL / CSng(timeFreq))

    End If
End Sub
Public Sub SongPosition(pSong As Song, ByVal pPos_0_To_1 As Single)
Dim timeNew As Single
Dim timePrv As Single
Dim L1 As Currency, L2 As Currency


    If pPos_0_To_1 < 0 Or pPos_0_To_1 > 1 Then Exit Sub
    If pSong.Proc.song_len <= 0 Then Exit Sub
    
    timePrv = gTimeNow
    timeNew = pPos_0_To_1 * pSong.Proc.song_len
'    z_AddTimeDelta mLngStart, Int(timefreq * (timePrv - timeNew) + 0.5) 'vb's Round() function sux ass
    Add mLngStart, Int(timeFreq * (timePrv - timeNew) + 0.5)
    
    GetTimeNow
    If gIsPaused Then
        L1 = timeNowL
'        z_SignedTimeDelta I, mLngStart
        Add L1, -mLngStart
        
        L2 = L1 / timeFreq
        If L2 > pSong.Proc.song_len Then
            gTimeNow = pSong.Proc.song_len
        ElseIf L1 < 0 Then
            gTimeNow = 0
        Else
            gTimeNow = L2
        End If
    End If
    
    For I = 1 To pSong.Proc.cTrack
        z_TrackToPosition pSong.Track(I).OnOff, timePrv
    Next
    
End Sub
Private Sub z_TrackToPosition(pSeqP As SeqAndPoly, pPosPrv As Single)
Dim lElem As Long

    If pSeqP.Seq.Events.StackPtr < 1 Then Exit Sub

    lElem = pSeqP.Seq.PlayElem
    mElem = pSeqP.Seq.Events.Link(lElem).Next
    If mElem = pSeqP.Seq.Events.LinkHead Then lElem = mElem
    
    If gTimeNow < pPosPrv Then 'rewinding
    
        Do While True
            mTimeNow = pSeqP.Seq.NINFO(lElem).abs_time
            If lElem = pSeqP.Seq.Events.LinkHead Then
                If mTimeNow > gTimeNow Then
                    z_LngPosSngD pSeqP.Seq.LngTimeD, mTimeNow - gTimeNow
                Else
                    If lElem = pSeqP.Seq.Events.LinkTail Then
                        pSeqP.Seq.LngTimeD = 0
                    Else
                        mElem = pSeqP.Seq.Events.Link(lElem).Next
                        mTimeNow = pSeqP.Seq.NINFO(mElem).abs_time
                        If mTimeNow > gTimeNow Then
                            z_LngPosSngD pSeqP.Seq.LngTimeD, mTimeNow - gTimeNow
                        Else
                            pSeqP.Seq.LngTimeD = 0
                        End If
                        Exit Do
                    End If
                End If
                z_NoteOnOff_Rewind pSeqP.Poly, lElem, pSeqP.Seq.NINFO(lElem), pSeqP.Seq
                lElem = pSeqP.Seq.Events.Link(lElem).Prev
                Exit Do
            ElseIf mTimeNow < gTimeNow Then
                If lElem <> pSeqP.Seq.Events.LinkTail Then
                    mElem = pSeqP.Seq.Events.Link(lElem).Next
                    z_LngPosSngD pSeqP.Seq.LngTimeD, pSeqP.Seq.NINFO(mElem).abs_time - gTimeNow
                End If
                Exit Do
            End If
            z_NoteOnOff_Rewind pSeqP.Poly, lElem, pSeqP.Seq.NINFO(lElem), pSeqP.Seq
            lElem = pSeqP.Seq.Events.Link(lElem).Prev
        Loop
        pSeqP.Seq.PlayElem = lElem

    ElseIf gTimeNow > pPosPrv Then 'going forward
    
        Do While True
            mTimeNow = pSeqP.Seq.NINFO(lElem).abs_time
            If mTimeNow > gTimeNow Then
                z_LngPosSngD pSeqP.Seq.LngTimeD, mTimeNow - gTimeNow
                Exit Do
            End If
            z_NoteOnOff pSeqP.Poly, lElem, pSeqP.Seq.NINFO(lElem)
            pSeqP.Seq.PlayElem = lElem
            If lElem = pSeqP.Seq.Events.LinkTail Then
                pSeqP.Seq.LngTimeD = 0
                Exit Do
            End If
            lElem = pSeqP.Seq.Events.Link(lElem).Next
            mTimePrv = mTimeNow
        Loop
    
    End If 'timePos < gTimeNow
    
End Sub
Private Sub z_NoteOnOff_Rewind(pPoly As PolyphonyCoupling, pElem As Long, pEvent As NoteEvent, pSeq As SequencerElement)
Dim L1 As Long, lInput As Long
        
    If pEvent.Event = NOTE_ON Then
        z_PolyRemove pPoly.Links, pElem
    ElseIf pEvent.Event = NOTE_OFF Then
        If z_PolyAppend(pPoly.Links, pEvent.Ref) Then
            pPoly.NI(pPoly.Links.LinkTail).dIncr = pSeq.NINFO(pElem).dFreq ' dIncr
        End If
    End If
    
End Sub
