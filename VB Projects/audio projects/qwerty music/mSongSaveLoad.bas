Attribute VB_Name = "mSSaveLoad"
Option Explicit

'+------------------+-----------------------------------+'
'| mSSaveLoad.bas   | developed in Visual Basic 6.0     |'
'+---------+--------+-----------------------------------+'
'| Release | Development - June 20 2006 - 60620         |'
'+---------+-------+------------------------------------+'
'| Original author | dafhi                              |'
'+-------------+---+------------------------------------+'
'| Description | Load / Save  + Save midi file          |'
'+-------------+                                        |'
'|                                                      |'
'| - Dependencies -                                     |'
'| mPolyRec.bas                                         |'
'| -> mGeneral.bas                                      |'
'| -> mAudioBase.bas                                    |'
'|                                                      |'
'| FileDlg2.cls                                         |'
'+------------------------------+-----------------------+'
'| Contributors / Modifications |                       |'
'+------------------------------+                       |'
'|                                                      |'
'|                                                      |'
'+------------------------------------------------------+'

Private Const SAVE_VERSION As Integer = 0

Dim mFreeFile As Integer

Dim mDialogShowing As Boolean
Dim CDLF As OSDialog

Dim IA As Long

Public gMidiDir  As String
Public gSongDir  As String
Dim mStrMidiFile As String
Dim mStrFile     As String
Dim mStr         As String
Dim mBytes()     As Byte    'settings save

Dim PPQN         As Long 'Midi Save

Dim mNoteE()     As NoteEvent 'Save Song (and nested subs)
Dim mRef         As Long

Dim mL1          As Long 'multi-purpose
Dim mL2          As Long
Dim mLA          As Long



Public Function SaveSong(pSong As Song, pFileSpec As String) As Boolean
Dim L1 As Long, lElem As Long, lcTrack As Integer
    
    Song_TogglePause
    
    SaveSong = DialogSuccess(pFileSpec, False, mStrFile, gSongDir, ".dat")
    
    lcTrack = pSong.Proc.cTrack
    For IA = 1 To pSong.Proc.cTrack
        If pSong.Track(lcTrack).OnOff.Seq.Events.StackPtr > 0 Then
            Exit For
        End If
        Add lcTrack, -1
    Next
        
    If lcTrack < 1 Then SaveSong = False
    
    If Not SaveSong Then Exit Function
    
    mStrMidiFile = Left$(mStrFile, Len(mStrFile) - Len(".dat"))
    
    'Open For Binary does not change size of existing file
    Open mStrFile For Output As #1
        Write #1,
    Close #1
    
    Open mStrFile For Binary As #mFreeFile
    
        Seek #mFreeFile, 1
        Put #mFreeFile, , SAVE_VERSION
        Put #mFreeFile, , lcTrack
        
        For IA = 1 To lcTrack
        
            mNoteE = pSong.Track(IA).OnOff.Seq.NINFO
            
            L1 = pSong.Track(IA).OnOff.Seq.Events.StackPtr
            Put #mFreeFile, , L1
            
            lElem = pSong.Track(IA).OnOff.Seq.Events.LinkHead
            For L1 = 1 To L1
                z_Save_MakeRefSequential mNoteE(lElem), L1, lElem, pSong.Track(IA).OnOff.Seq.NINFO
                lElem = pSong.Track(IA).OnOff.Seq.Events.Link(lElem).Next
            Next
            
            lElem = pSong.Track(IA).OnOff.Seq.Events.LinkHead
            For L1 = 1 To L1 - 1
                z_File_Note mNoteE(lElem)
                lElem = pSong.Track(IA).OnOff.Seq.Events.Link(lElem).Next
            Next
        Next
    Close #mFreeFile
    
    Song_TogglePause
    
End Function


Public Function LoadSong(pSong As Song, pFileSpec As String) As Boolean
Dim L1          As Long
Dim lTrackCount As Integer
Dim lVersion    As Integer
Dim lError      As Boolean
    
    LoadSong = DialogSuccess(pFileSpec, , mStrFile, gSongDir, ".dat")
    
    If LoadSong = False Then Exit Function
    
    Open mStrFile For Binary As #mFreeFile
        Seek #mFreeFile, 1
        Get #mFreeFile, , lVersion
        Get #mFreeFile, , lTrackCount 'Same byte length as pSong.Proc.cTrack
        
        Select Case lVersion
        Case 0
        
            If lTrackCount < 1 Or lTrackCount > 32 Then
                LoadSong = False
                GoTo OHNO
            End If
            
            SilenceTracks pSong
            pSong.Proc.cTrack = 0
            
            For L1 = 1 To lTrackCount
                Song_NewTrack pSong
            Next
            pSong.Proc.song_len = 0
    
            For L1 = 1 To pSong.Proc.cTrack
                Song_ClearData pSong, L1
                z_LoadSequence pSong, pSong.Track(L1).OnOff.Seq, lError
                If lError Then
                    LoadSong = False
                    GoTo OHNO
                End If
            Next
        
        End Select
        
        Song_NewTrack pSong
        
        mStrMidiFile = Left$(mStrFile, Len(mStrFile) - Len(".dat"))
        
OHNO:
    Close #mFreeFile
    
End Function


'*------------------------*'
'-                        -'
'-  Nested Load           -'
'-                        -'
'*------------------------*'

Private Sub z_LoadSequence(pSong As Song, pSeq As SequencerElement, pRetError As Boolean)
Dim L1 As Long, lElem As Long, lTime As Single, lTimeP As Single
    
    Get #mFreeFile, , lElem
    
    'making sure data is correct format
    If lElem > gMAX_EVENTS Or lElem < 1 Then
        pRetError = True
        Exit Sub
    End If
    
    pSeq.Events.StackPtr = lElem
    
    z_LoadSafeLims pSeq.NINFO(1), lTime, lTimeP
    z_LoadLngPosSngD pSeq.LngTimeD, lTime 'reset real-time playback
    
    pSeq.PlayElem = gMAX_EVENTS
    
    For L1 = 2 To pSeq.Events.StackPtr
        z_LoadSafeLims pSeq.NINFO(L1), lTime, lTimeP
    Next
    
    pSeq.Events.LinkTail = pSeq.Events.StackPtr
    pSeq.Events.LinkHead = 1
    
    lTime = pSeq.NINFO(pSeq.Events.StackPtr).abs_time
    If lTime > pSong.Proc.song_len Then pSong.Proc.song_len = lTime
    
    'redo the link system
        
    For L1 = 1 To gMAX_EVENTS - 1
        lElem = L1 + 1
        pSeq.Events.Link(L1).Next = lElem
        pSeq.Events.Link(lElem).Prev = L1
    Next
    
    pSeq.Events.Link(1).Prev = L1
    pSeq.Events.Link(L1).Next = 1
    
    pRetError = False

End Sub
Private Sub z_LoadSafeLims(pNote As NoteEvent, pRetTime As Single, pRetTimeP As Single)
Dim lMaxSecs As Single

    lMaxSecs = 7200

    z_File_Note pNote, True
    
    pRetTimeP = pRetTime
    pRetTime = pNote.abs_time
    
    z_LoadSetLim pRetTime, lMaxSecs, pRetTimeP
    
    pNote.abs_time = pRetTime
    
    z_LoadSetLim pNote.sNote, 120 'note index limit
    pNote.dFreq = GetFreq(pNote.sNote)
    
End Sub
Private Sub z_LoadSetLim(pVar As Variant, ByVal pHigh As Single, Optional ByVal pLow As Single = 0)
    If pVar > pHigh Then
        pVar = pHigh
    ElseIf pVar < pLow Then
        pVar = pLow
    End If
End Sub
Private Sub z_LoadLngPosSngD(pRetTimeD As Long, ByVal pSngTimeD As Single)
    pRetTimeD = Int(pSngTimeD * PCM.lSamplesPerSec + 0.5) 'global PCM from module mAudioBASE
End Sub

'*------------------------*'
'-                        -'
'-  Nested Save           -'
'-                        -'
'*------------------------*'


Private Sub z_Save_MakeRefSequential(pDest As NoteEvent, pL1 As Long, pElem As Long, pSourceAry() As NoteEvent)
    
    If pSourceAry(pElem).Event = NOTE_ON Then
        mRef = pSourceAry(pElem).Ref
        mNoteE(mRef).Ref = pL1
    ElseIf pSourceAry(pElem).Event = NOTE_OFF Then
        mRef = pSourceAry(pElem).Ref
        mNoteE(mRef).Ref = pL1
    End If

    pDest.abs_time = pSourceAry(pElem).abs_time
    pDest.Event = pSourceAry(pElem).Event
    pDest.sNote = pSourceAry(pElem).sNote

End Sub


'*------------------------*'
'-                        -'
'-  Common  Load / Save   -'
'-                        -'
'*------------------------*'

Public Function DialogSuccess(pFileSpec As String, Optional ByVal pIsLoad As Boolean = True, Optional pRetFileName As String, Optional pRetDir As String = "", Optional ByVal pForceExtension As String = "") As Boolean
DialogSuccess = False

    'Experimental

    If mDialogShowing Then Exit Function
    
    mStrFile = pFileSpec

    If Left$(pForceExtension, 1) <> "." Then pForceExtension = "." & pForceExtension
    
    Set CDLF = New OSDialog
    
    For IA = Len(mStrFile) To 1 Step -1
        If Mid$(mStrFile, IA, 1) = "\" Then
            CDLF.Directory = Left$(mStrFile, IA - 1)
            Exit For
        End If
    Next
    
    mDialogShowing = True
    
    mStr = "*" & pForceExtension
    
    If pIsLoad Then
        CDLF.ShowOpen mStrFile, , "(" & mStr & ")|" & mStr, pRetDir
        mDialogShowing = False
        If Not IsFile(mStrFile) Then
            Set CDLF = Nothing
            Exit Function
        End If
    Else
        If CDLF.ShowSave(mStrFile, , "(" & mStr & ")|" & mStr, pRetDir, pForceExtension) = "" Then
            mDialogShowing = False
            Set CDLF = Nothing
            Exit Function
        Else
            mDialogShowing = False
        End If
    End If
    
    For IA = 1 To Len(mStrFile)
        If Mid$(mStrFile, IA, 1) = "." Then
            If Len(pForceExtension) > 1 Then
                mStrFile = Left$(mStrFile, IA - 1) & pForceExtension
            Else
                pForceExtension = Right$(mStrFile, Len(mStrFile) - IA + 1)
            End If
            Exit For
        End If
    Next
    
    If Right$(mStrFile, Len(pForceExtension)) <> pForceExtension Then
        mStrFile = mStrFile & pForceExtension
    End If
    
    If Len(mStrFile) > Len(pForceExtension) Then
    
        pRetFileName = mStrFile
        pRetDir = CDLF.Directory
        pFileSpec = pRetDir & mStrFile
                
        mFreeFile = FreeFile

        DialogSuccess = True
        
    End If
    
    Set CDLF = Nothing

End Function
Private Sub z_File_Note(pNote As NoteEvent, Optional ByVal pIsLoading As Boolean)
    
    If pIsLoading Then
        Get #mFreeFile, , pNote.abs_time
        Get #mFreeFile, , pNote.Event
        Get #mFreeFile, , pNote.Ref
        Get #mFreeFile, , pNote.sNote
    Else
        Put #mFreeFile, , pNote.abs_time
        Put #mFreeFile, , pNote.Event
        Put #mFreeFile, , pNote.Ref
        Put #mFreeFile, , pNote.sNote
    End If

End Sub


'*------------------*'
'-                  -'
'-    Midi Save     -'
'-                  -'
'*------------------*'


Public Function SaveMidiFile(pSong As Song) As Boolean
Dim Str1 As String
Dim I1&, DeltaByteLen As Long
Dim Microsecs_QN As Long
Dim Beats_Min As Long
Dim GetNoteEventCount As Long
Dim cMidiBytes As Long

    SaveMidiFile = DialogSuccess(mStrMidiFile, False, mStrMidiFile, gMidiDir, ".mid")
    
    If Not SaveMidiFile Then Exit Function
    
    Open mStrMidiFile For Output As #1
        Print #1,
    Close #1
    
    Open mStrMidiFile For Binary As #1
    
    Put #1, 1, "MThd" 'midi header
    For I1 = 1 To 3
        Put #1, , CByte(0)
    Next
    Put #1, , CByte(6) '6 bytes for next record
    z_SaveMidi_WriteFixedLen 1, 2  'Midi Type 1 = multi tracks, sep channels
    z_SaveMidi_WriteFixedLen pSong.Proc.cTrack + 1, 2
    
    PPQN = 240 'pulses per quarter note
    z_SaveMidi_WriteFixedLen PPQN, 2
    
    'Creating a track that sets song information like tempo
    Put #1, , "MTrk"     'Each track header
    
    'Track length in bytes
    Put #1, , CByte(0)
    z_SaveMidi_WriteFixedLen 10, 3
    
    'META Key Signature - These are first 6 bytes of track length
    z_SaveMidi_WriteVariLen 0  'Delta Time Stamp
    Put #1, , CByte(&HFF)
    Put #1, , CByte(&H59)
    z_SaveMidi_WriteVariLen 2
    Put #1, , CByte(0) '0, -1 to -7 for number of flats, 1 to 7 for sharps
    Put #1, , CByte(0) '0 = major, 1 = minor
    
    z_SaveMidi_WriteVariLen 0 'Delta Time Stamp
    Put #1, , CByte(&HFF) 'META End of track
    Put #1, , CByte(&H2F)
    Put #1, , CByte(0)
    
    For IA = 1 To pSong.Proc.cTrack
        Put #1, , "MTrk"
        Put #1, , CByte(0) 'First byte
        
        DeltaByteLen = 0
        
        'obtain DeltaByteLen first
        z_SaveMidi_WriteNoteEvents pSong.Track(IA).OnOff.Seq, DeltaByteLen, GetNoteEventCount
        cMidiBytes = 4
        
        'Note On and Note Off events use 3 bytes each
        
        z_SaveMidi_WriteFixedLen GetNoteEventCount * 3 + cMidiBytes + DeltaByteLen, 3
        z_SaveMidi_WriteNoteEvents pSong.Track(IA).OnOff.Seq
        
    Next
    
    Close #1

End Function
Private Sub z_SaveMidi_WriteNoteEvents(pSeq As SequencerElement, Optional RetDeltaByteLen As Long = -1, Optional RetNoteEventCount As Long)
Dim Elem1&, Elem2&, J1&
Dim speed_mult!
Dim ByteLen1 As Long
Dim sNoteVal As Single
Dim sNoteVel As Single
Dim DigiTime As Long
Dim DigiDelta As Long
Dim sngTime As Single
Dim NoteSuccess As Boolean
Dim KeyPressed(127) As Boolean, lElem As Integer

    speed_mult = PPQN * 2
    
    RetNoteEventCount = 0
    
    lElem = pSeq.Events.LinkHead
    
    For J1 = 1 To pSeq.Events.StackPtr
    
        sngTime = pSeq.NINFO(lElem).abs_time * speed_mult
        DigiDelta = Int(sngTime - DigiTime + 0.5)
        If DigiDelta < 0 Then DigiDelta = 0
        DigiTime = DigiTime + DigiDelta
            
        If pSeq.NINFO(lElem).Event = NOTE_ON Then
            Elem2 = z_SaveMidi_IndexLimit(pSeq.NINFO(lElem).sNote)
            If Not KeyPressed(Elem2) Then
                If RetDeltaByteLen = -1 Then
                    'All Midi Messages begin with a Delta Time
                    z_SaveMidi_WriteVariLen DigiDelta
                    Put #1, , CByte(NOTE_ON) 'Note on, midi channel 1
                End If
                Add RetNoteEventCount, 1
                KeyPressed(Elem2) = True
                NoteSuccess = True
            End If
        ElseIf pSeq.NINFO(lElem).Event = NOTE_OFF Then
            Elem2 = z_SaveMidi_IndexLimit(pSeq.NINFO(lElem).sNote)
            If KeyPressed(Elem2) Then
                If RetDeltaByteLen = -1 Then
                    z_SaveMidi_WriteVariLen DigiDelta
                    Put #1, , CByte(NOTE_OFF)
                End If
                Add RetNoteEventCount, 1
                KeyPressed(Elem2) = False
                NoteSuccess = True
            End If
        End If
            
        If NoteSuccess Then
            sNoteVel = 86 ' pseq.NINFO(lElem).sVeloc / 2 '+ 96
            If sNoteVel > 126 Then sNoteVel = 126
            If sNoteVel < 0 Then sNoteVel = 0
            If RetDeltaByteLen = -1 Then
                'takes care of rounding errors
                Put #1, , CByte(z_SaveMidi_IndexLimit(pSeq.NINFO(lElem).sNote))
                Put #1, , CByte(Int(sNoteVel) + 0.5)
            Else
                z_SaveMidi_WriteVariLen DigiDelta, ByteLen1
                RetDeltaByteLen = RetDeltaByteLen + ByteLen1
            End If
            NoteSuccess = False
        End If 'NoteSuccess
        
        lElem = pSeq.Events.Link(lElem).Next
        
    Next
    
    If RetDeltaByteLen = -1 Then
        Put #1, , CByte(0)
        Put #1, , CByte(&HFF)
        Put #1, , CByte(&H2F)
        Put #1, , CByte(&H0)
    End If
    
End Sub
Private Function z_SaveMidi_ReadVarLen(ByVal InByte As Long) As Long
Dim I1&

    I1 = InByte * 128
    I1 = I1 Or (InByte And &H7F&)

End Function
Private Sub z_SaveMidi_WriteVariLen(ByVal InputVal As Long, Optional RetByteLen As Long = -1) ' Optional Max65535 As Boolean = False, Optional ForceWord As Boolean)
Dim ByteSim As Long
Dim UpperBound As Long
Dim Found1stChunk As Boolean
Dim MaskBits As Long

    UpperBound = 128 ^ 4
    MaskBits = 127 * (128 ^ 3) ' 1111111 followed by 21 zeros

    If RetByteLen > 0 Then RetByteLen = 0

    Do While UpperBound > 1
        ByteSim = InputVal And MaskBits
        UpperBound = UpperBound / 128
        ByteSim = ByteSim / UpperBound
        If ByteSim > 0 Or Found1stChunk Or UpperBound = 1 Then
            If UpperBound > 1 Then
                ByteSim = ByteSim Or &H80
            End If
            If RetByteLen = -1 Then
                Put #1, , CByte(ByteSim)
            Else
                RetByteLen = RetByteLen + 1
            End If
            Found1stChunk = True
        End If
        MaskBits = MaskBits / 128
    Loop
    
End Sub
Private Sub z_SaveMidi_WriteFixedLen(ByVal InputVal As Long, Optional MaxBytes As Long = 2)
Dim ByteSim As Long
Dim UpperBound As Long
Dim Found1stChunk As Boolean
Dim MaskBits As Long

    If MaxBytes > 3 Then MaxBytes = 3
    
    UpperBound = 256 ^ (MaxBytes)
    MaskBits = 255 * (256 ^ (MaxBytes - 1)) ' 1111111 followed by 21 zeros

    Do While UpperBound > 1
        ByteSim = InputVal And MaskBits
        UpperBound = UpperBound / 256
        ByteSim = ByteSim / UpperBound
        Put #1, , CByte(ByteSim)
        MaskBits = MaskBits / 256
    Loop
    
End Sub
Private Function z_SaveMidi_IndexLimit(ByVal sNoteVal As Single) As Long
    If sNoteVal < 0 Then sNoteVal = 0
    If sNoteVal > 126 Then sNoteVal = 126
    z_SaveMidi_IndexLimit = Int(sNoteVal + 0.5)
End Function


Public Sub DirectoryInfo(Optional pSave As Boolean = False)

    If Right$(App.Path, 1) = "\" Then
        mStrFile = App.Path & "settings.txt"
    Else
        mStrFile = App.Path & "\settings.txt"
    End If
    
    If pSave Then
        Open mStrFile For Output As #1
            Write #1,
        Close #1
        Open mStrFile For Binary As #1
            
            Seek #1, 1
            
            mStr = "SongDir " & Chr$(ASC_DOUBLE_QUOTE) & gSongDir & Chr$(ASC_DOUBLE_QUOTE) & vbCrLf
            Erase mBytes
            ReDim mBytes(Len(mStr) - 1)
            FillBytesFromString mBytes, mStr
            Put #1, , mBytes
            
            mStr = "MidiDir " & Chr$(ASC_DOUBLE_QUOTE) & gMidiDir & Chr$(ASC_DOUBLE_QUOTE)
            Erase mBytes
            ReDim mBytes(Len(mStr) - 1)
            FillBytesFromString mBytes, mStr
            Put #1, , mBytes
        
        Close #1
    Else
        If Not IsFile(mStrFile) Then Exit Sub
        Open mStrFile For Input As #1
            If Not EOF(1) Then
                mLA = 1
                Line Input #1, mStr
                z_GetNextDoubleQuote mStr, mLA, mL1
                z_GetNextDoubleQuote mStr, mLA, mL2, gSongDir
            End If
            If Not EOF(1) Then
                mLA = 1
                Line Input #1, mStr
                z_GetNextDoubleQuote mStr, mLA, mL1
                z_GetNextDoubleQuote mStr, mLA, mL2, gMidiDir
            End If
        Close #1
    End If
    
End Sub
Private Sub z_GetNextDoubleQuote(pStr As String, pPos As Long, pRetPos As Long, Optional pRetString As Variant)
    
    For IA = pPos To Len(mStr) 'IA seen by all mSSaveLoad.bas subs
        If Asc(Mid$(mStr, IA, 1)) = ASC_DOUBLE_QUOTE Then
            If Not IsMissing(pRetString) Then
                pRetString = Mid$(mStr, pPos, IA - pPos)
            End If
            pRetPos = IA
            pPos = IA + 1
            Exit For
        End If
    Next
    
End Sub

