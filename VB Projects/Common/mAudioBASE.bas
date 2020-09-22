Attribute VB_Name = "mAudioBASE"
Option Explicit

'Experimental audio base for VB6

'Dependency: mGeneral.bas

'Goals:  Test, in this order, whether FMOD, SDL, et. al.
'will work

'Module intended as a synth template which will eventually
'work with callbacks

Public LatencyAdjust&       'Experimental

Public SynthAPI_flags As Long    '

Public Const API_FMOD    As Long = 1
Private Const FSOUND_CDQuality_Signed As Long = FSOUND_16BITS Or FSOUND_STEREO
Private Const FS_FORMAT As Long = FSOUND_16BITS Or FSOUND_STEREO Or FSOUND_LOOP_NORMAL
Dim DSP_Unit    As Long

Public Const API_SDL     As Long = 2
Public Const API_BASS    As Long = 4
Public Const API_DirectX As Long = 8

Dim hChn        As Long
Dim SrcPtr      As Long
Dim DestPtr     As Long
Dim BufFormat   As Long
Dim Ptr_A       As Long
Dim Ptr_B       As Long
Dim Len_A       As Long
Dim Len_B       As Long


Dim SA      As SAFEARRAY1D
Dim mUseDSP As Boolean

'Helper Vars - future sound technology may use > 16 bits per sample.
'If that is the case ..
Public Const L_32767 As Long = 32767
Public Const L_Neg_32768 As Long = -32768

Public Const L_16383 As Long = 16383

Private Type WAVEFORMATEX
    nFormatTag      As Integer
    nChannels       As Integer
    lSamplesPerSec  As Long
    lAvgBytesPerSec As Long
    nBlockAlign     As Integer
    nBitsPerSample  As Integer
    cbSize          As Integer
End Type

Public PCM  As WAVEFORMATEX

' ======== softsynth =========

Public sRender()     As Single          'float array for main synth
Public lpBuffer() As Integer            '16 bits per sample final

Public mainVol       As Single          'experimental
Public SampleFreq    As Long
Public BufferSamples As Long
Public BufUBound     As Long

Public CPos          As Long            'sample position
Public PosPrevC      As Long
Public BufStart      As Long
Public BufStop       As Long
Public SamStart      As Long
Public SamStop       As Long
Public CalcStart     As Long
Public gSamPosP      As Long 'added 2006/03M/30

Public idelayPos      As Single              'anti-aliased delay
Public sDelayAmount   As Single
Public sDelayVolume   As Single
Private DelaybufferL()  As Integer
Private DelaybufferR()  As Integer
Private DelaybufferL2() As Integer
Private DelaybufferR2() As Integer
Private DBufUB          As Long
Private DBufUB2         As Long
Private DelaySamples    As Long
Private DelaySamples2   As Long
Private delayPosL       As Single
Private delayPosR       As Single
Private delayPosL2      As Single
Private delayPosR2      As Single
Private MicroDelay      As Long
Private MicroDelay2     As Long
Public SamWriteLength   As Long

Private m_triangle_freq As Double
Private m_wavetable_freq As Double

Public sSampleTime As Double

Public Const QWERTY_C As Double = 261.625565 / 16

Public Const QwertyNoteValueLB As Long = 0
Private Const QwertyNoteValueUB As Long = QwertyNoteValueLB + 108

Public NoteIndexFromQwerty(255) As Byte
Public Sub SoundPos(Optional RetStart&, Optional RetStop&, Optional Paused As Boolean)

    If (SynthAPI_flags And API_FMOD) > 0 Then
    
    '   SECOND PART of latency experiment .. doesn't really work
    '    CPos = CPos - BufferSamples * Int(CPos / BufferSamples)
     
        If mUseDSP Then
            PosPrevC = -1
            CPos = 0
    '        ZSynth_callback
    '        ZSynth_WriteNotesToSignal BufferSamples - 1, Paused
        Else
            PosPrevC = CPos
            CPos = FSOUND_GetCurrentPosition(hChn) '+ 1024  'latency experiment
            If PosPrevC < CPos Then
                ZSynth_WriteNotesToSignal CPos - 1, Paused
            ElseIf PosPrevC > CPos Then
                ZSynth_WriteNotesToSignal BufferSamples - 1, Paused
            End If
        End If
    
    ElseIf (SynthAPI_flags And API_DirectX) > 0 Then
    
    End If
    
    RetStart = BufStart
    
    If PosPrevC = CPos Then
        'create a skip-over condition to bypass synth algorithms.
        RetStop = BufStart - PCM.nChannels
    Else
        RetStop = BufStop
    End If
    
    SamWriteLength = (RetStop - RetStart + PCM.nChannels) / PCM.nChannels
 
End Sub
Public Function WriteBuf() As Long
Dim Len_&

    If SynthAPI_flags = 0 Then Exit Function

    If PosPrevC < CPos Then
        Len_ = CPos - CalcStart
'        FSOUND_Sample_Lock DestPtr, CalcStart, Len_, Ptr_A, Ptr_B, Len_A, Len_B
        ZSynth_WriteBuf Len_
'        FSOUND_Sample_Unlock DestPtr, Ptr_A, Ptr_B, Len_A, Len_B
        CalcStart = CPos
        WriteBuf = 1
    ElseIf PosPrevC > CPos Then
        Len_ = BufferSamples - CalcStart
'        FSOUND_Sample_Lock DestPtr, CalcStart, Len_, Ptr_A, Ptr_B, Len_A, Len_B
        ZSynth_WriteBuf Len_
'        FSOUND_Sample_Unlock DestPtr, Ptr_A, Ptr_B, Len_A, Len_B
        CalcStart = 0
        WriteBuf = 1
    Else
        WriteBuf = 0
    End If
    
End Function

Public Sub Init_Synth(Optional hwnd&, Optional BufSamples As Long = 4096, Optional SamplingFrequency& = 44100, Optional volume As Single = 0.2, Optional ByVal bUseDSP As Boolean = False, Optional ByVal pRequest_API As Long = API_FMOD)
    
    On Local Error GoTo OHNO
    
    Test_InitQWERTY NoteIndexFromQwerty

    '''''''''''''''''''''''''''
    ' experimental sub        '
    '''''''''''''''''''''''''''
    
    'set environment variables
    ZSynth_WaveProps SamplingFrequency, volume, BufSamples
    
    m_TryFMOD bUseDSP
    
    'high resolution synth buffer
    ReDim sRender(BufUBound)
    
OHNO:

    If SynthAPI_flags = 0 Then
        If pRequest_API And API_FMOD Then
'            MsgBox "Real-time audio requires fmod.dll to be placed in C:\Windows\System" & vbCrLf & vbCrLf & "www.fmod.org", , "Sound Init:"
        End If
        If pRequest_API And API_DirectX Then
            MsgBox "DirectX requested but not found"
        End If
    End If

End Sub
Private Sub m_TryFMOD(bUseDSP As Boolean)

    'found this value using FSound_Sample_GetFormat
    'on buffer created from cd quality wav file
    BufFormat = 338
    
    'init the sound environment!
    FSOUND_Init SampleFreq, 16, 0
    
    'specialized environment variable
    BufUBound = BufferSamples * PCM.nChannels - 1
    
    'vb array to point to previous buffer
    ReDim lpBuffer(BufUBound)
    SrcPtr = VarPtr(lpBuffer(0))
    
    mUseDSP = bUseDSP
    
    If bUseDSP Then
        DSP_Unit = FSOUND_DSP_Create(AddressOf ZSynth_callback, FSOUND_DSP_DEFAULTPRIORITY_USER, 0)
        FSOUND_DSP_SetActive DSP_Unit, True
    Else
        'create a buffer for playback
        DestPtr = FSOUND_Sample_Alloc(1, BufferSamples, BufFormat, SampleFreq, 255, FSOUND_STEREOPAN, 0)
        'start the sound engine
        hChn = FSOUND_PlaySound(FSOUND_FREE, DestPtr)
        
        'vb way of filling out pointer information
        FSOUND_Sample_Lock DestPtr, 0, BufferSamples, Ptr_A, Ptr_B, Len_A, Len_B
        SA.cbElements = PCM.nBitsPerSample / 8
        SA.cElements = BufferSamples * PCM.nChannels
        SA.cDims = 1
        SA.pvData = Ptr_A
        If Ptr_A <> 0 Then
            'clear previous memory reference
            CopyMemory ByVal VarPtrArray(lpBuffer), 0&, 4
            
            'create new reference
            CopyMemory ByVal VarPtrArray(lpBuffer), VarPtr(SA), 4&
            
        End If
        FSOUND_Sample_Unlock DestPtr, Ptr_A, Ptr_B, Len_A, Len_B
    End If
    
    SynthAPI_flags = SynthAPI_flags Or API_FMOD
        
End Sub
Public Function ZSynth_callback(ByVal buffer As Long, ByVal newbuffer As Long, ByVal length As Long, ByVal param As Long) As Long
    
    SA.cbElements = PCM.nBitsPerSample / 8
    SA.cElements = length * PCM.nChannels
    SA.cDims = 1
    SA.pvData = buffer
    
    'erase old pointer
    CopyMemory ByVal VarPtrArray(lpBuffer), 0&, 4

    CopyMemory ByVal VarPtrArray(lpBuffer), SA, 4
    
    ZSynth_WriteNotesToSignal length
    
'    ZSynth_WriteBuf length

End Function

Public Sub ZSynth_WriteNotesToSignal(Optional ByVal CalcStop1&, Optional PauseSignal As Boolean)
Dim I1&

    'ripped from a working synth project

    BufStart = CalcStart * PCM.nChannels
    BufStop = CalcStop1 * PCM.nChannels
    
    SamStart = CalcStart
    SamStop = CalcStop1
    
    m_ZeroSignal 'Reset the high resolution sound buffer
    
    If Not PauseSignal Then
        
        'AddController QwertyController, BufStart, BufStop
        
    End If
        
End Sub

Public Sub ZSynth_WriteBuf(length&)
Dim levL!
Dim bufLvl!
Dim I1&
Dim BufStop1&
Dim DelayPos1&
Dim DelayPos2&
Dim DelayVal1&
Dim DelayVal2&
Dim LDelayPosAhead&
Dim LDelayPosAhead2&
Dim Pos1&
Dim Pos2&
Dim Pos3&
Dim Pos4&
Dim Pos5&

    BufStop1 = BufStart + length * PCM.nChannels - 1
 
    If DBufUB > 0 Then 'delay buffer has been initialized
 
       'Rounding errors mess up the delay processor when increment
       'uses decimal point.  This code section keeps the sound quality respectful.
       'If there were no rounding errors, these should be placed inside WriteBuf()
       If delayPosR = Int(delayPosR) Then
           delayPosR = delayPosR - DelaySamples * Int(delayPosR / DelaySamples)
       End If
       
       If delayPosL = Int(delayPosL) Then
           delayPosL = delayPosL - DelaySamples * Int(delayPosL / DelaySamples)
       End If
       
       For I1 = BufStart To BufStop1 Step PCM.nChannels
     
             'linear interpolation
           DelayPos1 = Int(delayPosL)
           DelayPos2 = DelayPos1 + 1
           levL = delayPosL - DelayPos1
           LDelayPosAhead = DelayPos1 + MicroDelay
           DelayPos1 = DelayPos1 - DelaySamples * Int(DelayPos1 / DelaySamples)
           DelayPos2 = DelayPos2 - DelaySamples * Int(DelayPos2 / DelaySamples)
           DelayVal1 = DelaybufferL(DelayPos1)
           DelayVal2 = DelaybufferL(DelayPos2)
             
           bufLvl = (sRender(I1) * 32767& + sDelayAmount * (DelayVal1 + levL * (DelayVal2 - DelayVal1))) * mainVol
             
           If bufLvl > L_32767 Then
               bufLvl = L_32767
           ElseIf bufLvl < L_Neg_32768 Then
               bufLvl = L_Neg_32768
           End If
           
           lpBuffer(I1) = Int(bufLvl + 0.5)
             
           LDelayPosAhead = LDelayPosAhead - DelaySamples * Int(LDelayPosAhead / DelaySamples)
           DelayVal DelaybufferL(LDelayPosAhead), bufLvl, sDelayVolume
           DelaybufferL(LDelayPosAhead) = bufLvl * sDelayVolume
             
           delayPosL = delayPosL + idelayPos
     
       Next
    
       If PCM.nChannels = 2 Then
       
           For I1 = BufStart + 1 To BufStop1 Step PCM.nChannels
     
               'linear interpolation
               DelayPos1 = Int(delayPosR)
               DelayPos2 = DelayPos1 + 1
               levL = delayPosR - DelayPos1
               LDelayPosAhead = DelayPos1 + MicroDelay
               DelayPos1 = DelayPos1 - DelaySamples * Int(DelayPos1 / DelaySamples)
               DelayPos2 = DelayPos2 - DelaySamples * Int(DelayPos2 / DelaySamples)
               DelayVal1 = DelaybufferR(DelayPos1)
               DelayVal2 = DelaybufferR(DelayPos2)
                 
               bufLvl = (sRender(I1) * 32767& + sDelayAmount * (DelayVal1 + levL * (DelayVal2 - DelayVal1))) * mainVol
                 
               If bufLvl > L_32767 Then
                   bufLvl = L_32767
               ElseIf bufLvl < L_Neg_32768 Then
                   bufLvl = L_Neg_32768
               End If
                 
               lpBuffer(I1) = Int(bufLvl + 0.5)
                
               LDelayPosAhead = LDelayPosAhead - DelaySamples * Int(LDelayPosAhead / DelaySamples)
               DelayVal DelaybufferR(LDelayPosAhead), bufLvl, sDelayVolume
               DelaybufferR(LDelayPosAhead) = bufLvl * sDelayVolume
                 
               delayPosR = delayPosR + idelayPos
             
           Next
           
       End If 'PCM.nChannels = 2
 
    Else 'delay buffer has not been initialized
    
        levL = mainVol * 32767
 
        'If stereo, array pos of BufStart
        'is always even,
        'BufStop1 is always odd
        
        '| 4 | 5 | 6 | 7 | 8 | 9 |
        '| L | R | L | R | L | R |
 
        For I1 = BufStart To BufStop1
  
            bufLvl = sRender(I1) * levL
              
            If bufLvl > L_32767 Then
                bufLvl = L_32767
            ElseIf bufLvl < L_Neg_32768 Then
                bufLvl = L_Neg_32768
            End If
            lpBuffer(I1) = Int(bufLvl + 0.5)
  
        Next
 
    End If
 
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''
' These subs are "standard operating procedure"      '
'                                                    '
''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ZSynth_WaveProps(Optional SamplingFrequency& = 44100, Optional volume As Single = 1, Optional BufSamples As Long = 2048)

    SampleFreq = SamplingFrequency
    
    PCM.nFormatTag = 1 'Always WAVE_FORMAT_PCM
    PCM.nChannels = 2
    PCM.lSamplesPerSec = SampleFreq
    PCM.nBitsPerSample = 16
    PCM.nBlockAlign = PCM.nChannels * PCM.nBitsPerSample / 8
    PCM.lAvgBytesPerSec = PCM.lSamplesPerSec * PCM.nBlockAlign
    
    sSampleTime = PCM.nBitsPerSample * SampleFreq

    BufferSamples = BufSamples

    mainVol = volume
    
End Sub
Public Sub ZSynth_InitSynth(Optional freq_multiplier! = 1, Optional OscPerNote As Byte = 1, Optional SetPoly As Byte = 16, Optional AverageGranuleSize As Long = 22050)

     ' === custom software synthesizer vars
    m_triangle_freq = freq_multiplier * 8.20289 / SampleFreq
    m_wavetable_freq = freq_multiplier / 2 ^ 5 / SampleFreq
    
End Sub

Private Sub DelayVal(val1 As Integer, sVal2 As Single, sVal1 As Single)
    val1 = val1 + sVal1 * (sVal2 - val1)
End Sub
Private Sub m_ZeroSignal()
Dim I1&
    
    For I1 = BufStart To BufStop + PCM.nChannels - 1
        sRender(I1) = 0
    Next

End Sub


Public Sub StopSound()
'    If ZSynthAPI = API_FMOD Then
        Stop_Fmod
'    ElseIf ZSynthAPI = API_SDL Then
        Stop_SDL_Audio
'    ElseIf ZSynthAPI = API_DirectX Then
        Stop_DirectSound
'    End If
End Sub

Public Sub SilenceBuffer()
Dim I1&
    If SynthAPI_flags = 0 Then Exit Sub
    For I1 = 0 To BufUBound
        lpBuffer(I1) = 0
    Next
End Sub

Public Sub Stop_Fmod()

    If (SynthAPI_flags And API_FMOD) = 0 Then Exit Sub
    CopyMemory ByVal VarPtrArray(lpBuffer), 0&, 4
    
'    FSOUND_Sample_Unlock DestPtr, Ptr_A, Ptr_B, Len_A, Len_B
    FSOUND_DSP_Free DSP_Unit
    FSOUND_Close
    
End Sub
Public Sub Stop_DirectSound()
    
    'dummy sub

End Sub
Public Sub Stop_SDL_Audio()
    
    '

End Sub

'Audio 'Delay' calculations in WriteBuf()
Public Sub ResizeDelay(Samples As Long, Optional ByVal sAmount_ As Single = 26000 / 32767, Optional ByVal sVolume_ As Single = 32000 / 32767, Optional ByVal iPos_ As Single = 1)
    If Samples > 0 Then
        DelaySamples = Samples
        DBufUB = DelaySamples - 1
        Erase DelaybufferL
        ReDim DelaybufferL(DBufUB)
        Erase DelaybufferR
        ReDim DelaybufferR(DBufUB)
        delayPosL = 0
        delayPosR = 0
        MicroDelay = DBufUB
        sDelayAmount = sAmount_
        sDelayVolume = sVolume_
        idelayPos = iPos_
    End If
End Sub
Public Sub ResizeDelay2(Samples As Long)
    If Samples > 0 Then
        DelaySamples2 = Samples
        DBufUB2 = DelaySamples2 - 1
        Erase DelaybufferL2
        ReDim DelaybufferL2(DBufUB2)
        Erase DelaybufferR2
        ReDim DelaybufferR2(DBufUB2)
        delayPosL2 = 0
        delayPosR2 = 0
        MicroDelay2 = DBufUB2
    End If
End Sub


' -------- MATHS ---------

Private Sub m_SwapIt(Variana1, Variana2)
Dim Variana3!
    Variana3 = Variana1
    Variana1 = Variana2
    Variana2 = Variana3
End Sub
Private Sub m_IncLngAndRecord(Lng1&, Rec1&)
    Lng1 = Lng1 + 1
    Rec1 = Lng1
End Sub
Private Sub m_PullUp(Contender As Variant, Barlevel As Variant)
    If Contender < Barlevel Then Contender = Barlevel
End Sub


' -------- OTHER ---------

Public Sub Test_InitQWERTY(QwertyRefs() As Byte, Optional ByVal DoFullRange As Boolean)
 
    AssignNoteVal QwertyRefs, vbKeyZ, 4, 0
    AssignNoteVal QwertyRefs, vbKeyS, 4, 1
    AssignNoteVal QwertyRefs, vbKeyX, 4, 2
    AssignNoteVal QwertyRefs, vbKeyD, 4, 3
    AssignNoteVal QwertyRefs, vbKeyC, 4, 4
    AssignNoteVal QwertyRefs, vbKeyV, 4, 5
    AssignNoteVal QwertyRefs, vbKeyG, 4, 6
    AssignNoteVal QwertyRefs, vbKeyB, 4, 7
    AssignNoteVal QwertyRefs, vbKeyH, 4, 8
    AssignNoteVal QwertyRefs, vbKeyN, 4, 9
    AssignNoteVal QwertyRefs, vbKeyJ, 4, 10
    AssignNoteVal QwertyRefs, vbKeyM, 4, 11
    
    AssignNoteVal QwertyRefs, 188, 5, 0   ' <
    AssignNoteVal QwertyRefs, vbKeyL, 5, 1 ' L
    AssignNoteVal QwertyRefs, 190, 5, 2   ' >
    AssignNoteVal QwertyRefs, 186, 5, 3   ' semicolon
    AssignNoteVal QwertyRefs, 191, 5, 4   ' ?
 
    If DoFullRange Then
    
        AssignNoteVal QwertyRefs, vbKeyQ, 5, 5
        AssignNoteVal QwertyRefs, vbKey2, 5, 6
        AssignNoteVal QwertyRefs, vbKeyW, 5, 7
        AssignNoteVal QwertyRefs, vbKey3, 5, 8
        AssignNoteVal QwertyRefs, vbKeyE, 5, 9
        AssignNoteVal QwertyRefs, vbKey4, 5, 10
        AssignNoteVal QwertyRefs, vbKeyR, 5, 11
        
        AssignNoteVal QwertyRefs, vbKeyT, 6, 0
        AssignNoteVal QwertyRefs, vbKey6, 6, 1
        AssignNoteVal QwertyRefs, vbKeyY, 6, 2
        AssignNoteVal QwertyRefs, vbKey7, 6, 3
        AssignNoteVal QwertyRefs, vbKeyU, 6, 4
        
        AssignNoteVal QwertyRefs, vbKeyI, 6, 5
        AssignNoteVal QwertyRefs, vbKey9, 6, 6
        AssignNoteVal QwertyRefs, vbKeyO, 6, 7
        AssignNoteVal QwertyRefs, vbKey0, 6, 8
        AssignNoteVal QwertyRefs, vbKeyP, 6, 9
        AssignNoteVal QwertyRefs, 189, 6, 10 ' -
        AssignNoteVal QwertyRefs, 219, 6, 11 ' [
        AssignNoteVal QwertyRefs, 221, 7, 0 ' ]
        AssignNoteVal QwertyRefs, 220, 7, 1 ' \
        AssignNoteVal QwertyRefs, 13, 7, 2 'Enter
        AssignNoteVal QwertyRefs, 8, 7, 3  'backspace
    
    Else
    
        AssignNoteVal QwertyRefs, vbKeyQ, 5, 0
        AssignNoteVal QwertyRefs, vbKey2, 5, 1
        AssignNoteVal QwertyRefs, vbKeyW, 5, 2
        AssignNoteVal QwertyRefs, vbKey3, 5, 3
        AssignNoteVal QwertyRefs, vbKeyE, 5, 4
        AssignNoteVal QwertyRefs, vbKeyR, 5, 5
        AssignNoteVal QwertyRefs, vbKey5, 5, 6
        AssignNoteVal QwertyRefs, vbKeyT, 5, 7
        AssignNoteVal QwertyRefs, vbKey6, 5, 8
        AssignNoteVal QwertyRefs, vbKeyY, 5, 9
        AssignNoteVal QwertyRefs, vbKey7, 5, 10
        AssignNoteVal QwertyRefs, vbKeyU, 5, 11
        
        AssignNoteVal QwertyRefs, vbKeyI, 6, 0
        AssignNoteVal QwertyRefs, vbKey9, 6, 1
        AssignNoteVal QwertyRefs, vbKeyO, 6, 2
        AssignNoteVal QwertyRefs, vbKey0, 6, 3
        AssignNoteVal QwertyRefs, vbKeyP, 6, 4
        AssignNoteVal QwertyRefs, 219, 6, 5 ' [
        AssignNoteVal QwertyRefs, 187, 6, 6 ' + =
        AssignNoteVal QwertyRefs, 221, 6, 7 ' ]
        AssignNoteVal QwertyRefs, 220, 6, 8 ' \
        AssignNoteVal QwertyRefs, 13, 6, 9 'Enter
        AssignNoteVal QwertyRefs, 8, 6, 10 'backspace
    
    End If
 
End Sub
Private Sub AssignNoteVal(QwertyRefs() As Byte, KeyIndex As Byte, Octave_ As Byte, PianoKeyOffset As Byte)
    QwertyRefs(KeyIndex) = Octave_ * 12 + PianoKeyOffset + QwertyNoteValueLB
End Sub

Public Function GetFreq(ByVal pNoteIndex As Long, Optional ByVal pToneBase As Long = 12) As Double
    GetFreq = 0
    If pToneBase < 1 Or SampleFreq < 1 Then Exit Function
    GetFreq = QWERTY_C * 2 ^ (pNoteIndex / pToneBase) / SampleFreq
End Function

