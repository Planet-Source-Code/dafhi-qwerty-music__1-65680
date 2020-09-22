VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsbMinPoly 
      Height          =   2055
      Left            =   4320
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
      Begin VB.CheckBox chkNewEvery 
         Caption         =   "Auto new track"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.VScrollBar vsbTrack 
         Height          =   1095
         Left            =   1920
         Max             =   1
         Min             =   1
         TabIndex        =   3
         Top             =   120
         Value           =   1
         Width           =   255
      End
      Begin VB.Label lblqRec 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "paused"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblTrack 
         Alignment       =   2  'Center
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   15.75
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblqTrack 
         Caption         =   "Track"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.Label lblPoly 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblqPoly 
      Caption         =   "Polyphony"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   2070
      Width           =   735
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "save"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblLoad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "load"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblClear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "clear song"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblRestart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "restart"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblSaveMidi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "save midi"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblClear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "clear track"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'QWERTY Music - by dafhi - July 03 2006

' - Project Highlights - '

'1. time-accurate recording with QueryPerformanceCounter
'2. same-time events shuffling - note offs go before note on
' to maximize polyphony

Private Const MaxPoly  As Integer = 16
Private Const MinPoly  As Integer = 16

' MinPoly when set lower than MaxPoly can be used to ignore
'   the most distant notes, rather than remove them from memory
'   Example: (maxpoly 16) (minpoly 2)
'   1. hold in this order -> CDEF .. only EF will play
'   2. release F .. DE will play
'   3. release E .. CD will play
'------------------------------------------------------'

Private Type SliderProc
    lLefPlus   As Integer
    lRghM      As Integer
    lWidM2     As Integer
    lPos       As Integer
    lPosP      As Integer
    MousDown   As Integer
End Type

Private Type SliderControl
    lTop       As Integer
    lLeft      As Integer
    lBot       As Integer
    lRight     As Integer
    lColor     As Long
    val_0_to_1 As Single
    Proc       As SliderProc
End Type

Dim mSong        As Song

Dim mSlidePos    As SliderControl

Dim mSavChkEvery As Integer
Dim mHaltChkB    As Integer

Dim mPosVisX     As Integer 'cosmetic
Dim mPosVisY     As Integer

Dim m_SamPos1    As Long 'calculate signal
Dim m_SamPos2    As Long

Dim mL1          As Long 'multi-purpose

Private Sub Form_Load()

    Font = "M": FontSize = 8
    
    DirectoryInfo
    
    Init_Synth hwnd, 1024, 22050
    
    If SynthAPI_flags = 0 Then Caption = "fmod.org for audio output"
    
    On_RecToggle
    
    vsbMinPoly.min = LMax(MaxPoly, MinPoly)
    vsbMinPoly.Value = vsbMinPoly.min
    vsbMinPoly.max = 1
    
    mPosVisX = 18
    mPosVisY = 40
    
    Move (Screen.Width - Me.Width) / 3, (Screen.Height - Me.Height) / 3
    
    Show
    
    chkNewEvery_Click

    Do While DoEvents
    
        DoSound
        
        If mSong.Proc.song_len <> 0 Then
            SliderVal mSlidePos, gTimeNow / mSong.Proc.song_len
        End If
        
        Render_NoteGraphics mSong, mPosVisX, mPosVisY
    
    Loop
    
    StopSound
    
End Sub
Private Sub DoSound()

    SoundPos m_SamPos1, m_SamPos2
    
    Sequencer mSong, sRender, m_SamPos1, m_SamPos2
    
    If WriteBuf Then
        'there was a buffer update
        
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    Case vbKeySpace
        If Shift = 0 Then
            Song_TogglePause
        Else
            lblRestart_Click
        End If
        UpdateRecLabel
    
    Case Else
        If Shift = 0 Then
            If gIsPaused Then
                Song_TogglePause
            End If
            If Record_QWERTY(mSong, KeyCode, NOTE_ON) Then
                UpdateRecLabel
            End If
        End If
    End Select
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Record_QWERTY mSong, KeyCode
End Sub

Private Sub Form_Paint()
    
    Cls
    
    Blit gSDESC, mPosVisX, mPosVisY
    Print "Spacebar ( + Shift to restart) to toggle record / pause,"
    Print "or you can simply start typing ... (q w e .. = c d e .. )"
    DrawSlider Me, mSlidePos

End Sub

Private Sub Form_Resize()
    
    ScaleMode = vbPixels
    
    Init_NoteGraphics hDC, ScaleWidth - 2 * mPosVisX - 20, 50, 33
    
    SliderDims mSlidePos, 95, mPosVisY + gSDESC.High + 1, 185, 10
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If x > mSlidePos.lLeft And x < mSlidePos.lRight Then
            If y > mSlidePos.lTop And y < mSlidePos.lBot Then
                mSlidePos.Proc.MousDown = True
                Form_MouseMove Button, Shift, x, y
            End If
        End If
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton And mSlidePos.Proc.MousDown = True Then
        If mSlidePos.Proc.lWidM2 > 0 Then
            SliderVal mSlidePos, (x - mSlidePos.Proc.lLefPlus) / mSlidePos.Proc.lWidM2
            SongPosition mSong, mSlidePos.val_0_to_1
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mSlidePos.Proc.MousDown = False
End Sub


Private Sub lblClear_Click(WhichLbl As Integer)
Dim L1 As Long
    
    If WhichLbl = 1 Then
        Song_ClearData mSong 'clear all tracks
        vsbTrack.min = 1
    Else
        Song_ClearData mSong, mSong.Proc.Ptr 'current track
        Play_Restart mSong
        If Not gIsPaused Then Song_TogglePause
    End If

    UpdateRecLabel
    
    Render_NoteGraphics mSong, mPosVisX, mPosVisY, True
    
End Sub
Private Sub lblLoad_Click()
    If LoadSong(mSong, ".dat") Then
        For mL1 = vsbTrack.min + 1 To mSong.Proc.cTrack
            mSong.Track(mL1).OnOff.Poly.Links.MinPoly = vsbMinPoly.Value
        Next
        vsbTrack.min = mSong.Proc.cTrack
        vsbTrack.Value = mSong.Proc.Ptr
        Render_NoteGraphics mSong, mPosVisX, mPosVisY, True
        UpdateRecLabel
        DirectoryInfo True
    End If
End Sub
Private Sub lblSave_Click()
    If SaveSong(mSong, ".dat") Then
        DirectoryInfo True
    End If
End Sub
Private Sub lblSaveMidi_Click()
    If SaveMidiFile(mSong) Then
        DirectoryInfo True
    End If
End Sub

Private Sub lblRestart_Click()
    Play_Restart mSong
    MaybeNewTrack
End Sub

Private Sub vsbMinPoly_Change()
    vsbMinPoly_Scroll
End Sub

Private Sub vsbMinPoly_Scroll()
Dim L1 As Long

    For L1 = 1 To mSong.Proc.cTrack
        mSong.Track(L1).OnOff.Poly.Links.MinPoly = vsbMinPoly.Value
    Next
    lblPoly.Caption = vsbMinPoly.Value
    
End Sub

Private Sub vsbTrack_Scroll()
    vsbTrack_Change
    chkNewEvery.Value = vbUnchecked
End Sub
Private Sub vsbTrack_Change()
    mSavChkEvery = chkNewEvery.Value
    mHaltChkB = True
    chkNewEvery.Value = mSavChkEvery
    
    Song_CheckHeldEvents mSong 'prior to changing Proc.Ptr
    mSong.Proc.Ptr = vsbTrack.Value
    
    lblTrack.Caption = mSong.Proc.Ptr
End Sub

Private Sub chkNewEvery_GotFocus()
    mSavChkEvery = chkNewEvery.Value 'keep spacebar from changing chkbox state
    mHaltChkB = True
End Sub
Private Sub chkNewEvery_Click()
    If mHaltChkB Then
        chkNewEvery.Value = mSavChkEvery
        mHaltChkB = False
    End If
End Sub
Private Sub chkNewEvery_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then mSavChkEvery = chkNewEvery.Value
    mHaltChkB = True
End Sub

Private Sub On_RecToggle()
    If Not gIsPaused Then
        MaybeNewTrack
    Else
        Play_Restart mSong
    End If
End Sub
Private Sub MaybeNewTrack()
Dim lTrackReady As Long
    
    If chkNewEvery.Value = vbChecked Or mSong.Proc.Ptr = 0 Then
    
        If mSong.Proc.Ptr = 0 Then
            lTrackReady = Song_NewTrack(mSong, MaxPoly, MinPoly)
        ElseIf mSong.Track(mSong.Proc.Ptr).OnOff.Seq.Events.StackPtr > 0 Then
            lTrackReady = Song_NewTrack(mSong, MaxPoly, MinPoly)
        End If

    End If
    
    If lTrackReady > 0 Then
        mSong.Track(lTrackReady).OnOff.Poly.Links.MinPoly = vsbMinPoly.Value
        vsbTrack.min = mSong.Proc.cTrack
        vsbTrack.Value = lTrackReady
        chkNewEvery.Value = vbChecked
    End If

End Sub
Private Sub UpdateRecLabel()
    If gIsPaused Then
        lblqRec.Caption = "paused"
    Else
        lblqRec.Caption = "Recording"
    End If
End Sub

Private Sub DrawSlider(pForm As Form, pSlidr As SliderControl, Optional ByVal pNewValue As Single = -1)

    pForm.Line (pSlidr.lLeft, pSlidr.lTop)-(pSlidr.lRight, pSlidr.lBot), pSlidr.lColor, B
    If pNewValue <> -1 Then SliderVal pSlidr, pNewValue
        
    m_DrawSlidePos pForm, pSlidr
    
End Sub
Private Sub m_DrawSlidePos(pForm As Form, pSlidr As SliderControl)
    If pSlidr.Proc.lPos <> pSlidr.Proc.lPosP Then
        pForm.Line (pSlidr.Proc.lPosP, pSlidr.lTop + 1)-(pSlidr.Proc.lPosP, pSlidr.lBot), pForm.BackColor
        pForm.Line (pSlidr.Proc.lPos, pSlidr.lTop + 1)-(pSlidr.Proc.lPos, pSlidr.lBot), pSlidr.lColor
        pSlidr.Proc.lPosP = pSlidr.Proc.lPos
    End If
End Sub
Private Sub SliderVal(pSlidr As SliderControl, ByVal pNewValue As Single)
    If pNewValue < 0 Then
        pSlidr.val_0_to_1 = 0
    ElseIf pNewValue > 1 Then
        pSlidr.val_0_to_1 = 1
    Else
        pSlidr.val_0_to_1 = pNewValue
    End If
    pSlidr.Proc.lPos = pSlidr.Proc.lLefPlus + pSlidr.val_0_to_1 * pSlidr.Proc.lWidM2
    m_DrawSlidePos Me, mSlidePos
End Sub
Private Sub SliderDims(pSlidr As SliderControl, Optional ByVal pLeft As Integer = -1, Optional ByVal pTop As Integer = -1, Optional ByVal pWid As Integer = -1, Optional ByVal pHgt As Integer = -1)
    If pLeft < 0 Then pLeft = 0
    If pTop < 0 Then pTop = 0
    If pWid < 5 Then pWid = 5
    If pHgt < 5 Then pHgt = 5
    pSlidr.lLeft = pLeft
    pSlidr.lTop = pTop
    pSlidr.lBot = pTop + pHgt - 1
    pSlidr.lRight = pLeft + pWid - 1
    pSlidr.Proc.lLefPlus = pLeft + 1
    pSlidr.Proc.lRghM = pSlidr.lRight - 1
    pSlidr.Proc.lWidM2 = pSlidr.Proc.lRghM - pSlidr.Proc.lLefPlus
End Sub
