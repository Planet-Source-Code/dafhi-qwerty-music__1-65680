Attribute VB_Name = "mDrawNotes"
Option Explicit

Public gSDESC        As SurfaceDescriptor

Dim mPelWid1S        As Long
Dim mSecsPicWid      As Single
Dim mSecsWidBy2      As Single
Dim mTimeRightLim    As Single
Dim mTimeLeftLim     As Single
Dim mTmpSongLen      As Single
Dim mTimeV           As Long
Dim mTimeP           As Long
Dim mElem            As Long
Dim mNoteTime1       As Single
Dim mNoteTime2       As Single
Dim mX1 As Long
Dim mX2 As Long
Dim mLY As Long
Dim mL1 As Long
Dim mColor As Long

Private Type EraseInfo
    Stk        As Long
End Type

Dim mErasMain        As EraseInfo
Dim mPixelRef()      As Long

Dim mSA              As SAFEARRAY1D
Dim mSurf1D()        As Long
Dim mHSV()           As Integer
Dim mSaveX1          As Long
Dim mSaveX2          As Long

Dim mHueBase         As Integer
Dim iHue             As Integer

Public Sub Init_NoteGraphics(ByVal pDC As Long, ByVal pWid As Integer, ByVal pHgt As Integer, Optional ByVal PixWidForOneSecond As Integer = 50)

    If CLng(pWid) * CLng(pHgt) > 999999 Then Exit Sub
    
    mPelWid1S = PixWidForOneSecond
    SetSurfaceDesc gSDESC, gSDESC.Dib32, pDC, pWid, pHgt
    
    Erase mHSV
    ReDim mHSV(gSDESC.U1D)
    
    Erase mSurf1D
    ReDim mSurf1D(gSDESC.U1D)
    
    Erase mPixelRef
    ReDim mPixelRef(gSDESC.U1D)
    
    mErasMain.Stk = -1: mTimeP = -1
    
    mSA.cbElements = 4
    mSA.cDims = 1
    mSA.cElements = gSDESC.U1D + 1
    mSA.pvData = VarPtr(gSDESC.Dib32(0, 0))
    
    mSaveX2 = mSaveX1 - 1

End Sub
Public Sub Render_NoteGraphics(pSong As Song, ByVal pX As Integer, pY As Integer, Optional ByVal pForceRender As Boolean)
Dim L1 As Long
Dim lSecsTripRgh As Single
Dim lSecsTripLef As Single
Dim lTimeRendLef As Single

    mTimeV = Int(gTimeNow * mPelWid1S + 0.5)
    If mTimeP = mTimeV And pForceRender = False Then Exit Sub  'pixel pos has not changed
    mTimeP = mTimeV
    
    If gSDESC.U1D < 1 Or pSong.Proc.cTrack < 1 Then Exit Sub
    If mPelWid1S < 1 Then mPelWid1S = 1
    If mPelWid1S > 500 Then mPelWid1S = 500
    
    mTmpSongLen = pSong.Proc.song_len
    mSecsPicWid = gSDESC.WM / mPelWid1S
    
    'position bar will stay at half view width for a while
    'if song length is greater than visualization window
    mSecsWidBy2 = mSecsPicWid / 2
    
    L1 = Int(gSDESC.WM / 2) + 1
    lSecsTripLef = L1 / mPelWid1S
    
    GetTimeNow
    If mTmpSongLen < gTimeNow Then mTmpSongLen = gTimeNow
    
    If mTmpSongLen > mSecsPicWid Then
        mTimeLeftLim = mTmpSongLen - mSecsPicWid
        lSecsTripRgh = mTimeLeftLim + lSecsTripLef
        If gTimeNow > lSecsTripRgh Then
        
            'nothing to do here
            
        ElseIf gTimeNow > mSecsWidBy2 Then
            mTimeLeftLim = gTimeNow - mSecsWidBy2
        Else
            mTimeLeftLim = 0
        End If
    Else
        mTimeLeftLim = 0
    End If
    
    mTimeRightLim = mTimeLeftLim + mSecsPicWid
    
    'creating a 1d array same size as 2d gSDESC.Dib32
    CopyMemory ByVal VarPtrArray(mSurf1D), VarPtr(mSA), 4

    z_Erase mErasMain
    
    iHue = 35
    mHueBase = 0
    For L1 = 1 To pSong.Proc.cTrack
        z_RenderNoteGFX pSong.Track(L1).OnOff.Seq, gSDESC.Dib32
        Add mHueBase, 113
    Next
    
    z_DrawTimeBar
    
    CopyMemory ByVal VarPtrArray(mSurf1D), CLng(0), 4
    
    Blit gSDESC, pX, pY

End Sub
Private Sub z_Erase(pErase As EraseInfo)
    For pErase.Stk = pErase.Stk To 0 Step -1
        mHSV(mPixelRef(pErase.Stk)) = 0
        mSurf1D(mPixelRef(pErase.Stk)) = 0
    Next
    For mLY = mSaveX1 To mSaveX2 Step gSDESC.Wide
        mSurf1D(mLY) = vbBlack
    Next
End Sub
Private Sub z_RenderNoteGFX(pSeq As SequencerElement, p2D() As Long)
Dim lExit As Boolean

    mElem = pSeq.Events.LinkHead
    For mL1 = 1 To pSeq.Events.StackPtr
        z_If_Note_InView pSeq, pSeq.NINFO(mElem), p2D, lExit
        If lExit Then Exit For
        mElem = pSeq.Events.Link(mElem).Next
    Next
    
End Sub
Private Sub z_If_Note_InView(pSeq As SequencerElement, pNE As NoteEvent, p2D() As Long, pRetIfDone As Boolean)
Dim xTimesY_ As Long, lHue As Integer, lSat As Integer

    If pNE.Event = NOTE_ON Then
        mNoteTime1 = pNE.abs_time
        If pNE.Ref > 0 Then
            mNoteTime2 = pSeq.NINFO(pNE.Ref).abs_time
        Else 'note is being held down
            mNoteTime2 = gTimeNow
        End If
        mLY = pNE.sNote - 40 'pixel height
    ElseIf pNE.Event = NOTE_OFF Then
        mNoteTime2 = pNE.abs_time
        If pNE.Ref > 0 Then
            mNoteTime1 = pSeq.NINFO(pNE.Ref).abs_time
            mLY = pSeq.NINFO(pNE.Ref).sNote - 40
        Else
            mNoteTime1 = gTimeNow
        End If
    End If
    
    
    If mNoteTime2 >= mTimeLeftLim Then
        If mNoteTime1 < mTimeRightLim Then
            If mLY < gSDESC.High And mLY > -1 Then
                mX1 = Int((mNoteTime1 - mTimeLeftLim) * mPelWid1S + 0.5)
                mX2 = Int((mNoteTime2 - mTimeLeftLim) * mPelWid1S + 0.5)
                If mX1 < 0 Then mX1 = 0
                If mX2 > gSDESC.WM Then mX2 = gSDESC.WM
                xTimesY_ = mLY * gSDESC.Wide + mX1
                For mX1 = mX1 To mX2
                    If mSurf1D(xTimesY_) = 0 Then
                        mErasMain.Stk = mErasMain.Stk + 1
                        mPixelRef(mErasMain.Stk) = xTimesY_
                    End If
                    lHue = mHSV(xTimesY_)
                    mSurf1D(xTimesY_) = ARGBHSV(mHueBase + iHue * lHue, 1, 255)
                    mHSV(xTimesY_) = lHue + 1
                    xTimesY_ = xTimesY_ + 1
                Next
            End If
        Else
            pRetIfDone = True
        End If
    End If
    
End Sub
Private Sub z_DrawTimeBar()

    mColor = ARGBHSV(Rnd * 1530, 1, 255)
    
    mX1 = Int((gTimeNow - mTimeLeftLim) * mPelWid1S + 0.5)
    If mX1 > gSDESC.WM Then mX1 = gSDESC.WM
    mX2 = mX1 + gSDESC.HM * gSDESC.Wide
    
    For mLY = mX1 To mX2 Step gSDESC.Wide
        mSurf1D(mLY) = mColor
    Next
    
    mSaveX1 = mX1
    mSaveX2 = mX2

End Sub

