Attribute VB_Name = "mMySynth"
Option Explicit

' - Dependencies -

'mPolyRec.bas
'mGeneral.bas

    'This sub is called by m_PlayPoly2 (mPolyRec.bas)

Public Sub StereoArchitecture(pSeqP As SeqAndPoly, pFreq As NoteInfo, sngRender() As Single, ByVal SamStart As Long, ByVal SamStop As Long, Optional dblPeriod As Single = 1, Optional Use1stAlgo As Boolean = False)

    m_WriteToSignal pSeqP, sngRender, SamStart, SamStop, pFreq.InfoL, pFreq.dIncr, dblPeriod
    If PCM.nChannels = 2 Then
        m_WriteToSignal pSeqP, sngRender, SamStart + 1, SamStop + 1, pFreq.InfoR, pFreq.dIncr, dblPeriod
    End If
    
End Sub
Private Sub m_WriteToSignal(pSeqP As SeqAndPoly, sngRender() As Single, ByVal pStart As Long, ByVal pStop As Long, ChanI As ChanInf, ByVal iPos As Double, Optional dblPeriod As Single = 1)
Dim dPos As Double, L1 As Long
Dim dPos2 As Double, iPos2 As Double

    If dblPeriod <= 0 Then dblPeriod = 1
    
    'less typing ..
    dPos = ChanI.dPos
    iPos2 = iPos * TwoPi
    dPos2 = dPos * TwoPi
    If pStart < 0 Then Exit Sub
    
    ' Example synthesizer algorithm
    
    For L1 = pStart To pStop Step PCM.nChannels
        sngRender(L1) = sngRender(L1) + 0.23 * Sin(dPos2 + 8 * Triangle(dPos))
        dPos = dPos + iPos
        dPos2 = dPos2 + iPos2
    Next
    
    ' OVERFLOW PREVENTION based on a period
    If pSeqP.Poly.Links.StackPtr < 1 Then dPos = dPos - dblPeriod * Int(dPos / dblPeriod)
    ChanI.dPos = dPos

End Sub

