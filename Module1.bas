Attribute VB_Name = "Module1"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE
    dmDeviceName       As String * CCDEVICENAME
    dmSpecVersion      As Integer
    dmDriverVersion    As Integer
    dmSize             As Integer
    dmDriverExtra      As Integer
    dmFields           As Long
    dmOrientation      As Integer
    dmPaperSize        As Integer
    dmPaperLength      As Integer
    dmPaperWidth       As Integer
    dmScale            As Integer
    dmCopies           As Integer
    dmDefaultSource    As Integer
    dmPrintQuality     As Integer
    dmColor            As Integer
    dmDuplex           As Integer
    dmYResolution      As Integer
    dmTTOption         As Integer
    dmCollate          As Integer
    dmFormName         As String * CCFORMNAME
    dmUnusedPadding    As Integer
    dmBitsPerPel       As Integer
    dmPelsWidth        As Long
    dmPelsHeight       As Long
    dmDisplayFlags     As Long
    dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lptypDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Global hWaveMix As Long
Global lpWaveMix() As Long
Global WaveHandle As Long
Global WAVMIX_Quiet As Integer

Global Const WAVEMIX_MAXCHANNELS = 8

Type tChannelInfo
    Loops As Long
    WaveFile As String
End Type


Type WAVEMIXINFO
   wSize As Integer
   bVersionMajor As String * 1
   bVersionMinor As String * 1
   szDate(12) As String
   dwFormats As Long
End Type

Type MIXCONFIG
    wSize As Integer
    dwFlagsLo As Integer
    dwFlagsHi As Integer
    wChannels As Integer
    wSamplingRate As Integer
End Type

Private Type MIXPLAYPARAMS
    wSize         As Integer
    hMixSessionLo As Integer
    hMixSessionHi As Integer
    iChannelLo    As Integer
    iChannelHi    As Integer
    lpMixWaveLo   As Integer
    lpMixWaveHi   As Integer
    hWndNotifyLo  As Integer
    hWndNotifyHi  As Integer
    dwFlagsLo     As Integer
    dwFlagsHi     As Integer
    wLoops        As Integer
End Type

Declare Function WaveMixInit Lib "WAVMIX32.DLL" () As Long
Declare Function WaveMixConfigureInit Lib "WAVMIX32.DLL" (lpConfig As MIXCONFIG) As Long
Declare Function WaveMixActivate Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal fActivate As Integer) As Long
Declare Function WaveMixOpenWave Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, szWaveFilename As Any, ByVal hInst As Long, ByVal dwFlags As Long) As Long
Declare Function WaveMixOpenChannel Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal iChannel As Long, ByVal dwFlags As Long) As Long
Declare Function WaveMixPlay Lib "WAVMIX32.DLL" (lpMixPlayParams As Any) As Integer
Declare Function WaveMixFlushChannel Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal iChannel As Integer, ByVal dwFlags As Long) As Integer
Declare Function WaveMixCloseChannel Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal iChannel As Integer, ByVal dwFlags As Long) As Integer
Declare Function WaveMixFreeWave Lib "WAVMIX32.DLL" (ByVal hMixSession As Long, ByVal lpMixWave As Long) As Integer
Declare Function WaveMixCloseSession Lib "WAVMIX32.DLL" (ByVal hMixSession As Long) As Integer
Declare Sub WaveMixPump Lib "WAVMIX32.DLL" ()
Declare Function WaveMixGetInfo Lib "WAVMIX32.DLL" (lpWaveMixInfo As WAVEMIXINFO) As Integer
Dim Respons As Long
Dim Wait1 As Boolean
'------------------------------'
'         CaptiveX TM.         '
'   Writen by nofx (op-ivy)    '
'http://www.sharpnet.net/~nofx/'
'or visit us on EFNET #captivex'
'             P.S.             '
'           Have Fun           '
'------------------------------'

#If Win32 Then
Public Const HWND_TOPMOST& = -1
#Else
Public Const HWND_TOPMOST& = -1
#End If 'WIN32

#If Win32 Then
 Const SWP_NOMOVE& = &H2
 Const SWP_NOSIZE& = &H1
#Else
 Const SWP_NOMOVE& = &H2
 Const SWP_NOSIZE& = &H1
#End If 'WIN32

#If Win32 Then
 Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#Else
 Declare Sub SetWindowPos Lib "user" (ByVal hWnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#End If 'WIN32

Function StayOnTop(Form As Form) 'EX: Call StayOnTop(Me)
Dim lFlags As Long
Dim lStay As Long

'lFlags = SWP_NOSIZE Or SWP_NOMOVE
'lStay = SetWindowPos(Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Function DStayOnTop(Form As Form) 'EX: Call StayOnTop(Me)
Dim lFlags As Long
Dim lStay As Long

lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hWnd, -2, 0, 0, 0, 0, lFlags)
End Function


Private Function HiWord(ByVal l As Long) As Integer
    l = l \ &H10000
    
    HiWord = Val("&H" & Hex$(l))
End Function

Private Function LoWord(ByVal l As Long) As Integer
    l = l And &HFFFF&
    
    LoWord = Val("&H" & Hex$(l))
End Function


Function WAVMIX_AddFile(Filename As String) As Integer
'------------------------------------------------------------
' Open a wave file and assign it to the next available
' channel.
'------------------------------------------------------------
Dim wRtn As Long

    WAVMIX_AddFile = False
    If WAVMIX_Quiet Then Exit Function
    If WaveHandle + 1 = WAVEMIX_MAXCHANNELS Then Exit Function

    ReDim Preserve lpWaveMix(WaveHandle)
    lpWaveMix(WaveHandle) = WaveMixOpenWave(hWaveMix, ByVal Filename, 0, 0)
    wRtn = WaveMixOpenChannel(hWaveMix, WaveHandle, 0)
    WAVMIX_AddFile = WaveHandle
    WaveHandle = WaveHandle + 1
End Function

Sub WAVMIX_SetFile(Filename As String, AChannel As Long)
'------------------------------------------------------------
' Assign a new wave file, FileName, to the specified channel,
' AChannel.  If this channel is currently assigned another
' wave file, stop playing the channel and free the active
' wave file.
'------------------------------------------------------------
Dim wRtn As Long
    
    If WAVMIX_Quiet Then Exit Sub
    
    If AChannel > UBound(lpWaveMix) Then
        ReDim Preserve lpWaveMix(AChannel)
        WaveHandle = AChannel
    End If
    
    ' If another wave is currently assigned to this
    ' channel, free it.
    If lpWaveMix(AChannel) <> 0 Then
        WAVMIX_StopChannel AChannel
        wRtn = WaveMixFreeWave(hWaveMix, lpWaveMix(AChannel))
    End If
    
    ' Open the new wave and assign it to this channel.
    lpWaveMix(AChannel) = WaveMixOpenWave(hWaveMix, ByVal Filename, 0, 0)
    wRtn = WaveMixOpenChannel(hWaveMix, AChannel, 0)
End Sub


Sub WAVMIX_Close()
'------------------------------------------------------------
' Stop playing all channels and free all wave files, then
' close down this WaveMix session.
'------------------------------------------------------------
Dim wRtn As Long
Dim I As Integer, rc As Integer
    
    If WAVMIX_Quiet Then Exit Sub

    If (hWaveMix <> 0) Then
        For I = 0 To UBound(lpWaveMix)
            If lpWaveMix(I) <> 0 Then
                WAVMIX_StopChannel CLng(I)
                rc = WaveMixFreeWave(hWaveMix, lpWaveMix(I))
            End If
        Next
        wRtn = WaveMixCloseSession(hWaveMix)
        hWaveMix = 0
    End If
End Sub

Function WAVMIX_InitMixer() As Integer
'------------------------------------------------------------
' Initialize and activate the WaveMix DLL.
'------------------------------------------------------------
Dim wRtn As Long
Dim config As MIXCONFIG

    If WAVMIX_Quiet Then Exit Function

    WaveHandle = 0
    ReDim lpWaveMix(0)
    ChDir App.Path
    
    config.wSize = Len(config)
    config.dwFlagsHi = 1
    config.dwFlagsLo = 0
    'Allow stereo sound
    config.wChannels = 2
    hWaveMix = WaveMixConfigureInit(config)
    wRtn = WaveMixActivate(hWaveMix, True)

    If (wRtn <> 0) Then
        WAVMIX_InitMixer = False
        Call WaveMixCloseSession(hWaveMix)
        hWaveMix = 0
    Else
        WAVMIX_InitMixer = True
    End If
End Function

Sub WAVMIX_StopChannel(ByVal ChannelNum As Long)
'------------------------------------------------------------
' Stop playing the specified channel.
'------------------------------------------------------------
Dim rc As Integer
If WAVMIX_Quiet Then Exit Sub
If (hWaveMix = 0) Then Exit Sub
rc = WaveMixFlushChannel(hWaveMix, ChannelNum, 0)
End Sub

Sub WAVMIX_Activate(Activate As Long)
'------------------------------------------------------------
' Activate the WaveMix DLL.
'------------------------------------------------------------
Dim rc As Integer

    If WAVMIX_Quiet Then Exit Sub
    If (hWaveMix = 0) Then Exit Sub

    rc = WaveMixActivate(hWaveMix, Activate)
End Sub

Sub WAVMIX_PlayChannel(ChannelNum As Long, LoopWave As Long)
'------------------------------------------------------------
' Play a specified channel, and indicate whether the sound
' should be looped.
'------------------------------------------------------------
Dim params As MIXPLAYPARAMS
Dim wRtn As Long
If WAVMIX_Quiet Then Exit Sub
If ChannelNum > UBound(lpWaveMix) Then Exit Sub
If (hWaveMix = 0) Then Exit Sub
Dim rc As Integer
rc = WaveMixFlushChannel(hWaveMix, ChannelNum, 2)
params.wSize = Len(params)
params.hMixSessionLo = LoWord(hWaveMix)
params.hMixSessionHi = HiWord(hWaveMix)
params.iChannelLo = LoWord(ChannelNum)
params.iChannelHi = HiWord(ChannelNum)
params.lpMixWaveLo = LoWord(lpWaveMix(ChannelNum))
params.lpMixWaveHi = HiWord(lpWaveMix(ChannelNum))
params.hWndNotifyLo = 0
params.hWndNotifyHi = 0
params.dwFlagsHi = 0
params.dwFlagsLo = 0
params.wLoops = LoopWave
wRtn = WaveMixPlay(params)
End Sub

Function PlayBackGround(Filename As String)
sndPlaySound Filename, &H1 Or &H8
End Function

Function Stop_Playing()
sndPlaySound vbNullString, &H0
WAVMIX_Close
End Function

Function Play(Number As Long)
WAVMIX_PlayChannel Number, 0
Form1.Timer2.Interval = 10000
End Function

Function SetFile(v As Boolean)
Dim Filename As String
Select Case Int((2 * Rnd) + 1)
Case 1
Filename = App.Path & "\Sound\Food1.wav"
Case 2
Filename = App.Path & "\Sound\Food2.wav"
End Select
WAVMIX_SetFile Filename, 2
If F = True Then Exit Function
Select Case Int((3 * Rnd) + 1)
Case 1
Filename = App.Path & "\Sound\gameover1.wav"
Case 2
Filename = App.Path & "\Sound\gameover2.wav"
Case 3
Filename = App.Path & "\Sound\gameover3.wav"
End Select
WAVMIX_SetFile Filename, 3
Select Case Int((2 * Rnd) + 1)
Case 1
Filename = App.Path & "\Sound\levelwin1.wav"
Case 2
Filename = App.Path & "\Sound\levelwin2.wav"
End Select
WAVMIX_SetFile Filename, 1
WAVMIX_SetFile App.Path & "\Sound\Food.wav", 4
WAVMIX_SetFile App.Path & "\Sound\Start.wav", 5
WAVMIX_SetFile App.Path & "\Sound\Superapple.wav", 6
WAVMIX_SetFile App.Path & "\Sound\st.wav", 7
'WAVMIX_SetFile App.Path & "\Sound\Hiscore.wav", 8
End Function

Function SetFile2()
WAVMIX_SetFile App.Path & "\Sound\Menu1.wav", 1
WAVMIX_SetFile App.Path & "\Sound\Menu2.wav", 2
End Function
