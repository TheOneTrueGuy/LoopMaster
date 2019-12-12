Attribute VB_Name = "Module1"
 Public Type CUSTOMVERTEX
    X As Single         'x in screen space
    Y As Single         'y in screen space
    z  As Single        'normalized z
    rhw As Single
    color As Long      'vertex color
End Type
Private Type WaveFormat
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
End Type









'
'Private Sub Command4_Click()
'Dim intbuffer() As Integer
'
'Dim bufs As Collection
'Set bufs = New Collection
''ReDim bufs(ccount)
'ReDim bufoffset(ccount) As Long
'For tzl = 0 To ccount - 1
'' get buffer from each child
'' chop and blend
'' play output
'intbuffer = chils(tzl).dropLoad
'bufs.Add intbuffer
'Next tzl
'
'Dim ldx As DirectX8
'Dim lds As DirectSound8
'Dim lbuf As DirectSoundSecondaryBuffer8
'Dim bufdesc As DSBUFFERDESC
'bufdesc.fxFormat = chils(0).WaveEx(22050, 1, 16)
'bufdesc.lBufferBytes = CLng(22000) * 5
'bufdesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
'Set ldx = New DirectX8
'Set lds = ldx.DirectSoundCreate(vbNullString)
'Set lbuf = lds.CreateSoundBuffer(bufdesc)
'Dim bufit(CLng(22000) * 5) As Integer
'Dim bzl As Long
'Dim factor
'For bzl = 1 To (CLng(22000) * 5) - ccount Step ccount
'For tzl = 1 To ccount
''If bzl + tzl > UBound(bufit) Then Exit For
'bufi = bufs.Item(tzl)
'bufoffset(tzl) = bufoffset(tzl) + 1
''If bufOffset(tzl) > UBound(bufi) Then bufOffset(tzl) = 1 ' tiles, no stretch yet
'bufoffset(tzl) = bufoffset(tzl) Mod UBound(bufi)
'
'' segment from each array.
'bufit(bzl + tzl) = bufi(bufoffset(tzl))
'
'Next tzl
'Next bzl
'lbuf.WriteBuffer 0, CLng(22000) * 5, bufit(0), DSBLOCK_ENTIREBUFFER
'
'lbuf.SaveToFile "c:\windows\desktop\testZZ.wav"
'lbuf.Play DSBPLAY_DEFAULT
'
'
'
'End Sub
Private Type ChunkHeader
    lType As Long
    lLen As Long
End Type


Private Type FileHeader
    lRiff As Long
    lFileSize As Long
    lWave As Long
    lFormat As Long
    lFormatLength As Long
End Type





Sub Main()
    
End Sub

