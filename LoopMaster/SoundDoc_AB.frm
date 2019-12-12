VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSoundDoc 
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   9765
   Begin VB.OptionButton Option1 
      Caption         =   "Repeat"
      Height          =   240
      Left            =   60
      TabIndex        =   4
      Top             =   1755
      Width           =   1245
   End
   Begin VB.CommandButton Loopit 
      Caption         =   "Loopit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5325
      TabIndex        =   3
      Top             =   1695
      Width           =   1110
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   330
      Left            =   2460
      TabIndex        =   2
      Top             =   1710
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   330
      Left            =   1395
      TabIndex        =   1
      Top             =   1710
      Width           =   960
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   9060
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   1485
      Left            =   45
      ScaleHeight     =   1425
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   135
      Width           =   9675
   End
   Begin VB.Menu fyl 
      Caption         =   "File"
      Begin VB.Menu opn 
         Caption         =   "Open"
      End
      Begin VB.Menu sayv 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "frmSoundDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Implements DirectXEvent8
Public mynumber As Long
Dim dx3d As Direct3D8
Dim dx3ddev(2) As Direct3DDevice8
'Dim dxev As DirectXEvent
Dim ds As DirectSound8

Private capFormat As WAVEFORMATEX
Dim DSBPOSIT As DSBPOSITIONNOTIFY
Dim dx8ev As DirectXEvent8
Dim eventid As Long
Dim dsb As DirectSoundSecondaryBuffer8
Dim dsbd As DSBUFFERDESC
Dim dscd As DSCBUFFERDESC
Dim wavformat As WAVEFORMATEX
Dim capCURS As DSCURSORS
Dim ByteBuffer() As Integer
Dim OByteBuffer() As Integer
Dim CNT As Integer
Dim Caps As DSCAPS
Public numbytes As Long
Public guid As String
Dim devent As DirectXEvent8, hndl_Devent As Long
Dim dsb_PoNo(3) As DSBPOSITIONNOTIFY 'Dim dx8ev As DirectXEvent8

'_________________ bitblt and other scrolling stuff
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Const LR_LOADFROMFILE As Long = &H10
'Const LR_CREATEDIBSECTION As Long = &H2000
'****************************************
Dim alreadymade As Boolean
Dim vertices1() As CUSTOMVERTEX
Dim g_VB As Direct3DVertexBuffer8
Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)
Dim samplesize As Integer, sizeOfVertex As Integer
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim dx As New DirectX8 'Our DirectX object
Dim sourceBuf As DirectSoundSecondaryBuffer8, reflen As Long 'reference SoundBuffer
Dim ds3dBuffer As DirectSound3DBuffer8 'We need to get a 3DSoundBuffer
Private oFX As DirectSoundFXEcho8, mlIndex As Long
Dim exo() As DSFXECHO
Dim delayarray() As Single
Dim tempomarkers() As Single
Dim segBufs() As DirectSoundSecondaryBuffer8
Dim markcount As Integer
Dim oPos As D3DVECTOR 'Position
Dim fMouseDown As Boolean 'Is the mouse down?
Dim proBuf As DirectSoundSecondaryBuffer8
Dim intbuf() As Integer
Public Function LoadFile()
'
'On Error GoTo erhandl
initDevices

Dim reffyl$
cdlFile.Filter = "Wav files (*.wav)|*.wav"

cdlFile.ShowOpen
reffyl = cdlFile.FileName
If reffyl = "" Or LCase(Right(reffyl, 3)) <> "wav" Then Exit Function
Set sourceBuf = ds.CreateSoundBufferFromFile(reffyl, dsbd)
Debug.Print "dsbd.lBufferBytes "; dsbd.lBufferBytes
Debug.Print "dsbd.lFlags "; dsbd.lFlags
Debug.Print "dsbd.fxFormat.lAvgBytesPerSec "; dsbd.fxFormat.lAvgBytesPerSec
Debug.Print "dsbd.fxFormat.lSamplesPerSec "; dsbd.fxFormat.lSamplesPerSec

reflen = dsbd.lBufferBytes + 1
ReDim intbuf(reflen)
sourceBuf.ReadBuffer 0, 0, intbuf(0), DSBLOCK_ENTIREBUFFER
reflen = UBound(intbuf)
Debug.Print reflen
'hndl_Devent = dx.CreateEvent(Me)
'dsb_PoNo(0).hEventNotify = hndl_Devent
'dsb_PoNo(0).lOffset = 1000
'dsb_PoNo(1).hEventNotify = hndl_Devent
'dsb_PoNo(1).lOffset = 1200
'dsb_PoNo(2).hEventNotify = hndl_Devent
'dsb_PoNo(2).lOffset = 1300
'refBuf.SetNotificationPositions 3, dsb_PoNo
'
Dim res As Boolean

If Not drawRef Then Stop
Exit Function
erhandl:
' check if posnotify is too long?


End Function
Public Function drawRef() As Boolean
On Error GoTo erhandl
' draw the reference buffer into picbox1

drawRef = True
Exit Function
erhandl:
drawRef = False
End Function

Private Sub Command1_Click()
If Option1.Value Then
sourceBuf.Play DSBPLAY_LOOPING
Else
sourceBuf.Play DSBPLAY_DEFAULT
End If
End Sub

Private Sub Command2_Click()
sourceBuf.Stop
End Sub




Private Sub Command5_Click()
Dim bufdesc As DSBUFFERDESC
Dim blen As Long, lsys$
blen = CLng(InputBox("Number of seconds to generate"))
lsys = InputBox("Enter Lsys sequence")

bufdesc.fxFormat = WaveEx(22050, 1, 16)
bufdesc.lBufferBytes = CLng(22000) * blen
bufdesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
Set proBuf = ds.CreateSoundBuffer(bufdesc)
ReDim bufit(CLng(22000) * blen) As Integer
Dim tzl As Long
Dim factor, lcount As Integer
factor = Rnd * 100
For tzl = 1 To CLng(22000) * blen
'If tzl / 100 = Int(tzl / 100) Then factor = Rnd * 100
'bufit(tzl) = (500 + factor) - (Cos(tzl / factor) * (900 + factor))
If lcount > Len(lsys) Then lcount = 0
lcount = lcount + 1
If Mid(lsys, lcount, 1) = "u" Then factor = factor + 1000 Else factor = factor - 1000
If factor > 32000 Then factor = 32000
If factor < -32000 Then factor = -32000
bufit(tzl) = factor


Next tzl
proBuf.WriteBuffer 0, CLng(22000) * blen, bufit(0), DSBLOCK_ENTIREBUFFER
proBuf.SaveToFile "c:\windows\desktop\test.wav"
proBuf.Play DSBPLAY_DEFAULT
End Sub


Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
'
End Sub

Private Sub Form_Load()
'
End Sub

Private Sub Form_Unload(Cancel As Integer)
sourceBuf.Stop
Set sourceBuf = Nothing
Set ds = Nothing
Set dx = Nothing
End Sub

Public Sub initDevices()
  Set ds = dx.DirectSoundCreate(vbNullString)
  ds.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL
'   Set dsb = ds.CreateSoundBuffer(dsbd)
 Call Initbuffer
End Sub
Private Sub Initbuffer()
    ds.GetCaps Caps
    wavformat = WaveEx(22050, 1, 16)
    dsbd.fxFormat = wavformat
    dsbd.lBufferBytes = wavformat.lAvgBytesPerSec * 20
    dsbd.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
      
    'Set dsb = ds.CreateSoundBuffer(dsbd)
End Sub
Public Function WaveEx(Hz As Long, Channels As Integer, BITS As Integer) As WAVEFORMATEX

    WaveEx.nFormatTag = WAVE_FORMAT_PCM
    WaveEx.nChannels = Channels
    WaveEx.lSamplesPerSec = Hz
    WaveEx.nBitsPerSample = BITS
    WaveEx.nBlockAlign = Channels * BITS / 8
    WaveEx.lAvgBytesPerSec = WaveEx.lSamplesPerSec * WaveEx.nBlockAlign
    WaveEx.nSize = 0

End Function

Public Function dropLoad() As Integer()
ReDim bufout(reflen) As Integer

sourceBuf.ReadBuffer 0, 1, bufout(0), DSBLOCK_ENTIREBUFFER
dropLoad = bufout

End Function

Private Sub Loopit_Click()
' don't alter beginning of loop (necessarily)
' just alter end to match.
' Either average last n samples to match last sample to beginning
' amplitude.



'Error 1
' get beginning and end of buffer
' compare amplitudes.
' general pitch of amplitude slope
' and final amplitude values.
' adjust last n samples to match/overlap
'intbuf(reflen)
Dim tzl, dzl
Dim sam1(20), sam2(20)
For tzl = 1 To 20
sam1(tzl) = intbuf(tzl)
sam2(tzl) = intbuf(UBound(intbuf) - 20 + tzl)
Next tzl

Dim bufdesc As DSBUFFERDESC
Dim blen As Long, lsys$
bufdesc.fxFormat = WaveEx(22050, 1, 16)
bufdesc.lBufferBytes = reflen ' same length as reference
bufdesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPOSITIONNOTIFY
Set proBuf = ds.CreateSoundBuffer(bufdesc)
ReDim bufit(CLng(44000) * blen) As Integer
Dim tzl As Long


proBuf.WriteBuffer 0, CLng(22000) * blen, bufit(0), DSBLOCK_ENTIREBUFFER
proBuf.SaveToFile "c:\windows\desktop\test.wav"
proBuf.Play DSBPLAY_DEFAULT



End Sub

Private Sub opn_Click()
LoadFile
End Sub

