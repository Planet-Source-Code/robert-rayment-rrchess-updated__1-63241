Attribute VB_Name = "ModResWav"
' ModResWav.bas   ~RRChess~

Option Explicit

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Purges better ?
Private Declare Function PlayResWAV Lib "winmm" Alias "sndPlaySoundA" _
   (ByVal lpszName&, ByVal dwFlags&) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, Source As Any, ByVal Length As Long)

'Private Const SND_ASYNC = &H1         ' return to program immediately
Private Const SND_MEMORY = &H4        ' play the sound from memory
Private Const SND_NODEFAULT = &H2     ' don't play the default sound if not available
'Private Const SND_NOSTOP = &H10       ' don't stop a currently playing sound
'Private Const SND_NOWAIT = &H2000     ' return immediately if driver not available
Private Const SND_PURGE = &H40        ' purge non-static events for task
'Private Const SND_LOOP = &H8          ' loop the sound until next sndPlaySound
'Private Const SND_RESOURCE = &H40004  ' play from resource

' sndPlaySound WAVData, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

'   Call PlaySoundMem(VarPtr(Sound(0)), 0, SND_NOWAIT Or SND_NODEFAULT _
'       Or SND_MEMORY Or SND_ASYNC Or SND_NOSTOP)

''''' Used :- ''''''''''''''''''''''''''''''''''''''''''''''
'    "String" to bytes:
'    CopyMemory ByteArr(SIndex), ByVal AString$, Len

'    Bytes to "string":
'    AString$ = Space$(Len)
'    CopyMemory ByVal AString$, ByteArr(SIndex), Len
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public aSound As Boolean

Private WAVData$()
Private Sound() As Byte


Public Sub LoadWavs()
Dim U As Long
Dim k As Long
   ReDim WAVData$(0 To 8)
   For k = 0 To 8
      Select Case k
      Case 0: Sound = LoadResData(107, "STEP")
      Case 1: Sound = LoadResData(102, "CHECK")
      Case 2: Sound = LoadResData(103, "CHECKMATE")
      Case 3: Sound = LoadResData(108, "CHECKMAN")
      Case 4: Sound = LoadResData(109, "CHECKMATEMAN")
      Case 5: Sound = LoadResData(110, "HALTED")
      Case 6: Sound = LoadResData(111, "STALEMATE")
      Case 7: Sound = LoadResData(112, "DRAW")
      Case 8: Sound = LoadResData(113, "PROMOTE")
      End Select
      U = UBound(Sound()) + 1
      WAVData$(k) = Space$(U)
      CopyMemory ByVal WAVData$(k), Sound(0), U
   Next k
End Sub
Public Sub Play(Index As Integer)
   If Not aSound Then Exit Sub
   ' Complete play first
   sndPlaySound ByVal WAVData$(Index), SND_NODEFAULT Or SND_MEMORY 'Or SND_NOSTOP
'   DoEvents
End Sub

Public Sub StopPlay()
  ' sndPlaySound 0, SND_PURGE Or SND_ASYNC
   PlayResWAV 0, SND_PURGE
End Sub
