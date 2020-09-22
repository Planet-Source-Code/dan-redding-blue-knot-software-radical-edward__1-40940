Attribute VB_Name = "modResSound"
Option Explicit


'Public Declare Function waveOutGetNumDevs _
'    Lib "winmm" () As Long
    
Private Declare Function sndPlaySound _
    Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
    
    
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)


'Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SND_ASYNC = &H1   '  play asynchronously
Private Const SND_MEMORY = &H4  '  lpszSoundName points to a memory file
Private Const SND_SYNC = &H0
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const lngFlags As Long = SND_ASYNC Or SND_NODEFAULT

Public Sub PlayResSound(intSndID As Integer, Optional blnSync As Boolean = False)
Dim bSound() As Byte, aSound As String
    Debug.Print "!"
    bSound = LoadResData(intSndID, "WAVE")
    aSound = Space$(UBound(bSound) + 1)
    CopyMemory ByVal aSound, bSound(0), Len(aSound)
    sndPlaySound aSound, IIf(blnSync, SND_SYNC, SND_ASYNC) Or SND_MEMORY
    Erase bSound
End Sub

