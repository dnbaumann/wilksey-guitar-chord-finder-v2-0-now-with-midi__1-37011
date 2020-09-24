Attribute VB_Name = "MidiDefs"
'Guitar Chord Finder MIDI Playing Module
'©2002 Wilksey
'Designed for version 2.0+
'Following Defs taken from MIDI Piano.
Option Explicit

'Consts
Public Const MAXPNAMELEN = 32
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)
Public Const MIDIERR_BASE = 64
Public Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)
Public Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)
Public Const MIDIERR_BADOPENMODE = (MIDIERR_BASE + 6)

'Types
Public Type MIDIOUTCAPS
    wMid As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
    wTechnology As Integer
    wVoices As Integer
    wNotes As Integer
    wChannelMask As Integer
    dwSupport As Long
End Type

'Declarations
Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
