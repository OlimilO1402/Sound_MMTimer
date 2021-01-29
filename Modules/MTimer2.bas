Attribute VB_Name = "MTimer2"
Option Explicit
'wType
'Time format. It can be one of the following values.
Public Const TIME_BYTES   As Long = &H0&  'Current byte offset from beginning of the file.
Public Const TIME_MS      As Long = &H1&  'Time in milliseconds.
Public Const TIME_SAMPLES As Long = &H2&  'Number of waveform-audio samples.
Public Const TIME_SMPTE   As Long = &H8&  'SMPTE (Society of Motion Picture and Television Engineers) time.
Public Const TIME_MIDI    As Long = &H10& 'MIDI time.
Public Const TIME_TICKS   As Long = &H20& 'Ticks within a MIDI stream.

'https://docs.microsoft.com/en-us/previous-versions//dd757347(v=vs.85)

'typedef struct mmtime_tag {
'  UINT  wType;
'  union {
'    DWORD  ms;
'    DWORD  sample;
'    DWORD  cb;
'    DWORD  ticks;
'    struct {
'      BYTE hour;
'      BYTE min;
'      BYTE sec;
'      BYTE frame;
'      BYTE fps;
'      BYTE dummy;
'      BYTE pad[2];
'    } smpte;
'    struct {
'      DWORD songptrpos;
'    } midi;
'  } u;
'} MMTIME, *PMMTIME, *LPMMTIME;
'Type Smpte
'    hour  As Byte
'    min   As Byte
'    sec   As Byte
'    frame As Byte
'
'    fps   As Byte 'oder einfach fps As Long
'    dummy As Byte
'    pad0  As Byte
'    pad1  As Byte
'End Type
'
'Type Midi
'    songptrpos As Long
'End Type
'
'Type UnionU
'    ms_sample_cb_ticks_smpte_hour_min_sec_frame_midi_SongPtrPos As Long
'    fps As Long
'End Type
'
'Type MMTIME
'    wType As Long
'    u As UnionU
'End Type

Public Type Smpte
    hour  As Byte
    min   As Byte
    sec   As Byte
    frame As Byte
    
    'Frames per second (24, 25, 29 (30 drop), or 30).
    fps   As Byte 'oder einfach nur fps As Long??
    dummy As Byte
    pad0  As Byte
    pad1  As Byte
End Type

Public Type Midi
    songptrpos As Long
End Type

'https://docs.microsoft.com/en-us/previous-versions/ms712191(v=vs.85)
Public Type MMTIME
    wType As Long
    u_ms_sample_cb_ticks_smpte_midi0 As Long
    fps As Long
End Type

'timeGetSystemTime
'https://docs.microsoft.com/en-us/windows/win32/api/timeapi/nf-timeapi-timegetsystemtime
Public Declare Function timeGetSystemTime Lib "winmm.dll" ( _
     ByRef lpTime As MMTIME, _
     ByVal uSize As Long) As Long

'https://docs.microsoft.com/en-us/previous-versions/ms713414(v=vs.85)
Public Type TIMECAPS
    wPeriodMin As Long 'Minimum supported resolution.
    wPeriodMax As Long 'Maximum supported resolution.
End Type
 
