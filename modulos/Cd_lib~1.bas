Attribute VB_Name = "CD_Library"
'
' MCI BASED CD PLAYER CONTROL LIBRARY.
' BY MARK WILLS - EMAIL : mark@stratoblaster.screaming.net
' IF YOU FIND THIS LIBRARY USEFUL, THEN PLEASE DROP ME AN EMAIL!
' Release 1.1 - 9th June 2000'
' NOTE - THIS LIBRARY WAS DEVELOPED IN VB 5, BUT, TO MY KNOWLEDGE, IT SHOULD RUN IN
' ANY 32BIT VERSION, SUCH AS VB 4/32
' [ MARK WILLS IS A 29 YEAR OLD FORK LIFT TRUCK DRIVER FROM THE UK ]
' Revision History:
' Rel : 1.0 - Worked ok. However, if functions were called without having opened the CD
'             player, or, if the cd player failed to open caused runtime errors.
' Rel : 1.1 - It's up to you the programmer to check that the CD opened properly,
'             but seeing as all programmers are lazy, (!) I have added a lot of bomb
'             proofing. As a result, you should never get a runtime error from the
'             library, even if you call functions without first opening the cd player.
'             If you DO get any runtime errors FROM THE LIBRARY, these can be considered
'             bugs, and as such, reported to me at the above email address.
'             Pause Function also added.
'
'declare our windows based DLL functions that we'll be using (only mciSendString)
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public mciOpen As Boolean ' global variable used by the library.
Public drawer As Boolean

'CloseCD
' This function closes the logical connection to the CD Player control library.
' A bit like closing a file when you have finished with it.
' Do not confuse this with closing the drawer on the cd player!
' The ClosePlayer function always returns true
Function CloseCD() As Boolean
mciOpen = False
CloseCD = Not CBool(mciSendString("close cdr", vbNullString, 0, 0))
drawer = False
End Function

'OpenCD
' This function creates the CD alias.
' Pass the drive that you wish to open as a CDAudio drive as a string.
' eg: OpenCD("D:\")
' Only the FIRST letter of the drive name counted. Ie, if you pass "D:\temp" (which is
' invalid anyway) the function will truncate it to "D:\"
' Note, the function does not check to check that the drive you are attempting to open
' is of the appropriate type, so check before you try to open A:\ as a CDAudio!
' This function sets the CD Time format to Track:Mins:Secs:Frames
' True is returned if the CD is succesfully opened.
' None of the functions will operate until the device is opened.
Function OpenCD(CD_Drive As String) As Boolean
Dim ReturnString As String * 30
CloseCD
CD_Drive = Mid$(CD_Drive, 1, 1) + ":\"
OpenCD = Not CBool(mciSendString("open " + CD_Drive + " Type cdaudio alias cdr wait shareable", ReturnString, Len(ReturnString), 0))
If OpenCD Then mciOpen = True Else mciOpen = False
mciSendString "set cdr time format tmsf wait", vbNullString, 0, 0
drawer = False
End Function

'MediaPresent
' This function checks to see if a CD is actually inserted in the selected CD Drive.
' Returns true is a CD is present, else false.
Function MediaPresent() As Boolean
Dim ReturnString As String * 30
MediaPresent = False
If mciOpen Then
    mciSendString "status cdr media present", ReturnString, Len(ReturnString), 0
    MediaPresent = CBool(ReturnString)
End If
End Function

'GetCDID
' This function returns the CDID code - a code that uniquely identifies every audio CD.
' The Windows CD Player uses this code to identify exactly which CD has been inserted
' into the player, so that it can give you the correct album title etc. With access to
' the CDid information, it would be very easy to build your own CD player application
' with a built in Album/Track data base.
Function GetCDID() As String
Dim ReturnString As String, i As Integer
ReturnString = Space$(64)
If mciOpen Then
    mciSendString "info cdr identity", ReturnString, 64, 0
    GetCDID = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
    Else
    GetCDID = "device_not_open"
End If
End Function

'GetNumberOfTracks
' This function will return the number of audio tracks on a CD, in your selected and
' opened CD drive. The value returned is an Integer.
' Note: If 0 is returned the disc is either damaged or the wrong format (or not present)
' If 1 is returned, be aware that the disc MAY be a CD Rom type disc.
Function GetNumberOfTracks() As Integer
Dim ReturnString As String * 30
If MediaPresent Then
    mciSendString "status cdr number of tracks wait", ReturnString, Len(ReturnString), 0
    GetNumberOfTracks = CInt(Mid$(ReturnString, 1, 2))
    Else
    GetNumberOfTracks = 0
End If
End Function

'GetCDLength
' This function returns the length of the CD, as string, formatted MM:SS:FF
' If no disc is detected, then "no_disc_present" is returned. Your software can
' test for this.
Function GetCDLength() As String
Dim ReturnString As String * 30
If mciOpen Then
    If MediaPresent Then
        mciSendString "status cdr length wait", ReturnString, Len(ReturnString), 0
        GetCDLength = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
        Else
        GetCDLength = "no_disc_present"
    End If
    Else
    GetCDLength = "device_not_open"
End If
End Function

'GetCDStatus
' This function returns the current state of the cd player as a string.
' If the device was not opened, then "device_not_open" is returned.
Function GetCDStatus() As String
Dim ReturnString As String * 30
If mciOpen Then
    mciSendString "status cdr mode", ReturnString, Len(ReturnString), 0
    GetCDStatus = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
    Else
    GetCDStatus = "device_not_open"
End If
End Function

'SetCurrentTrack
' This function will allow you to set the current track to play. It does not play the
' track, it just selects the track. Returns a TRUE if successful.
' The function checks that the track number that you pass to it is within legal bounds.
' If not, FALSE is returned.
Function SetCurrentTrack(TrackNumber As Long) As Boolean
If (TrackNumber <= CLng(GetNumberOfTracks)) And (TrackNumber > 0&) Then
    If Not CBool(mciSendString("seek cdr to " & TrackNumber, vbNullString, 0, 0)) Then
        SetCurrentTrack = True
        Else
        SetCurrentTrack = False
    End If
    Else
    SetCurrentTrack = False
End If
End Function

'SetCurrentTime
' This function will set the CD player to a specified point in time, on the specified
' track. The function takes a string argument, in the form tt:mm:ss:ff
' eg SetCurrentTime("04:03:10:50") will start the CD player playing track 4, at 3 mins,
' 10 secs, and 50 frames.
' The Function returns TRUE if successful, or else returns FALSE
' Note: The CD playe should be STOPPED before calling this routine, or else FALSE is
' returned.
Function SetCurrentTime(TimePoint As String) As Boolean
If Len(TimePoint) <> 11 Then SetCurrentTime = False
If IsPlaying Then SetCurrentTime = False: Exit Function
If Not CBool(mciSendString("seek cdr to " + TimePoint, vbNullString, 0, 0)) Then
    SetCurrentTime = True
    Else
    SetCurrentTime = False
End If
End Function

'GetTrackLength
' This function returns the length of the selected track as a string, in the format
' mm:ss:ff
' Use GetNumberOfTracks to determine how many tracks are available.
Function GetTrackLength(Track As Long) As String
Dim ReturnString As String * 30
If mciOpen Then
    If GetNumberOfTracks > 0& Then
        mciSendString "status cdr length track " & Val(Track), ReturnString, Len(ReturnString), 0
        GetTrackLength = Mid$(ReturnString, 1, 8)
        Else
        GetTrackLength = "no_tracks"
    End If
    Else
    GetTrackLength = "device_not_open"
End If
End Function

'GetCurrentPosition
' This function returns a string, containing the current position of the CD player,
' in the format tt:mm:ss:ff (track:minutes:seconds:frames)
' Note that the time returned is relative to the beginning of the selected track.
' If the CD is stopped, the string 01:00:00:00 is returned.
' If no disc is present, then "no_disc_present" is returned. Your software can test for
' this.
Function GetCurrentPosition() As String
Dim ReturnString As String * 30, Status As String
GetCurrentPosition = "device_not_open"
If mciOpen Then
    If MediaPresent Then
        mciSendString "status cdr position", ReturnString, Len(ReturnString), 0
        GetCurrentPosition = Mid$(ReturnString, 1, InStr(1, ReturnString, Chr$(0)) - 1)
        Else
        GetCurrentPosition = "no_disc_present"
    End If
End If
End Function

'PlayCD
' This functions starts the CD playing.
' Returns TRUE if successful.
Function PlayCD() As Boolean
PlayCD = Not CBool(mciSendString("play cdr", vbNullString, 0, 0))
End Function

'StopCD
' This function stops the CD player, and returns the player to beginning of track 1,
' like a conventional CD player.
' Returns TRUE if successful.
Function StopCD() As Boolean
StopCD = Not CBool(mciSendString("stop cdr wait", vbNullString, 0, 0))
SetCurrentTrack (1)
End Function

'PauseCD
' Emulates the pause function of a conventional CD player.
' Returns TRUE if successful.
Function PauseCD() As Boolean
PauseCD = Not CBool(mciSendString("stop cdr wait", vbNullString, 0, 0))
End Function

'EjectCD
' Opens the CD drawer, so you can load a disc.
Function EjectCD() As Boolean
EjectCD = False
If mciOpen Then
    mciSendString "set cdr door open", vbNullString, 0, 0
    EjectCD = True
    drawer = True
End If
End Function

'ShutCD
' Shuts the CD drawer.
Function ShutCD() As Boolean
ShutCD = False
If mciOpen Then
    mciSendString "set cdr door closed", vbNullString, 0, 0
    ShutCD = True
    drawer = False
End If
End Function

'IsStopped
' A general purpose function, provided for convenience. Returns true is the CD
' is stopped.
' Note, an IsStopped and IsPlaying function is provided, which are essentially the
' same, except the logic is reversed. This allows you to choose the function that is
' most applicable the context in which you want to 'ask the question'
' Ie : If you wanted to know if the CD player is playing, you could use IsStopped,
' and test for a FALSE, however, it is nicer (for the programmer!) to use the IsPlaying
' function, which keeps the context correct.
' Some of the above functions use these, so don't delete them!
Function IsStopped() As Boolean
IsStopped = False
If mciOpen Then IsStopped = CBool(InStr(1, GetCDStatus, "stopped"))
End Function
Function IsPlaying() As Boolean
IsPlaying = False
If mciOpen Then IsPlaying = CBool(InStr(1, GetCDStatus, "playing"))
End Function

'Function SETCDVOLUME(volumen As Integer) As String
' Dim cmdToDo As String
' Dim dwReturn As Long
' Dim ret As String * 128
' If mciOpen Then
'
'    Rem cmdToDo = "set cdr volume to 0"
'    dwReturn = mciSendString("set wave_out volume to 0", vbNullString, 0, 0)
'
'  If Not dwReturn = 0 Then  'not success
'      mciGetErrorString dwReturn, ret, 128 'Get the error
'      SETCDVOLUME = ret
'      Exit Function
'  End If
' End If
'End Function

Function LoadCDBase(Lista As Object, ID As String, Titulo As Object, Interpre As Object, mode As Integer) As Boolean
 LoadCDBase = False
 Dim artista As String
 Dim check As Boolean
 Dim POSIT As Integer
 Dim a$
 check = False
 POSIT = 0
 If Dir("c:\windows\cdplayer.ini", vbArchive) = "" Then Exit Function
 Open "c:\windows\cdplayer.ini" For Input As #1
 Do
  Line Input #1, a$
  If InStr(a$, "[" & Hex(ID) & "]") > 0 Then
   check = True
   LoadCDBase = True
  End If
  If check = True Then
   If InStr(a$, "order=") > 0 Then Exit Do
   Select Case POSIT
   Case Is = 2
    If mode = 1 Then
     artista = right(a$, Len(a$) - 7)
    Else
     Interpre = right(a$, Len(a$) - 7)
    End If
   Case Is = 3
    If mode = 2 Then Titulo = right(a$, Len(a$) - 6)
   Case Is >= 5
    If Mid(a$, 1, 1) = "[" Then Exit Do
    If mode = 1 Then
     Lista.AddItem (Mid(a$, InStr(a$, "=") + 1, Len(a$) - InStr(a$, "="))) + " - " + artista
    Else
     Lista.AddItem (Mid(a$, InStr(a$, "=") + 1, Len(a$) - InStr(a$, "=")))
    End If
   End Select
   POSIT = POSIT + 1
  End If
 Loop While Not (EOF(1))
 Close #1
End Function
