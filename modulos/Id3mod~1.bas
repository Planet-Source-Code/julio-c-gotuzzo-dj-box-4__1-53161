Attribute VB_Name = "Id3Module"
Public Type Id3                 'This type is standard for
Title As String * 30            ' Id3 Tags
Artist As String * 30           ' Although later versions
Album As String * 30            ' use comments for 28 bytes
sYear  As String * 4            ' and they use the 2 remaining  bytes for "TrackNumber"!
Comments As String * 30
Genre As Byte
End Type

Public id3info As Id3           ' Declare a variable as the id3 type

Public Function GetId3(Filename As String) As Boolean
Dim Tag As String * 3               ' We use this variable to make sure the file has an ID3TAG
id3info.Album = ""
id3info.Artist = ""
id3info.Comments = ""
id3info.Genre = 0
id3info.sYear = ""
id3info.Title = ""
Open Filename For Binary As #1      ' we open the file as binary for total control (we need it for the Genre part)
Get #1, FileLen(Filename) - 127, Tag    ' Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
If Tag = "TAG" Then                     ' "TAG" is put at position filesize-127 to show that this file indeed contains an Id3
Get #1, FileLen(Filename) - 124, id3info    ' if the file has a tag, we put it into our earlier declared variable id3info
Else
 GetId3 = False
End If
Close #1                                            ' close the file
 GetId3 = True
End Function

Public Function SaveId3(Filename As String, MP3Info As Id3) As Boolean
Dim Tag As String * 3               ' We use this variable to make sure the file has an ID3TAG
SaveId3 = True
On Error GoTo MAL
Open Filename For Binary As #1      ' we open the file as binary for total control (we need it for the Genre part)
Get #1, FileLen(Filename) - 127, Tag    ' Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
If Tag = "TAG" Then                     ' "TAG" is put at position filesize-127 to show that this file indeed contains an Id3
Put #1, FileLen(Filename) - 124, MP3Info    ' if the file has a tag, we put our new information in the file
Else
Put #1, FileLen(Filename) - 127, "TAG"      ' else we put the "TAG" there first,
Close #1
Call SaveId3(Filename, MP3Info)                               ' then we call this function again so we fill the info this time
End If
Close #1                                            ' close the filE
Exit Function
MAL:
 Close #1                                            ' close the filE
 SaveId3 = False
End Function

Public Function GetGenre(Numero As Byte, Reverse As Boolean) As Byte
 If Reverse = True Then
  Select Case Numero
  Case Is = 82
   GetGenre = 18
  Case Is = 80
   GetGenre = 18
  Case Is = 115
   GetGenre = 18
  Case Is = 52
   GetGenre = 17
  Case Is = 4
   GetGenre = 16
  Case Is = 22
   GetGenre = 15
  Case Is = 3
   GetGenre = 14
  Case Is = 2
   GetGenre = 10
  Case Is = 0
   GetGenre = 6
  Case Is = 99
   GetGenre = 1
  Case Is = 40
   GetGenre = 45
  Case Is = 20
   GetGenre = 2
  Case Is = 116
   GetGenre = 3
  Case Is = 135
   GetGenre = 4
  Case Is = 138
   GetGenre = 5
  Case Is = 32
   GetGenre = 9
  Case Is = 68
   GetGenre = 41
  Case Is = 16
   GetGenre = 42
  Case Is = 76
   GetGenre = 43
  Case Is = 17
   GetGenre = 44
  Case Is = 143
   GetGenre = 47
  Case Is = 114
   GetGenre = 48
  Case Is = 5
   GetGenre = 19
  Case Is = 137
   GetGenre = 21
  Case Is = 7
   GetGenre = 22
  Case Is = 35
   GetGenre = 23
  Case Is = 33
   GetGenre = 24
  Case Is = 8
   GetGenre = 25
  Case Is = 86
   GetGenre = 26
  Case Is = 142
   GetGenre = 30
  Case Is = 9
   GetGenre = 31
  Case Is = 10
   GetGenre = 32
  Case Is = 11
   GetGenre = 33
  Case Is = 103
   GetGenre = 34
  Case Is = 12
   GetGenre = 35
  Case Is = 75
   GetGenre = 37
  Case Is = 13
   GetGenre = 38
  Case Is = 109
   GetGenre = 20
  Case Is = 43
   GetGenre = 39
  Case Is = 15
   GetGenre = 40
  Case Is = 42
   GetGenre = 51
  Case Is = 24
   GetGenre = 55
  Case Is = 83
   GetGenre = 52
  Case Is = 106
   GetGenre = 50
  Case Is = 113
   GetGenre = 53
  Case Is = 18
   GetGenre = 54
  Case Is = 144
   GetGenre = 57
  Case Is = 31
   GetGenre = 56
  Case Is = 28
   GetGenre = 59
  Case Is = 200
   GetGenre = 7
  Case Is = 201
   GetGenre = 8
  Case Is = 202
   GetGenre = 11
  Case Is = 203
   GetGenre = 12
  Case Is = 204
   GetGenre = 13
  Case Is = 205
   GetGenre = 27
  Case Is = 206
   GetGenre = 28
  Case Is = 207
   GetGenre = 29
  Case Is = 208
   GetGenre = 36
  Case Is = 209
   GetGenre = 46
  Case Is = 210
   GetGenre = 58
  Case Is = 211
   GetGenre = 60
  Case Else
   GetGenre = 49
  End Select
 Else
  Select Case Numero
  Case Is = 18
   GetGenre = 82
  Case Is = 18
   GetGenre = 80
  Case Is = 18
   GetGenre = 115
  Case Is = 17
   GetGenre = 52
  Case Is = 16
   GetGenre = 4
  Case Is = 15
   GetGenre = 22
  Case Is = 14
   GetGenre = 3
  Case Is = 10
   GetGenre = 2
  Case Is = 6
   GetGenre = 0
  Case Is = 1
   GetGenre = 99
  Case Is = 45
   GetGenre = 40
  Case Is = 2
   GetGenre = 20
  Case Is = 3
   GetGenre = 116
  Case Is = 4
   GetGenre = 135
  Case Is = 5
   GetGenre = 138
  Case Is = 9
   GetGenre = 32
  Case Is = 41
   GetGenre = 68
  Case Is = 42
   GetGenre = 16
  Case Is = 43
   GetGenre = 76
  Case Is = 44
   GetGenre = 17
  Case Is = 47
   GetGenre = 143
  Case Is = 48
   GetGenre = 114
  Case Is = 19
   GetGenre = 5
  Case Is = 21
   GetGenre = 137
  Case Is = 22
   GetGenre = 7
  Case Is = 23
   GetGenre = 35
  Case Is = 24
   GetGenre = 33
  Case Is = 25
   GetGenre = 8
  Case Is = 26
   GetGenre = 86
  Case Is = 30
   GetGenre = 142
  Case Is = 31
   GetGenre = 9
  Case Is = 32
   GetGenre = 10
  Case Is = 33
   GetGenre = 11
  Case Is = 34
   GetGenre = 103
  Case Is = 35
   GetGenre = 12
  Case Is = 37
   GetGenre = 75
  Case Is = 38
   GetGenre = 13
  Case Is = 20
   GetGenre = 109
  Case Is = 39
   GetGenre = 43
  Case Is = 40
   GetGenre = 15
  Case Is = 51
   GetGenre = 42
  Case Is = 55
   GetGenre = 24
  Case Is = 52
   GetGenre = 83
  Case Is = 50
   GetGenre = 106
  Case Is = 53
   GetGenre = 113
  Case Is = 54
   GetGenre = 18
  Case Is = 57
   GetGenre = 144
  Case Is = 56
   GetGenre = 31
  Case Is = 59
   GetGenre = 28
  Case Is = 7
   GetGenre = 200
  Case Is = 8
   GetGenre = 201
  Case Is = 11
   GetGenre = 202
  Case Is = 12
   GetGenre = 203
  Case Is = 13
   GetGenre = 204
  Case Is = 27
   GetGenre = 205
  Case Is = 28
   GetGenre = 206
  Case Is = 29
   GetGenre = 207
  Case Is = 36
   GetGenre = 208
  Case Is = 46
   GetGenre = 209
  Case Is = 58
   GetGenre = 210
  Case Is = 60
   GetGenre = 211
  End Select
 End If
End Function

Public Function GenreStr(Numero As Byte) As String
  Select Case Numero
  Case Is = 82
   GenreStr = "Folklore"
  Case Is = 80
   GenreStr = "Folklore"
  Case Is = 115
   GenreStr = "Folklore"
  Case Is = 52
   GenreStr = "Electrónica"
  Case Is = 4
   GenreStr = "Disco"
  Case Is = 22
   GenreStr = "Death Metal"
  Case Is = 3
   GenreStr = "Dance"
  Case Is = 2
   GenreStr = "Country"
  Case Is = 0
   GenreStr = "Blues"
  Case Is = 99
   GenreStr = "Acústicos"
  Case Is = 40
   GenreStr = "Rock Alternativo"
  Case Is = 20
   GenreStr = "Música Alternativa"
  Case Is = 116
   GenreStr = "Baladas"
  Case Is = 135
   GenreStr = "Ritmo Beat"
  Case Is = 138
   GenreStr = "Black Metal"
  Case Is = 32
   GenreStr = "Música Clásica"
  Case Is = 68
   GenreStr = "Rave"
  Case Is = 16
   GenreStr = "Reggae"
  Case Is = 76
   GenreStr = "Retro"
  Case Is = 17
   GenreStr = "Rock"
  Case Is = 143
   GenreStr = "Salsa"
  Case Is = 114
   GenreStr = "Samba"
  Case Is = 5
   GenreStr = "Funk"
  Case Is = 137
   GenreStr = "Heavy Metal"
  Case Is = 7
   GenreStr = "Hip-Hop"
  Case Is = 35
   GenreStr = "House"
  Case Is = 33
   GenreStr = "Música Intrumental"
  Case Is = 8
   GenreStr = "Jazz"
  Case Is = 86
   GenreStr = "Temas Latinos"
  Case Is = 142
   GenreStr = "Merengue"
  Case Is = 9
   GenreStr = "Metal"
  Case Is = 10
   GenreStr = "New Age"
  Case Is = 11
   GenreStr = "Oldies"
  Case Is = 103
   GenreStr = "Opera"
  Case Is = 12
   GenreStr = "Temas Varios"
  Case Is = 75
   GenreStr = "Polka"
  Case Is = 13
   GenreStr = "Pop"
  Case Is = 109
   GenreStr = "Groove"
  Case Is = 43
   GenreStr = "Punk"
  Case Is = 15
   GenreStr = "Rap"
  Case Is = 42
   GenreStr = "Soul"
  Case Is = 24
   GenreStr = "Temas De Películas"
  Case Is = 83
   GenreStr = "Swing"
  Case Is = 106
   GenreStr = "Sinfonías"
  Case Is = 113
   GenreStr = "Tango"
  Case Is = 18
   GenreStr = "Techno"
  Case Is = 144
   GenreStr = "Trash Metal"
  Case Is = 31
   GenreStr = "Trance"
  Case Is = 28
   GenreStr = "Vocal"
  Case Is = 200
   GenreStr = "Boleros"
  Case Is = 201
   GenreStr = "Temas De Boliche"
  Case Is = 202
   GenreStr = "Cuarteto"
  Case Is = 203
   GenreStr = "Cumbias"
  Case Is = 204
   GenreStr = "Cumbia Villera"
  Case Is = 205
   GenreStr = "Lentos"
  Case Is = 206
   GenreStr = "Mambo"
  Case Is = 207
   GenreStr = "Melódicos"
  Case Is = 208
   GenreStr = "Sólo Piano"
  Case Is = 209
   GenreStr = "Rock Nacional"
  Case Is = 210
   GenreStr = "Música Tropical"
  Case Is = 211
   GenreStr = "Waltz"
  Case Else
   GenreStr = "-1"
  End Select
End Function
