' This is a generic song class module that uses auto-implemented properites.
' Because the destinations are xml or text, most properties are String()

Public Class Song

    Public EmptyName As Integer = 0
    Public Property Name As String
    Public Property ChildTracks As List(Of Track)
    Public Property SongMarkers As List(Of String)
    Public Property SampleRate As String
    Public Property Tempo As String
    Public Property TimeSig As String
    Public Property Numerator As Integer
    Public Property Denominator As Integer
    Public Property Source() As String()
    Public Property TrackCount As String
    Public Property MasterBusID As String
    Public Property AudioFilePath As String
    Public Property Notes As String
    Public Property Markers() As New List(Of String)

    Public Sub New()
    End Sub
    ' Song Methods
    Public Function FormatTrackName(name As String)
    
        ' strip any quote chars
        name = Replace(Trim(name), Chr(34), "")

        If Trim(name) = "" Then
            EmptyName = EmptyName + 1
            name = "[Track] " & EmptyName
            Return name
        End If

        ' format and make name title case
        Dim myTI As Globalization.TextInfo = New Globalization.CultureInfo("en-US", False).TextInfo
        name = myTI.ToTitleCase(name)
        Return name
    End Function

    '*************************************************************************************
    '           A METHOD TO PRINT TO CONSOLE A NICELY FORMATTED OVERVIEW 
    '               OF THE DATA IN THE LOADEDSONG SONG CLASS INSTANCE
    '*************************************************************************************

    Public Sub PrintSongData()
        Dim T As Integer = 1, C As Integer = 1

        With loadedSong
            Console.WriteLine("SONG  **************************************************")
            Console.WriteLine("     Name             " & .Name)
            Console.WriteLine("     Tempo            " & .Tempo)
            Console.WriteLine("     Time Sig         " & .TimeSig)
            Console.WriteLine("     Sample Rate:     " & .SampleRate)
            Console.WriteLine("     Markers          " & .Markers.Count)
            Console.WriteLine("     Notes:           " & .Notes)
            Console.WriteLine("")

            Console.WriteLine("MARKERS  ----------------------------------------------")

            For i = 0 To .Markers.Count - 1
                Console.WriteLine("     Pos / Name:  " & .Markers(i))
            Next

            Console.WriteLine("")

            For Track = 0 To .ChildTracks.Count - 1
                Console.WriteLine("TRACK " & T & "  ==============================================")
                Console.WriteLine("     Name:    " & .ChildTracks(Track).Name)
                Console.WriteLine("     ID:      " & .ChildTracks(Track).TrackID)
                Console.WriteLine("     Color    " & .ChildTracks(Track).Color)
                Console.WriteLine("     Volume:  " & .ChildTracks(Track).Volume)
                Console.WriteLine("     Pan      " & .ChildTracks(Track).Pan)
                Console.WriteLine("     Mute     " & .ChildTracks(Track).Mute)
                Console.WriteLine("     Clips     " & .ChildTracks(Track).ChildClips.Count)

                With loadedSong.ChildTracks(Track)
                    For Clip = 0 To .ChildClips.Count - 1
                        Console.WriteLine("             ------------  Clip " & C & " --------------------")
                        Console.WriteLine("             Name:  " & .ChildClips(Clip).Name)
                        Console.WriteLine("             ID:  " & .ChildClips(Clip).ClipID)
                        Console.WriteLine("             Parent:  " & .ChildClips(Clip).ParentTrack)
                        Console.WriteLine("             Start:  " & .ChildClips(Clip).Start)
                        Console.WriteLine("             Length:  " & .ChildClips(Clip).Length)
                        Console.WriteLine("             Offset:  " & .ChildClips(Clip).Offset)
                        Console.WriteLine("             Gain:  " & .ChildClips(Clip).Gain)
                        Console.WriteLine("             Pitch:  " & .ChildClips(Clip).Pitch)
                        Console.WriteLine("             Tune:  " & .ChildClips(Clip).Tune)
                        Console.WriteLine("             Fade In:  " & .ChildClips(Clip).FadeIn)
                        Console.WriteLine("             Fade Out:  " & .ChildClips(Clip).FadeOut)
                        Console.WriteLine("             Filename:  " & .ChildClips(Clip).FileName)
                        Console.WriteLine("")
                        C = C + 1
                    Next
                End With
                T = T + 1
                C = 1
            Next
        End With
    End Sub

    Public Function createUID()
        Dim ID = UCase(Guid.NewGuid().ToString)
        Return ID



    End Function

    Public Function crossCheckTrackNames(name)
        'For I = 0 To Tracks.Count - 1
        '    If Trim(name) = Trim(Tracks(I).Name) Then
        '        IncrementTrackNames = IncrementTrackNames + 1
        '        name = Trim(name) & " " & IncrementTrackNames
        '        Return name
        '    Else
        '        Return name
        '    End If
        'Next
        Return name
    End Function

End Class

Public Class Marker
    Public Property Position As String
    Public Property Name As String
End Class

'===============================================================

Public Class Track
    Public Property Name As String
    Public Property Channels As String
    Public Property Volume As String
    Public Property Pan As String
    Public Property Color As String
    Public Property Mute As String
    Public Property XMLTrack As String
    Public Property TrackID As String
    Public Property PlayListID As String
    Public Property RefID As String
    Public Property ChildClips() As List(Of Clip)

    Public Sub New()
    End Sub


End Class
Public Class Clip
    Public Property ParentTrack As String
    Public Property Name As String
    Public Property Type As String
    Public Property Start As String
    Public Property Length As String
    Public Property Offset As String
    Public Property Mute As String
    Public Property Gain As String
    Public Property Speed As String
    Public Property Pitch As String
    Public Property Tune As String
    Public Property FadeIn As String
    Public Property FadeOut As String
    Public Property ClipID As String
    Public Property FileName As String
    Public Property PoolID As String
    Public Property XMLClipID As String
    Public Property Origin As String ' path?


    Public Property UseCount As Integer

    Public Sub New()
    End Sub


End Class
