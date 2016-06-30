Module Read_Reaper_Song

    ' Seqentially read through a REAPER song file line by line and extract the song
    ' and markers and the tracks and clips and all relevant properties and  settings 

    Dim curSong As Song, curTrack As Track, curClip As Clip
    Dim isTrack As Boolean, isItem As Boolean, trackCompleted As Boolean

    Public Sub Read_Reaper_Song_Info()

        ' open a file dialog to open a Reaper file
        Dim f As New OpenFileDialog
        f.InitialDirectory = ""
        f.Filter = "Reaper Project (*.RPP)|*.RPP"
        f.Title = "Open a Reaper project file..."
        f.ShowDialog()

        ' if dialog cancelled then exit
        If My.Computer.FileSystem.FileExists(f.FileName) = False Then Exit Sub

        ' read the Reaper file into a string array
        Dim SongFile = IO.File.ReadAllLines(f.FileName)
        'For i = 0 To SongFile.Count - 1
        
        ' init a new song class object
        curSong = New Song
        curSong.ChildTracks = New List(Of Track)
        curSong.Markers = New List(Of String)

        ' get the root path of the song in case you need it
        ' later for audio file path references
        curSong.Name = Replace(f.SafeFileName, ".RPP", "")
        curSong.AudioFilePath = Replace(f.FileName, "\" & f.SafeFileName, "")

        '***********************************************************************
        '            READ THE REAPER RPP SONG FILE LINE BY LINE
        '***********************************************************************
        For I = 0 To SongFile.Count - 1
            If InStr(SongFile(I), "TEMPO ") > 0 AndAlso isTrack = False Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curSong.Tempo = data(1)
                curSong.TimeSig = data(2) & "/" & data(3)
            End If

            If InStr(SongFile(I), "MARKER ") > 0 Then
                Dim data = Split(Trim(SongFile(I)), " ")
                data(3) = Replace(data(3), Chr(34), "")
                If Trim(data(3)) = "" Then data(3) = "No Name"
                curSong.Markers.Add(data(2) & " " & data(3))
            End If

            If InStr(SongFile(I), "SAMPLERATE ") > 0 AndAlso isTrack = False Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curSong.SampleRate = data(1)
            End If

            '*****************************************************************
            '            READ THE DATA FOR THE CURRENT TRACK
            '*****************************************************************
            ' get all the tracks and clips
            If InStr(SongFile(I), "<TRACK") > 0 AndAlso isTrack = False Then

                ' when you enter a <TRACK tag, push the previous track object
                 if any, into the song tracks list array
                Try
                    If curTrack.Name > "" Then curSong.ChildTracks.Add(curTrack)
                Catch
                End Try

                ' set the isTrack flag to know what section we're in
                isTrack = True

                ' init a new Track class object
                curTrack = New Track
                
                ' init the List(Of Clip) array for the ChildClips property
                curTrack.ChildClips = New List(Of Clip)
                
                ' create a new UID for each track
                curTrack.TrackID = UCase(Guid.NewGuid.ToString)

                Dim name = Replace(SongFile(I + 1), "NAME", "")

                ' use a method to format the name and fix empty track names
                curTrack.Name = curSong.FormatTrackName(name)
            End If

            ' track color  (optional) get and convert to HTML
            If InStr(SongFile(I), "PEAKCOL") > 0 AndAlso isTrack = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                Dim tcolor = Split(Trim(data(1)), " ")
                Dim trackColor = ColorTranslator.FromOle(tcolor(0))
                curTrack.Color = UCase(trackColor.Name)
            End If

            If InStr(SongFile(I), "VOLPAN") > 0 AndAlso isTrack = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curTrack.Volume = data(1)

                '****************************************************
                '  Reaper's pan ranges from  Left -1, Center 0, Right 
                '  keep that in mind when converting pan values to
                '  other products where the ranges may be different
                '****************************************************
                curTrack.Pan = data(2)
            End If

            If InStr(SongFile(I), "MUTESOLO") > 0 AndAlso isTrack = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curTrack.Mute = data(1)
            End If

            '*****************************************************************
            '            READ THE DATA FOR THE CURRENT ITEM (CLIP)
            '*****************************************************************
            ' read the clips beneath the current track
            If InStr(SongFile(I), "<ITEM") > 0 Then
                curClip = New Clip

                ' switch the boolean flags so we know we're in an <ITEM section
                isItem = True
                isTrack = False

                ' assoicate later if necessary by using the parent track UID
                curClip.ParentTrack = curTrack.TrackID
            End If

            If InStr(SongFile(I), "POSITION") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.Start = data(1)
            End If

            If InStr(SongFile(I), "LENGTH") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.Length = data(1)
            End If

            If InStr(SongFile(I), "FADEIN") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.FadeIn = data(2)
            End If

            If InStr(SongFile(I), "FADEOUT") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.FadeOut = data(2)
            End If

            If InStr(SongFile(I), "MUTE") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.Mute = data(1)
            End If

            If InStr(SongFile(I), "NAME") > 0 AndAlso isItem = True Then
                Dim name = Split(Trim(SongFile(I)), " ")
                curClip.Name = CheckName(name(1))
            End If

            If InStr(SongFile(I), "VOLPAN") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.Gain = data(1)
            End If

            If InStr(SongFile(I), "SOFFS") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I)), " ")
                curClip.Offset = data(1)
            End If

            If InStr(SongFile(I), "PLAYRATE ") > 0 AndAlso isItem = True Then
                Dim Data = Split(Trim(SongFile(I)), " ")
                curClip.Speed = Trim(Data(1))

                ' Reaper uses a siNgle value for pitch so we have to split
                ' it to get both Transpose and Tune values
                Dim FullPitch = Data(3)

                If InStr(FullPitch, ".") > 0 Then
                    Dim ary = Split(FullPitch, ".")
                    ' transpose
                    Dim pitch = ary(0)

                    Dim tune = CInt(("0." & ary(1)) * 100).ToString

                    curClip.Pitch = pitch.ToString
                    curClip.Tune = tune.ToString
                Else
                    curClip.Pitch = FullPitch
                    curClip.Tune = "0"
                End If
            End If

            '******************************************************************
            '      WE KNOW THIS TAG IS AT THE END OF EVERY <ITEM SECTION
            '******************************************************************
            If InStr(SongFile(I), "<SOURCE WAVE") > 0 AndAlso isItem = True Then
                Dim data = Split(Trim(SongFile(I + 1)), "FILE ")
                curClip.FileName = Replace(data(1), Chr(34), "")

                ' If it then it's a short filename use the Reaper root song folder path
                ' to build a full filename string
                If InStr(curClip.FileName, ":\") = 0 Then
                    curClip.FileName = curSong.AudioFilePath & "\" & curClip.FileName
                End If

                ' add this clip to the current track
                curTrack.ChildClips.Add(curClip)
                'curClip = Nothing

                'Switch off the flags
                isItem = False
                isTrack = False

            End If

            '***********************************************************************
            ' when you hit the <EXTENSIONS tag you know there are no more tracks
            ' push the last track object into the song list array and exit the loop
            '***********************************************************************
            If InStr(SongFile(I), "<EXTENSIONS") > 0 Then
                Try
                    curSong.ChildTracks.Add(curTrack)
                Catch
                End Try
                Exit For
            End If
        Next
        
        '***********************************************************
        ' push all of the data into public loadedSong class object
        ' curSong is local to this module and will lose scope
        '***********************************************************
        loadedSong = curSong

        '*************************************************************************
        '     OPTIONALLY TEST PRINT THE SONG DATA BY SENDING IT TO THE CONSOLE
        '*************************************************************************
        Exit Sub  ' comment to print the results
        loadedSong.PrintSongData()

    End Sub

End Module
