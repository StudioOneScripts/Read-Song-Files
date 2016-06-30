' this module needs a bit of cleaning up and refactoring
' and more code commenting.  colors also not working yet

Imports System.Xml

Module Read_Ardour_Song

    ' var used to increment numbers when cross checking track names
    ' with the Song.crossCheckTrackNames() Method
    'Public IncrementTrackNames As Integer = 0
    'Public IncrementAudioChNames As Integer = 100

    Public Sub Read_Ardour_Song()

        Dim TrackNodes As XmlNodeList, ClipNodes As XmlNodeList
        Dim PoolFiles As New List(Of Clip), Tracks As New List(Of Track), Clips As New List(Of Clip)
        Dim AllClips As New List(Of Clip)
        loadedSong = New Song

        '*****************************************************************
        '                  OPRN AN *.AROUR XML SONG FILE
        '*****************************************************************
        Dim f As New OpenFileDialog
        'f.InitialDirectory = Form1.lbl[DAWName]Path.Text
        f.Filter = "Ardour XML (*.ardour)|*.ardour"
        f.Title = "Open a MixBus Ardour Session..."
        f.InitialDirectory = ArdourFolder
        f.ShowDialog()

        ' if OpenFileDialog is cancelled, exit
        If My.Computer.FileSystem.FileExists(f.FileName) = False Then Exit Sub

        ' get the short file name for later
        Dim ShortFileName = Replace(f.SafeFileName, ".ardour", "")

        '*****************************************************************
        '             LOAD SONG FILE INTO AN XML DOCUMENT
        '*****************************************************************

        Dim reader As XmlTextReader = New XmlTextReader(f.FileName)
        Dim songText = My.Computer.FileSystem.ReadAllText(f.FileName)
        Dim songXML As New XmlDocument
        songXML.LoadXml(songText)

        Tracks.Clear()
        Clips.Clear()
        AllClips.Clear()
        PoolFiles.Clear()

        '*************************************************************
        '                 INIT SONG CLASS OBJECT
        '*************************************************************

        Dim curSong As New Song
        curSong.ChildTracks = New List(Of Track)
        curSong.Name = Replace(f.SafeFileName, ".ardour", "")

        ' create a new shared UID for Master and all track connections

        curSong.AudioFilePath = Replace(f.FileName, "\" & f.SafeFileName, "")

        '  Load the 'Tracknodes' XMLNodeList with all <Playlist> nodes
        TrackNodes = songXML.SelectNodes("/Session/Playlists/Playlist")
        curSong.TrackCount = TrackNodes.Count

        '   Set the Song Class Tempo Sample Rate & Time Sig Properties
        Dim curNode = songXML.SelectSingleNode("/Session/TempoMap/Tempo")
        curSong.Tempo = curNode.Attributes("beats-per-minute").Value

        ' Ardour uses sample for placement so this will help the conversion
        Dim BeatPerSecond = (curSong.Tempo / 60)

        curNode = songXML.SelectSingleNode("/Session")
        curSong.SampleRate = curNode.Attributes("sample-rate").Value

        curNode = songXML.SelectSingleNode("/Session/TempoMap/Meter")
        curSong.TimeSig = CInt(curNode.Attributes(
            "divisions-per-bar").Value).ToString & "/" & CInt(
             curNode.Attributes("note-type").Value).ToString

        Dim ts = Split(curSong.TimeSig, "/")
        curSong.Numerator = ts(0)
        curSong.Denominator = ts(1)

        '*************************************************************
        '  Build an audio pool file from the <Sources> tag

        Dim PoolClipCollection As String = ""
        Dim SourceNodes = songXML.SelectNodes("/Session/Sources/Source")

        PoolFiles.Clear()

        '  create unique UID's for the pool clips
        For s = 0 To SourceNodes.Count - 1
            ' create a new clip
            Dim pClip = New Clip
            pClip.Name = SourceNodes(s).Attributes("name").Value
            pClip.ClipID = UCase(Guid.NewGuid.ToString)

            If My.Computer.FileSystem.FileExists(SourceNodes(s).Attributes("origin").Value) Then
                ' if the file is external the full path name will be in the "origin" value
                pClip.FileName = Replace(SourceNodes(s).Attributes("origin").Value, "\", "/")
            Else
                'GoTo startover
                ' use the Ardour \audiofiles folder path when a clip isn't external and doesn't have a full pathname
                pClip.FileName = Replace(curSong.AudioFilePath & "\interchange\" & ShortFileName & "\audiofiles\" & pClip.Name, "\", "/")
            End If
            pClip.XMLClipID = SourceNodes(s).Attributes("id").Value
            PoolFiles.Add(pClip)
        Next
   
        '*************************************************************
        '    Loop through all tracks and clips in the song 
        '    song markers property, which is a string array
        '*************************************************************

        Dim SongList As New ListBox
        Dim Row As Integer = 0
        
        curSong.ChildTracks = New List(Of Track)
       
        '*************************************************************
        '      START READING THE TRACKS AND CLIPS IN A FOR LOOP   
        '************************************************************
        For I = 0 To TrackNodes.Count - 1

            ' create  track
            Dim curTrack = New Track
            curTrack.ChildClips = New List(Of Clip)

            ' Create a unique UID for every new track. Method: createUID()
            curTrack.TrackID = curSong.createUID()

            '  Set current xml node to <Playlist> (I)
            curNode = TrackNodes.ItemOf(I)

            '  Load the ClipNodes XMLNodeList with array of all 
            '  of the <Region> nodes under the current <Playlist>
            ' to reference later when we loop through clips
            ClipNodes = curNode.ChildNodes

            '  Get the reference ID
            curTrack.RefID = curNode.Attributes("orig-track-id").Value

            '  Track Name
            If curNode.Attributes.ItemOf("name") IsNot Nothing Then
                curTrack.Name = curNode.Attributes("name").Value
            Else
                IncrementTrackNames = IncrementTrackNames + 1
                curTrack.Name = "Name: [Track] " & IncrementTrackNames
            End If

            '  Track Mute refID is the ardour track id for cross referencing
            '  things like the volume and pan values
            Dim refID = curTrack.RefID
            curNode = songXML.SelectSingleNode("/Session/Routes/Route[@id=" & refID & "]/Controllable[@name='mute']")
            curTrack.Mute = CInt(curNode.Attributes.ItemOf("value").Value).ToString

            '  Track Fader
            '  refID is the ardour track id for cross referencing
            curNode = songXML.SelectSingleNode("/Session/Routes/Route[@id=" & refID & "]/Processor/Controllable[@name='" & "gaincontrol" & "']")
            curTrack.Volume = curNode.Attributes.ItemOf("value").Value

            curNode = songXML.SelectSingleNode("/Session/Routes/Route[@id=" & refID & "]/Pannable/Controllable[@name='" & "pan-azimuth" & "']")
            curTrack.Pan = curNode.Attributes.ItemOf("value").Value

            '********************************************
            '  Track Color
            '  the RGB codes are decimal, convert to hex
            '********************************************

            'curNode = songXML.SelectSingleNode("/Session/Routes/Route[@id=" & refID & "]/PresentationInfo")
            'Dim clr = Hex(curNode.Attributes.ItemOf("color").Value)
            'curTrack.Color = clr

            '***************************************************
            '      READ THE CLIP BENEATH THE <PLAYLIST> TAG  
            '***************************************************

            For C = 0 To ClipNodes.Count - 1

                ' New clip object instance from Clip Class
                Dim curClip As New Clip

                '  Only extract from <Region> nodes In Ardour Start = Offset / Position = Start
                Select Case ClipNodes(C).Name
                    Case = "Region"
                        curNode = ClipNodes(C)
                        curClip.ParentTrack = curTrack.Name
                        curClip.Name = curNode.Attributes("name").Value
                        curClip.Type = curNode.Attributes("type").Value
                        curClip.Start = (curNode.Attributes("position").Value / curSong.SampleRate) * BeatPerSecond
                        curClip.Length = (curNode.Attributes("length").Value / curSong.SampleRate) '* BeatPerSecond
                        curClip.Offset = (curNode.Attributes("start").Value / curSong.SampleRate) * BeatPerSecond
                        curClip.XMLClipID = curNode.Attributes("master-source-0").Value
                        curClip.Mute = curNode.Attributes("muted").Value
                        curClip.Gain = curNode.Attributes("scale-amplitude").Value
                        curClip.Speed = curNode.Attributes("stretch").Value

                        ' get the fade in/out values
                        If curNode.HasChildNodes = True Then
                            Try
                                Dim data = Split(curNode.FirstChild.NextSibling.FirstChild.FirstChild.InnerText, vbLf)
                                Dim fadein = Split(data(1), " ")
                                curClip.FadeIn = (fadein(0) / curSong.SampleRate) '* BeatPerSecond
                                data = Split(curNode.FirstChild.NextSibling.NextSibling.NextSibling.FirstChild.FirstChild.InnerText, vbLf)
                                Dim fadeout = Split(data(1), " ")
                                curClip.FadeOut = (fadeout(0) / curSong.SampleRate) ' * BeatPerSecond
                            Catch
                            End Try
                        End If

                        '****************  Get use counts
                        For p = 0 To PoolFiles.Count - 1
                            If curClip.XMLClipID = PoolFiles(p).XMLClipID Then
                                curClip.ClipID = PoolFiles(p).ClipID
                                curClip.FileName = PoolFiles(p).FileName
                                PoolFiles(p).UseCount = PoolFiles(p).UseCount + 1
                                Exit For
                            End If
                        Next

                End Select

                '  Add the current Clip to the List(of Clip) array
                Clips.Add(curClip)

                ' push to track array
                curTrack.ChildClips.Add(curClip)
            Next

            ' Add this collection of clips to the AllClips List() before cycle
            For P = 0 To Clips.Count - 1
                AllClips.Add(Clips(P))
            Next

            ' Clear the array before the next cycle
            Clips.Clear()

            '   Add the current track to List(of Track) array
            Tracks.Add(curTrack)

            ' push to the main song object
            curSong.ChildTracks.Add(curTrack)
        Next

        '  Method:  validate and format track names
        For Each Track In Tracks
            Track.Name = curSong.FormatTrackName(Track.Name)
        Next

        '****************************************************************
        '                   GET THE SONG MARKERS 
        '****************************************************************
        curNode = songXML.SelectSingleNode("/Session/Locations/Location")
        curNode = curNode.NextSibling

        Do
            Try
                If curNode.Attributes("flags").Value = "IsMark" Then
                    curSong.Markers.Add((curNode.Attributes("start").Value / curSong.SampleRate) * BeatPerSecond & "|" & curNode.Attributes("name").Value)
                    curNode = curNode.NextSibling
                End If
            Catch
                Exit Do
            End Try

        Loop

        loadedSong.Markers = New List(Of String)
        loadedSong.Markers = curSong.Markers

        loadedSong = curSong

        '*****************************************************
        '        OPTIONALLY PRINT FORMATTED RESULTS
        '*****************************************************
        Exit Sub  ' optionally print data result
        loadedSong.PrintSongData()

    End Sub
End Module

