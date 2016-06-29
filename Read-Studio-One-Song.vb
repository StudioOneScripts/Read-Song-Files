Public Sub Read_Studio_One_Song()


        '******************************************************************************************************************
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////   OPEN AND UNZIP A STUDIO ONE SONG FILE TO GET THE DATA   ////////////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '******************************************************************************************************************

        ' uses Ionic.zip dll to unzip the file  https://dotnetzip.codeplex.com/

        Dim f As New OpenFileDialog
        f.InitialDirectory = StudioOneFolder
        f.Filter = "Studio One Song (*.song)|*.song"
        f.Title = "Open a Studio One Song ..."
        f.ShowDialog()
        If My.Computer.FileSystem.FileExists(f.FileName) = False Then Exit Sub

        ' a generic identifier var 
        DAW = "Studio One"

        Form1.lblExportType.Text = ""

        ' remove temp folder and create a new one
        Try
            My.Computer.FileSystem.DeleteDirectory(TempPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Catch ex As Exception
            ' MsgBox("remove temp dir: " & ex.Message)
        End Try

        My.Computer.FileSystem.CreateDirectory(TempPath)


        '******************************************************************************************************************
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////   EXTRACT THE FILES BELOW FROM THE STUDIO ONE SONG FILE   ////////////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '******************************************************************************************************************
        Try
            Using zip As ZipFile = ZipFile.Read(f.FileName)
                For Each item As ZipEntry In zip
                    If item.FileName = "metainfo.xml" Then item.Extract(TempPath, ExtractExistingFileAction.OverwriteSilently)
                    If item.FileName = "Song/song.xml" Then item.Extract(TempPath, ExtractExistingFileAction.OverwriteSilently)
                    If item.FileName = "Song/mediapool.xml" Then item.Extract(TempPath, ExtractExistingFileAction.OverwriteSilently)
                    If item.FileName = "Devices/audiomixer.xml" Then item.Extract(TempPath, ExtractExistingFileAction.OverwriteSilently)
                    If item.FileName = "notes.txt" Then item.Extract(TempPath, ExtractExistingFileAction.OverwriteSilently)
                Next
            End Using
        Catch ex1 As Exception
            MsgBox("zip file exception:    {0}", ex1.ToString)
        End Try

        '******************************************************************************************************************
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////   INSTANCE A NEW SONG CLASS AND PARSE THE METAINFO FILE   ////////////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '******************************************************************************************************************

        Dim curSong As New Song, TotalClips As Integer

        ' load the mediapool.xml file  (needed for clip filenames)
        Dim mediaPool = My.Computer.FileSystem.ReadAllText(TempPath & "\Song\mediapool.xml")
        Dim mediaPoolXML As New XmlDocument
        mediaPoolXML.LoadXml(Replace(mediaPool, "x:", ""))

        ' load the audiomixer.xml file (needed for fader and pan
        Dim audioMixer = My.Computer.FileSystem.ReadAllText(TempPath & "\Devices\audiomixer.xml")
        Dim audioMixerXML As New XmlDocument
        audioMixerXML.LoadXml(Replace(audioMixer, "x:", ""))

        curSong.Name = Replace(f.SafeFileName, ".song", "")

        ' load the metainfo.xml file
        Dim metaInfo = My.Computer.FileSystem.ReadAllText(TempPath & "\metainfo.xml")
        Dim metaXML As New XmlDocument
        metaXML.LoadXml(metaInfo)

        'quote char
        Dim Q = Chr(34)

        ' get the tempo, sample rate and time sig from the metainfo.xml file
        curSong.SampleRate = metaXML.SelectSingleNode("/MetaInformation/Attribute[@id=" & Q & "Media:SampleRate" & Q & "]").Attributes("value").Value
        curSong.Tempo = metaXML.SelectSingleNode("/MetaInformation/Attribute[@id=" & Q & "Media:Tempo" & Q & "]").Attributes("value").Value
        curSong.Numerator = metaXML.SelectSingleNode("/MetaInformation/Attribute[@id=" & Q & "Media:TimeSignatureNumerator" & Q & "]").Attributes("value").Value
        curSong.Denominator = metaXML.SelectSingleNode("/MetaInformation/Attribute[@id=" & Q & "Media:TimeSignatureDenominator" & Q & "]").Attributes("value").Value
        curSong.TimeSig = curSong.Numerator & "/" & curSong.Denominator


        ' get the song notes 
        Try
            curSong.Notes = My.Computer.FileSystem.ReadAllText(TempPath & "\notes.txt")
        Catch
            curSong.Notes = " "
        End Try


        '******************************************************************************************************************
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '//////////////////   GET EVERYTHING ELSE FROM THE SONG.XML FILE:  MARKERS, SONGS AND CLIPS   /////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '******************************************************************************************************************

        ' load the song.xml into an XML document  (also replacing "x:" which is formatting VS has issue with
        Dim songInfo = Replace(My.Computer.FileSystem.ReadAllText(TempPath & "\Song\song.xml"), "x:", "")
        Dim songXML As New XmlDocument
        songXML.LoadXml(songInfo)

        ' get the song markers, skip the first two (Start and End)
        curSong.SongMarkers = New List(Of String)
        Dim Markers = songXML.SelectSingleNode("/Song/Attributes/List/MarkerTrack").ChildNodes

        ' loop through the array and add the markers to the song's property array
        For I = 2 To Markers.Count - 1
            curSong.SongMarkers.Add(Markers(I).Attributes("start").Value & "|" & Markers(I).Attributes("name").Value)
        Next

        '******************************************************************************************************************
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////  READ ALL OF THE TRACKS AND CLIPS FROM THE SONG FILE   /////////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '******************************************************************************************************************

        ' this node list is a collection of all of the matching nodes below, audio tracks
        Dim TrackList As XmlNodeList = songXML.SelectNodes("/Song/Attributes/List/MediaTrack[@mediaType=" & Q & "Audio" & Q & "]")

        ' song.ChildTracks is a List(Of Track) property for all of the tracks
        curSong.ChildTracks = New List(Of Track)

        ' toop through the tracklist node list and get each track
        For i = 0 To TrackList.Count - 1
            Dim curTrack As New Track
            curTrack.ChildClips = New List(Of Clip)
            curTrack.Name = TrackList(i).Attributes("name").Value


            ' get the volume and pan for the curren track by cross
            ' referencing the GUID of the track with the audiomixer.xml

            ' get the channel ID
            curTrack.TrackID = TrackList(i).FirstChild.NextSibling.Attributes("uid").Value

            ' get the reference track from the audiomixer.xml file
            Dim RefTrack As XmlNode = audioMixerXML.SelectSingleNode("/AudioMixer/Attributes/ChannelGroup/AudioTrackChannel/UID[@uid=" & Q & curTrack.TrackID & Q & "]").ParentNode

            Try
                curTrack.Volume = RefTrack.Attributes("gain").Value
            Catch
                curTrack.Volume = "0"
            End Try

            Try
                curTrack.Pan = RefTrack.Attributes("pan").Value
                If Trim(curTrack.Pan) = "" Then curTrack.Pan = "0.5"
            Catch
                curTrack.Pan = "0.5"
            End Try


            Try
                curTrack.Mute = TrackList(i).Attributes("mute").Value
            Catch ex As Exception
                curTrack.Mute = "0"
            End Try

            ' loop through all of the child nodes beneath this
            ' track and get all of the clips on this track
            Dim ChildList As XmlNodeList = TrackList(i).ChildNodes
            For c = 0 To ChildList.Count - 1
                If ChildList(c).Name = "List" Then
                    Dim Eventlist As XmlNodeList = ChildList(c).ChildNodes

                    For aEvent = 0 To Eventlist.Count - 1
                        Dim curClip As New Clip
                        curClip.ParentTrack = curTrack.TrackID
                        Try
                            curClip.Name = Eventlist(aEvent).Attributes("name").Value

                            curClip.ClipID = Eventlist(aEvent).Attributes("clipID").Value
                        Catch
                            GoTo skipclip
                        End Try
                        Try
                            curClip.Mute = Eventlist(aEvent).Attributes("mute").Value
                        Catch
                            curClip.Mute = "0"
                        End Try
                        Try
                            curClip.Start = Eventlist(aEvent).Attributes("start").Value
                        Catch
                            curClip.Start = "0"
                        End Try

                        curClip.Length = Eventlist(aEvent).Attributes("length").Value

                        Try
                            curClip.Offset = Eventlist(aEvent).Attributes("offset").Value
                        Catch
                            curClip.Offset = "0"
                        End Try

                        curClip.Speed = Eventlist(aEvent).Attributes("speed").Value
                        curClip.Pitch = Eventlist(aEvent).Attributes("transpose").Value
                        curClip.Tune = Eventlist(aEvent).Attributes("tune").Value
                        Try
                            curClip.Gain = Eventlist(aEvent).FirstChild.Attributes("level").Value
                        Catch
                            curClip.Gain = "1"
                        End Try
                        Try
                            curClip.FadeIn = Eventlist(aEvent).FirstChild.Attributes("fadeIn.length").Value
                        Catch
                            curClip.FadeIn = "0"
                        End Try
                        Try
                            curClip.FadeOut = Eventlist(aEvent).FirstChild.Attributes("fadeOut.length").Value
                        Catch
                            curClip.FadeOut = "0"
                        End Try


                        ' get the filename by referenceing the mediapool xml file
                        Dim poolID = mediaPoolXML.SelectSingleNode("/MediaPool/Attributes/MediaFolder/AudioClip[@mediaID=" & Q & curClip.ClipID & Q & "]/Url").Attributes("url").FirstChild.Value

                        ' get the clip filename from the mediapool
                        curClip.FileName = Replace(Replace(poolID, "file:///", ""), "/", "\")

                        ' the Track.ChildClips property is List(Of Clip) to hold all of
                        ' the clips for a given track in that property's array



                        '  Stop
                        ' add the current clip to the Track property array
                        curTrack.ChildClips.Add(curClip)


                        ' update the total clips var because the Clips object will 
                        ' go out of scope  at the end and we'll  need this value
                        TotalClips = TotalClips + 1
                        ' Stop
                    Next
skipclip:
                End If
            Next

            ' add the current Track to the Song ChildTracks property array
            curSong.ChildTracks.Add(curTrack)
        Next

        ' copy the current Song over to the public LastSong Song
        ' object before all of the data goes out of scope
        ' we'll use that object's data for exporting
        loadedSong = curSong
        ' Stop
        '******************************************************************************************************************
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '/////////////////////////////////////////  UPDATE THE MAIN UI LABEL FIELDS  //////////////////////////////////////
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        '******************************************************************************************************************
        Form1.lblSongName.Text = curSong.Name
        Form1.lblSampleRate.Text = curSong.SampleRate
        Form1.lblTempo.Text = curSong.Tempo
        Form1.lblTimeSig.Text = curSong.TimeSig
        Form1.lblTrackCount.Text = curSong.ChildTracks.Count
        Form1.lblClipCount.Text = TotalClips
        Form1.lblMarkerCount.Text = curSong.SongMarkers.Count
        Try
            Form1.txtNotes.Text = My.Computer.FileSystem.ReadAllText(TempPath & "\notes.txt")
        Catch
        End Try
        Form1.lblImportType.Text = "In Format: " & DAW


    End Sub
