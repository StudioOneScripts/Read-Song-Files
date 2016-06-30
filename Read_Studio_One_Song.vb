     '*************************************************************************************
     ' This module will read a Presonus Studio One *.song file and extract relevant
     ' details including fader, pand and clip settings.  It may be useful if you want 
     ' to convert a Studio One song to another populr audio workstation's song format.

     ' This uses Ionic.zip dll to unzip the file  https://dotnetzip.codeplex.com/

       '/// **************  CODE TO ALWAYS COPY THE ZIP DLL RESOURCE IF NOT EXISTING  ************///
       ' Use this code in Form_Load and make sure you put the dll in your project resources.
      
       '' Get the path / folder that the application was launched from
       ' Dim strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)

       '' if the zip dll doesn't already exist there, copy it there from resources 
       ' Dim zipdll = Replace(strPath, "file:\", "") & "\Ionic.Zip.dll"
       ' strPath = Replace(strPath, "file:\", "")
       ' If My.Computer.FileSystem.FileExists(zipdll) = False Then
       '     File.WriteAllBytes(strPath & "\Ionic.Zip.dll", My.Resources.Ionic_Zip)
       ' End If
       '/// **************************************************************************************///
       
    '**************************************************************************************

Imports System.Xml
Imports Ionic.Zip

Module Read_Studio_One_Song

    ' temp path folder for the Studio One song file extraction
    Dim TempPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\SongTemp"
    
    ' variable for the class objects and total clips counter
    Dim curSong As Song, curTrack As Track, TotalClips As Integer

    Public Sub Read_Studio_One_Song_File()

        '******************************************************************************************************************
        '                           OPEN AND UNZIP A STUDIO ONE SONG FILE TO GET THE DATA  
        '******************************************************************************************************************

        ' open a file dialog to choose the file
        Dim f As New OpenFileDialog
        f.InitialDirectory = ""
        f.Filter = "Studio One Song (*.song)|*.song"
        f.Title = "Open a Studio One Song ..."
        f.ShowDialog()

        ' if user cancels, exit
        If My.Computer.FileSystem.FileExists(f.FileName) = False Then Exit Sub

        ' remove temp folder from any previous operation and create a new one
        Try
            My.Computer.FileSystem.DeleteDirectory(TempPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Catch
            ' silent catch
        End Try

        ' create a new temp directory
        My.Computer.FileSystem.CreateDirectory(TempPath)


        '******************************************************************************************************************
        '                            EXTRACT THE FILES BELOW FROM THE STUDIO ONE SONG FILE PACKAGE   
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
            MsgBox("Zip extraction error." & " " & vbNewLine & ex1.ToString, vbOKOnly + vbInformation, "Unzip Song Error")

            ' no point in proceeding if you can't open the song package
            Exit Sub
        End Try

        '******************************************************************************************************************
        '                             INSTANCE A NEW SONG CLASS AND PARSE THE METAINFO.XML FILE    
        '******************************************************************************************************************

        ' init a song class object
        Dim curSong As New Song
        
        ' set the song name in the class object
        curSong.Name = Replace(f.SafeFileName, ".song", "")

        ' load the mediapool.xml file  (needed for clip filenames)
        Dim mediaPool = My.Computer.FileSystem.ReadAllText(TempPath & "\Song\mediapool.xml")
        Dim mediaPoolXML As New XmlDocument
        mediaPoolXML.LoadXml(Replace(mediaPool, "x:", ""))

        ' load the audiomixer.xml file (needed for fader and pan
        Dim audioMixer = My.Computer.FileSystem.ReadAllText(TempPath & "\Devices\audiomixer.xml")
        Dim audioMixerXML As New XmlDocument
        audioMixerXML.LoadXml(Replace(audioMixer, "x:", ""))

        ' load the metainfo.xml file
        Dim metaInfo = My.Computer.FileSystem.ReadAllText(TempPath & "\metainfo.xml")
        Dim metaXML As New XmlDocument
        metaXML.LoadXml(metaInfo)

        'quote char for the strings we'll build below
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
        '                    GET SOME OTHER THINGS FROM THE SONG.XML FILE:  MARKERS, SONGS AND CLIPS   
        '******************************************************************************************************************

        ' load the song.xml into an XML document while also replacing "x:" which is formatting VS has issue with
        Dim songInfo = Replace(My.Computer.FileSystem.ReadAllText(TempPath & "\Song\song.xml"), "x:", "")
        
        ' create a new xml document and load the file 
        Dim songXML As New XmlDocument
        songXML.LoadXml(songInfo)

        ' get the song markers from song.xml file
        curSong.Markers = New List(Of String)
        Dim Markers = songXML.SelectSingleNode("/Song/Attributes/List/MarkerTrack").ChildNodes

        ' loop through the array and add the markers to the song's property array
        ' skip the first two markers which are always the default Start and End markers
        For I = 2 To Markers.Count - 1
            curSong.Markers.Add(Markers(I).Attributes("start").Value & "|" & Markers(I).Attributes("name").Value)
        Next

        '******************************************************************************************************************
        '                           READ IN ALL OF THE TRACKS AND CLIPS FROM THE SONG FILE   
        '******************************************************************************************************************

        ' this node list is a collection of all of the matching nodes below, audio tracks
        Dim TrackList As XmlNodeList = songXML.SelectNodes("/Song/Attributes/List/MediaTrack[@mediaType=" & Q & "Audio" & Q & "]")

        ' song.ChildTracks is a List(Of Track) property for holding all of the tracks
        curSong.ChildTracks = New List(Of Track)

        ' loop through the tracklist node list and get each track's properties
        For i = 0 To TrackList.Count - 1

            ' init a new Track class object on every cycle
            curTrack = New Track

            ' init the new object's list of clip array on every cycle
            curTrack.ChildClips = New List(Of Clip)

            ' set the track name with the FormatTrackName() method
            ' to check for empty track names and other things
            curTrack.Name = curSong.FormatTrackName(TrackList(i).Attributes("name").Value)
            curTrack.Color = TrackList(i).Attributes("color").Value

            ' *********************************************************
            ' get the volume and pan for the current track by cross
            ' referencing the GUID of the track with the audiomixer.xml
            '***********************************************************

            ' get the channel ID
            curTrack.TrackID = TrackList(i).FirstChild.NextSibling.Attributes("uid").Value

            ' get the reference track from the audiomixer.xml file by matching the uid there
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

            '**************************************************************************
            '           READ ALL OF THE CLIPS ON THE CURRENT TRACK
            '**************************************************************************

            ' loop through all of the child nodes beneath this
            ' track and get all of the clips on this track
            Dim ChildList As XmlNodeList = TrackList(i).ChildNodes
            For c = 0 To ChildList.Count - 1
                If ChildList(c).Name = "List" Then

                    Dim Eventlist As XmlNodeList = ChildList(c).ChildNodes

                    '*************************  CLIPS *****************************
                    ' <AudioEvent> tags.  The Try/Catch blocks are to ensure that
                    ' every attribute we're looking for exists.  If not, create it.
                    ' Some attributes only get created when the user changes them
                    ' and we may need them all to translate to other products.
                    '**************************************************************

                    For aEvent = 0 To Eventlist.Count - 1

                        ' create a new Clip class object
                        Dim curClip As New Clip

                        '******************************************************
                        ' if the currrent clip doesn't have both a name and
                        ' a clipID attribute, something is wrong, skip it
                        '******************************************************
                        
                        Try
                            curClip.Name = Eventlist(aEvent).Attributes("name").Value

                            curClip.ClipID = Eventlist(aEvent).Attributes("clipID").Value
                        Catch
                            GoTo skipclip
                        End Try

                        ' assign the current track UID to the clip for reference
                        curClip.ParentTrack = curTrack.TrackID
                        
                        ' beging setting some of the more typical clip properties
                        curClip.Length = Eventlist(aEvent).Attributes("length").Value
                        curClip.Speed = Eventlist(aEvent).Attributes("speed").Value
                        curClip.Pitch = Eventlist(aEvent).Attributes("transpose").Value
                        curClip.Tune = Eventlist(aEvent).Attributes("tune").Value

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

                        Try
                            curClip.Offset = Eventlist(aEvent).Attributes("offset").Value
                        Catch
                            curClip.Offset = "0"
                        End Try

                        Try
                            curClip.Gain = Eventlist(aEvent).FirstChild.Attributes("level").Value
                        Catch
                            curClip.Gain = "1"
                        End Try

                        Try
                            curClip.Pitch = Eventlist(aEvent).FirstChild.Attributes("transpose").Value
                        Catch
                            curClip.Pitch = "1"
                        End Try

                        Try
                            curClip.Tune = Eventlist(aEvent).FirstChild.Attributes("tune").Value
                        Catch
                            curClip.Tune = "0"
                        End Try

                        Try
                            curClip.Speed = Eventlist(aEvent).FirstChild.Attributes("speed").Value
                        Catch
                            curClip.Speed = "1"
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

                        '*************************************************************************
                        '        FIND AND FORMAT THE FILENAME FOR THE CURRENT AUDIO CLIP
                        '*************************************************************************
                        
                        ' get the filename by cross-referencing the mediapool.xml file UID
                        Dim poolID = mediaPoolXML.SelectSingleNode("/MediaPool/Attributes/MediaFolder/AudioClip[@mediaID=" & Q & curClip.ClipID & Q & "]/Url").Attributes("url").FirstChild.Value

                        ' get the clip filename from the mediapool.xml reference and format it
                        curClip.FileName = Replace(Replace(poolID, "file:///", ""), "/", "\")

                        '****************************************************************
                        ' the Track.ChildClips property is List(Of Clip) to hold all of
                        ' the clips for a given track in that property's array
                        '****************************************************************

                        ' add the current clip to the Track.ChildClips property array
                        curTrack.ChildClips.Add(curClip)

                        ' update the totalclips var because the curClip object will 
                        ' go out of scope on every loop cycle  and we'll need this value
                        TotalClips = TotalClips + 1
                    Next
skipclip:
                End If
            Next

            ' add the current Track to the Song ChildTracks property array
            curSong.ChildTracks.Add(curTrack)
        Next

        '*************************************************************
        ' copy the current Song object over to the public loadedSong 
        ' class object because curSong will go out of scope
        ' ************************************************************
        loadedSong = curSong


        '*****************************************************************
        ' OPTIONALY PRINT A NICELY FORMATTED OVERVIEW OF THE SONG DETAILS
        '*****************************************************************

        Exit Sub  ' uncomment to always get test printout of the entire song
        loadedSong.PrintSongData()

    End Sub

End Module

