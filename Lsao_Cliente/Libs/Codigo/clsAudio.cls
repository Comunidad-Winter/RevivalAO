VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsAudio.cls - Controls all audio using DirectSound and DirectMusic
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@hotmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Private Const NumSoundBuffers As Byte = 30

'WAVE Sounds
Private Type SoundBuffer
    FileName As String
    looping As Boolean
    Buffer As DirectSoundBuffer
End Type

Public Enum LoopStyle
    Default = 0     'If wave is in memory use it's looping setting, otherwise false
    Disabled = 1    'Never loop
    Enabled = 2     'Loop
End Enum

Dim DirectSound As DirectSound
Dim DSBuffers(1 To NumSoundBuffers) As SoundBuffer

Dim SoundPath As String

Dim SndVolume As Long   'Sound volume

'Music
Dim Performance As DirectMusicPerformance
Dim Segment As DirectMusicSegment
Dim Loader As DirectMusicLoader
Dim SegState As DirectMusicSegmentState

Dim Music As Boolean        'True if music is supported by HW

Dim MusicPath As String

Private Sub Class_Initialize()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'
'***************************************************
End Sub

Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Destroy's all DirectMusic and DirectSound objects
'***************************************************
On Local Error Resume Next
    Dim loopc As Long
    
    'Stop every channel being used and destroy the buffer
    For loopc = 1 To NumSoundBuffers
        Call DSBuffers(loopc).Buffer.Stop
        Set DSBuffers(loopc).Buffer = Nothing
    Next loopc
    
    'Stop music
    Call Performance.Stop(Segment, SegState, 0, 0)
    
    Call Performance.CloseDown
    
    Set DirectSound = Nothing
    
    Set Segment = Nothing
    Set Performance = Nothing
    Set Loader = Nothing
End Sub

Public Function Initialize(ByRef DX As DirectX7, ByVal hwnd As Long, ByVal sound_path As String, ByVal music_path As String) As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Creates all DS and DM objects
'***************************************************
On Error GoTo ErrHandler

    'Validate DX object
    If DX Is Nothing Then GoTo ErrHandler
    
    'Initialize DirectSound
    Set DirectSound = DX.DirectSoundCreate("")
    
    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
    
    'Set sound path
    SoundPath = sound_path
    
    'Set default volume
    SndVolume = 0
    
    'Initialize DirectMusic
    Music = InitializedirectMusic(DX, hwnd)
    
    'Set music path
    MusicPath = music_path
    
    Initialize = True
Exit Function

ErrHandler:
End Function

Private Function InitializedirectMusic(ByRef DX As DirectX7, ByVal hwnd As Long) As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Initialize DM. Returns false if music isn't supported by HW
'***************************************************
On Error GoTo ErrHandler
    Set Loader = DirectX.DirectMusicLoaderCreate()
    
    Set Performance = DX.DirectMusicPerformanceCreate()
    Performance.Init DirectSound, hwnd
    Performance.SetMasterAutoDownload True     'Enable auto doanload of instruments
    
    'Set to port default, 1 channel group is enough for MIDIs
    Performance.SetPort -1, 1
    
    InitializedirectMusic = True
Exit Function

ErrHandler:
End Function

Public Function PlayWave(ByVal FileName As String, Optional ByVal LoopSound As LoopStyle = LoopStyle.Default) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/04/06
'Loads a new wave to memory (if it isn't already loaded) and starts playing it
'Returns the index in which the sound is playing
'***************************************************
On Error GoTo ErrHandler
    Dim bufferIndex As Long

    Dim data() As Byte
    Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", UCase$(FileName), data, WAV_RESOURCE_FILE)
    Open TemporalFile & ".WAV" For Binary Access Write As #1
        Put #1, , data
    Close #1

    'Get the buffer index were wave was loaded
    bufferIndex = LoadWave(FileName, LoopSound)
    If bufferIndex = 0 Then Exit Function   'If an error ocurred abort
    
    With DSBuffers(bufferIndex)
        If .looping Then
            .Buffer.Play DSBPLAY_LOOPING
        Else
            .Buffer.Play DSBPLAY_DEFAULT
        End If
        
        'Apply volume
        Call .Buffer.SetVolume(SndVolume)
    End With
    
    PlayWave = bufferIndex
    
    Kill TemporalFile & ".WAV"
Exit Function

ErrHandler:
End Function

Private Function LoadWave(ByVal FileName As String, ByVal looping As LoopStyle) As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 10/04/06
'Loads a new wave into memory (if it's not already loaded) and returns
'the buffer index
'***************************************************
On Error GoTo ErrHandler
    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Dim I As Long
    
    FileName = UCase$(FileName)
    
    'Check if the buffer is in memory and not playing
    For I = 1 To NumSoundBuffers
        If DSBuffers(I).FileName = FileName Then
            If DSBuffers(I).Buffer.GetStatus = 0 Or DSBuffers(I).Buffer.GetStatus = DSBSTATUS_BUFFERLOST Then
                'Found it!!! We just play this one :)
                LoadWave = I
                
                'Set looping if needed
                If looping <> LoopStyle.Default Then DSBuffers(I).looping = (looping = LoopStyle.Enabled)
                
                Exit Function
            End If
        End If
    Next I
    
    'Not in memory, search for an empty buffer
    For I = 1 To NumSoundBuffers
        If DSBuffers(I).Buffer Is Nothing Then
            Exit For
        End If
    Next I
    
    'No empty buffers?? Search for one not being played and use that slot
    If I > NumSoundBuffers Then
        For I = 1 To NumSoundBuffers
            If DSBuffers(I).Buffer.GetStatus = DSBSTATUS_BUFFERLOST Or DSBuffers(I).Buffer.GetStatus = 0 Then
                Exit For
            End If
        Next I
    End If
    
    'Still no free buffers?? Get the first non-looping sound
    If I > NumSoundBuffers Then
        For I = 1 To NumSoundBuffers
            If Not (DSBuffers(I).Buffer.GetStatus And DSBSTATUS_LOOPING) Then
                Exit For
            End If
        Next I
    End If
    
    'If every sound is looping, just ignore new sound unless it's looping too (in which case we set it at the first slot by default)
    If I > NumSoundBuffers Then
        If Not looping Then Exit Function
        I = 1
    End If
    
    'Create new sound
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2
    waveFormat.lSamplesPerSec = 22050
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    
    With DSBuffers(I)
        Set .Buffer = Nothing   'Get rid of any previous data
        
        .FileName = FileName
        .looping = (looping = LoopStyle.Enabled)
        
        Set .Buffer = DirectSound.CreateSoundBufferFromFile(TemporalFile & ".WAV", bufferDesc, waveFormat)
    End With
    
    LoadWave = I
Exit Function

ErrHandler:
End Function

Public Sub StopWave(Optional ByVal bufferIndex As Long = 0)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Stops the sound in the given buffer index, or every sound if
'index is invalid
'***************************************************
    If bufferIndex > 0 And bufferIndex <= NumSoundBuffers Then
        If DSBuffers(bufferIndex).Buffer.GetStatus And DSBSTATUS_PLAYING Then
            Call DSBuffers(bufferIndex).Buffer.Stop
        End If
    Else
        Dim I As Long
        For I = 1 To NumSoundBuffers
            If Not DSBuffers(I).Buffer Is Nothing Then
                If DSBuffers(I).Buffer.GetStatus <> 0 And DSBuffers(I).Buffer.GetStatus <> DSBSTATUS_BUFFERLOST Then
                    Call DSBuffers(I).Buffer.Stop
                End If
            End If
        Next I
    End If
End Sub

Public Property Get PlayingSound() As Boolean
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns true if there is at least one sound playing
'***************************************************
    Dim I As Long
    
    For I = 1 To NumSoundBuffers
        If DSBuffers(I).Buffer.GetStatus And DSBSTATUS_PLAYING Then
            PlayingSound = True
            Exit Property
        End If
    Next I
End Property

Public Property Let SoundVolume(ByVal Volume As Long)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the sound's volume
'***************************************************
    Dim I As Long
    
    For I = 1 To NumSoundBuffers
        If DSBuffers(I).Buffer.GetStatus And DSBSTATUS_PLAYING Then
            Call DSBuffers(I).Buffer.SetVolume(Volume)
        End If
    Next I
    
    SndVolume = Volume
End Property

Public Property Get SoundVolume() As Long
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume
'***************************************************
    SoundVolume = SndVolume
End Property

Public Function PlayMIDI(Optional ByVal file As String = "", Optional ByVal Loops As Long = -1) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi and start to play it
'***************************************************
On Error GoTo ErrHandler
    If Not Music Then Exit Function
    
    If PlayingMusic Then StopMidi
    
    If UCase$(file) = "0.MID" Then Exit Function
    
    If file <> "" Then
        Dim data() As Byte
        Call modCompression.Get_File_Data(App.Path & "\..\Recursos\", UCase$(file), data, MIDI_RESOURCE_FILE)
        Open TemporalFile For Binary Access Write As #1
            Put #1, , data
        Close #1
        
        If Not LoadMidi(TemporalFile) Then Exit Function
        Kill TemporalFile
    Else
        'Make sure we have a loaded segment
        If Segment Is Nothing Then Exit Function
    End If
    
    'Play it
    Segment.SetRepeats Loops
    
    Set SegState = Performance.PlaySegment(Segment, 0, 0)
    
    'Set tempo to 0 and volume to max
    MusicTempo = 0
    MusicVolume = 200
    
    PlayMIDI = True
Exit Function

ErrHandler:
End Function

Private Function LoadMidi(ByVal file As String) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Loads a midi to memory
'***************************************************
On Error GoTo ErrHandler
    If Not FileExist(file, vbArchive) Then Exit Function
    
    Call StopMidi
    
    'Destroy old object
    Set Segment = Nothing
    
    Set Segment = Loader.LoadSegment(file)
    
    If Segment Is Nothing Then
        Exit Function
    End If
    
    Segment.SetStandardMidiFile
    
    LoadMidi = True
Exit Function

ErrHandler:
End Function

Public Sub StopMidi()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Stops playing a midi
'***************************************************
On Error GoTo ErrHandler
    If PlayingMusic Then
        Call Performance.Stop(Segment, SegState, 0, 0)
    End If
Exit Sub

ErrHandler:
End Sub

Public Property Let MusicTempo(ByVal Tempo As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    Performance.SetMasterTempo Tempo
End Property

Public Property Let MusicVolume(ByVal Volume As Long)
'***************************************************
'Autor: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Sets the music's volume
'***************************************************
    If Not Music Then Exit Property
    
    Performance.SetMasterVolume Volume
End Property

Public Property Get MusicTempo() As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's tempo
'***************************************************
    If Not Music Then Exit Property
    
    MusicTempo = Performance.GetMasterTempo
End Property

Public Property Get MusicVolume() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the music's volume
'***************************************************
    If Not Music Then Exit Property
    
    MusicVolume = Performance.GetMasterVolume
End Property

Public Property Get PlayingMusic() As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns True if a midi is being played
'***************************************************
    If Not Music Or Segment Is Nothing Then Exit Property
    
    PlayingMusic = Performance.IsPlaying(Segment, SegState)
End Property

Public Property Get MusicLenght() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 16/03/04
'Returns the lenght of the current midi
'***************************************************
    If Not Music Then Exit Property
    
    MusicLenght = Segment.GetLength
End Property
