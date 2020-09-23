Attribute VB_Name = "WaveInfo"
'Waveform Processor
'Coded by Lance Fitz-Herbert (phrizer)
'ICQ: 23549284

Type RiffHdr_
    Id          As String * 4       'identifier string = "RIFF"
    Len         As Long             'remaining length *after* this header
    DataType    As String * 4       'type of data. wav = WAVE
End Type


Type ChunkHdr_                      'CHUNK 8-byte header
    Id          As String * 4       'identifier, e.g. "fmt " or "data"
    Len         As Long             'remaining chunk length *after* header
End Type                            'data bytes follow chunk header



Type WAVEFORMATEX                   'FMT Chunk
    wFormatTag      As Integer      'Format category
    nChannels       As Integer      'Number of channels
    nSamplesPerSec  As Long         'Sampling rate
    nAvgBytesPerSec As Long         'For buffer estimation
    nBlockAlign     As Integer      'Data block size
    wBitsPerSample  As Integer
    cbSize          As Integer
End Type

Type FactChunk_                     'Not always present
    dwFileSize As Long              'Number Of Samples
End Type

Type WaveData8bit_
    ChannelData() As Byte           '8bit samples, (channel)(samples)
End Type

Type WaveData16bit_
    ChannelData() As Integer        '16bit samples, (channel)(samples)
End Type


Dim RiffHdr As RiffHdr_
Dim FMTChunk As WAVEFORMATEX
Dim FACTChunk As FactChunk_
Dim ChunkHdr As ChunkHdr_
Dim WaveData8bit As WaveData8bit_
Dim WaveData16bit As WaveData16bit_
'
'-
'Number of Samples Per Channel is Calculated By:
'(The Length Of The DataChunk devide by the number of channels)
'devide by the number of bytes per sample
'-

Public Function GetWavInfo(FileName As String, Picbox As PictureBox)
    Dim FreeNum As Integer
    Dim TmpSeek As Long

    DoEvents
        
    FreeNum = FreeFile
    Open FileName For Binary Access Read As FreeNum
    
    'Get RIFF header
    Get #FreeNum, 1, RiffHdr
    
    If RiffHdr.Id <> "RIFF" Or RiffHdr.DataType <> "WAVE" Then
        MsgBox "Not A WaveForm File"
        Exit Function
    End If
    
    'Read Each Chunk in File
    Do
        'Save current position in file
        TmpSeek = Seek(FreeNum)
        'Read Next Chunk Header
        Get #FreeNum, , ChunkHdr
        
        'Proccess Chunks
        If ChunkHdr.Id = "fmt " Then
            Get #FreeNum, , FMTChunk
        End If
        
        If ChunkHdr.Id = "fact" Then
            Get #FreeNum, , FACTChunk
        End If
        
        If ChunkHdr.Id = "data" Then
            
            If FMTChunk.wBitsPerSample = 8 Then
                ReDim WaveData8bit.ChannelData(1 To FMTChunk.nChannels, 1 To ChunkHdr.Len / FMTChunk.nChannels)
                Get #FreeNum, , WaveData8bit.ChannelData
                DrawWave8 WaveData8bit, 8, Picbox
            End If
            
            If FMTChunk.wBitsPerSample = 16 Then
                ReDim WaveData16bit.ChannelData(1 To FMTChunk.nChannels, 1 To (ChunkHdr.Len / FMTChunk.nChannels) / 2)
                Get #FreeNum, , WaveData16bit.ChannelData
                DrawWave16 WaveData16bit, 16, Picbox
            End If
            
            
        End If
        
        'Test If Last Chunk Has been found
        If ChunkHdr.Len + Seek(FreeNum) >= LOF(FreeNum) Then
            Exit Do
        End If
        
        
        'Move Position in file to next chunk
        Seek #FreeNum, TmpSeek + ChunkHdr.Len + Len(ChunkHdr)
                    
    Loop
       
        
    Close FreeNum
    
End Function

Public Function DrawWave8(Samples As WaveData8bit_, Bits As Byte, Picbox As PictureBox)
    Dim nChannels As Integer    'Number of Channels in WaveForm
    Dim nSamples As Long        'Number of Samples per channel
    Dim Loopsamples As Long
    
    Channels = UBound(Samples.ChannelData, 1)
    nSamples = UBound(Samples.ChannelData, 2)
    
    Picbox.ScaleMode = 0
    Picbox.ScaleHeight = 2 ^ Bits
    Picbox.ScaleWidth = nSamples

    Picbox.CurrentY = (2 ^ Bits) / 2
    
    Picbox.Visible = False
    For Loopsamples = 1 To nSamples
        Picbox.Line -(Loopsamples, Samples.ChannelData(1, Loopsamples))
    Next
    Picbox.Visible = True
End Function
Public Function DrawWave16(Samples As WaveData16bit_, Bits As Byte, Picbox As PictureBox)
    Dim nChannels As Integer    'Number of Channels in WaveForm
    Dim nSamples As Long        'Number of Samples per channel
    Dim Loopsamples As Long
    
    Channels = UBound(Samples.ChannelData, 1)
    nSamples = UBound(Samples.ChannelData, 2)
    Picbox.ScaleMode = 0
    Picbox.ScaleHeight = (2 ^ Bits)
    Picbox.ScaleWidth = nSamples

    Picbox.CurrentY = (2 ^ Bits) / 2
     
    Picbox.Visible = False
    For Loopsamples = 1 To nSamples
        Picbox.Line -(Loopsamples, Samples.ChannelData(1, Loopsamples) + (2 ^ Bits) / 2)
    Next
    Picbox.Visible = True
End Function
