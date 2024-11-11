'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.IO
Imports System.Linq
Imports System.Text.Encoding

'This module contains this program's core procedures.
Public Module CoreModule
   Private Const BITS_PER_SAMPLE As Short = 8                   'Defines the bits per sample in the output WAVE file.
   Private Const CHANNEL_COUNT As Short = 1                     'Defines the number of channels in the output WAVE file.
   Private Const CHUNK_SIZE As Integer = 16                     'Defines the chunk size in the output WAVE file.
   Private Const CVF_HEADER As String = "Creative Voice File"   'Defines the CVF header.
   Private Const PCM_AUDIO_FORMAT As Short = 1                  'Defines the PCM audio format indicator in the output WAVE file.
   Private Const SAMPLE_RATE As Integer = 16000                 'Defines the sample rate in the output WAVE file.
   Private Const WAVE_HEADER_SIZE As Integer = 36               'Defines the output WAVE file header's size.

   'This procedure is executed when this program is started.
   Public Sub Main()
      Try
         Dim CVFData As List(Of Byte) = Nothing
         Dim InFile As String = If(GetCommandLineArgs().Count > 1, GetCommandLineArgs().Last().Trim(), Nothing)
         Dim OutFile As String = $"{InFile}.wav"

         If InFile = Nothing Then
            With My.Application.Info
               Console.WriteLine($"{ .Title} v{ .Version.ToString()} - by: { .CompanyName}, { .Copyright}")
               Console.WriteLine($"{NewLine}{ .Description}{NewLine}")
               Console.WriteLine($"Specify the input file as a command line argument.")
            End With
         Else
            CVFData = New List(Of Byte)(File.ReadAllBytes(InFile))
            If ASCII.GetString(CVFData.ToArray()).Substring(0, CVF_HEADER.Length) = CVF_HEADER Then
               CVFToWAVE(CVFData.GetRange(CVF_HEADER.Length, CVFData.Count - CVF_HEADER.Length).ToArray(), OutFile)
               Console.WriteLine($"""{OutFile}"" created.")
            Else
               Console.WriteLine("Not a ""Create Voice File"" file.")
            End If
         End If
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try
   End Sub

   'This procedure writes the specified CVF data as a WAVE file at the specified path.
   Public Sub CVFToWAVE(RawCVFData() As Byte, OutPath As String)
      Try
         Using FileO As New BinaryWriter(New FileStream(OutPath, FileMode.Create, FileAccess.Write))
            FileO.Write(ASCII.GetBytes("RIFF"))
            FileO.Write(WAVE_HEADER_SIZE + RawCVFData.Count)
            FileO.Write(ASCII.GetBytes("WAVE"))
            FileO.Write(ASCII.GetBytes("fmt "))
            FileO.Write(CHUNK_SIZE)
            FileO.Write(PCM_AUDIO_FORMAT)
            FileO.Write(CHANNEL_COUNT)
            FileO.Write(SAMPLE_RATE)
            FileO.Write(SAMPLE_RATE * CHANNEL_COUNT * BITS_PER_SAMPLE \ 8)
            FileO.Write(CShort(CHANNEL_COUNT * BITS_PER_SAMPLE \ 8))
            FileO.Write(BITS_PER_SAMPLE)
            FileO.Write(ASCII.GetBytes("data"))
            FileO.Write(RawCVFData.Count)
            FileO.Write(RawCVFData)
         End Using
      Catch ExceptionO As Exception
         Console.WriteLine(ExceptionO.Message)
      End Try
   End Sub

End Module
