Imports SpeechLib

Public Class Form1

    Private WithEvents Voice As SpVoice = Nothing

    Private Sub SayText(ByVal TextToSay As String)
        If Voice Is Nothing Then Voice = New SpVoice
        Voice.Speak(TextToSay, SpeechVoiceSpeakFlags.SVSFlagsAsync)
    End Sub

    Private Sub Voice_EndStream(StreamNumber As Integer, StreamPosition As Object) Handles Voice.EndStream
        Voice = Nothing
    End Sub

    Private Function ISpVoice() As Object
        Throw New NotImplementedException
    End Function

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            GC.Collect()
            SayText(Clipboard.GetText)
        Catch ex As Exception
            'nix machen wenn fehler
        End Try
    End Sub

End Class
