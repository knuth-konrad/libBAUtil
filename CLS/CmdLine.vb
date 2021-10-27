Imports libBAUtil.StringUtil

Public Class CmdLine

    Private msArgsDelimiter As String = "-/"
    Private msParamDelimiter As String = ":="

    Public Property ArgsDelimiter() As String
        Get
            Return msArgsDelimiter
        End Get
        Set(ByVal sValue As String)
            msArgsDelimiter = sValue
        End Set
    End Property

    Public Property ParamDelimiter() As String
        Get
            Return msParamDelimiter
        End Get
        Set(ByVal sValue As String)
            msParamDelimiter = sValue
        End Set
    End Property

    Public Function ArgsCount() As Integer
        ArgsCount = My.Application.CommandLineArgs.Count
    End Function

    Public Function ArgsGetByNameS(ByVal sParam As String) As String

        Try
            For i As Integer = 1 To ArgsDelimiter.Length
                For j As Integer = 1 To ParamDelimiter.Length
                    For Each s As String In My.Application.CommandLineArgs
                        If s.ToLower.StartsWith(Mid(ArgsDelimiter, i, 1) & sParam & Mid$(ParamDelimiter, j, 1)) Then
                            Return s.Remove(0, (Mid(ArgsDelimiter, i, 1) & sParam & Mid$(ParamDelimiter, j, 1)).Length)
                        End If
                    Next s
                Next j
            Next i
        Catch e As Exception
            Return ""
        End Try

        Return ""

    End Function

    Public Function ArgsGetByNameN(ByVal sParam As String) As Integer

        Try
            For i As Integer = 1 To ArgsDelimiter.Length
                For j As Integer = 1 To ParamDelimiter.Length
                    For Each s As String In My.Application.CommandLineArgs
                        If s.ToLower.StartsWith(Mid$(ArgsDelimiter, i, 1) & sParam & Mid$(ParamDelimiter, j, 1)) Then
                            Return CType(s.Remove(0, Len(Mid$(ArgsDelimiter, i, 1) & sParam & Mid$(ParamDelimiter, j, 1))), Integer)
                        End If
                    Next s
                Next j
            Next i
        Catch e As Exception
            Return 0
        End Try

    End Function

    Public Function ArgsExist(ByVal sParam As String) As Boolean

        Try
            For i As Integer = 1 To ArgsDelimiter.Length
                For Each s As String In My.Application.CommandLineArgs
                    If s.ToLower.StartsWith(Mid$(ArgsDelimiter, i, 1) & sParam) Then Return True
                Next s
            Next i
        Catch e As Exception
            Return False
        End Try

        Return False

    End Function

    Private Sub ParseCommandLineArgs()
        Dim inputArgument As String = "/input="
        Dim inputName As String = ""

        For Each s As String In My.Application.CommandLineArgs
            If s.ToLower.StartsWith(inputArgument) Then
                inputName = s.Remove(0, inputArgument.Length)
            End If
        Next

        If inputName = "" Then
            MsgBox("No input name")
        Else
            MsgBox("Input name: " & inputName)
        End If
    End Sub



End Class
