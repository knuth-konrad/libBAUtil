Imports libBAUtil.StringUtil

Public Class CmdLineParser
   '------------------------------------------------------------------------------
   'File:    : baCmdLine.vb
   'Purpose  : Command line parameter parsing
   '
   'Prereq.  : -
   'Note     : -
   '
   '   Author: Knuth Konrad
   '     Date: 07.06.2018
   '   Source: -
   '  Changed: -
   '------------------------------------------------------------------------------
#Region "Declares"

   Private msArgsDelimiter As String = "-/"     ' Default argument delimiters, e.g. -help or /help
   Private msParamDelimiter As String = ":="    ' Default parameter delimiters, e.g. -arg=val or /arg:val

   ' The actual command line as passed to the class or retrieved via Environment.CommandLine
   Private msCommandLine As String = String.Empty

#End Region

#Region "Properties - Private"
#End Region

#Region "Properties - Public"

   Public Property ArgsDelimiter() As String
      Get
         Return msArgsDelimiter
      End Get
      Set(ByVal sValue As String)
         msArgsDelimiter = sValue
      End Set
   End Property

   Public Property CommandLine As String
      Get
         Return msCommandLine
      End Get
      Set(value As String)
         msCommandLine = value
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

#End Region

#Region "Implementations"
#End Region

#Region "Methods - Private"

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

#End Region

#Region "Methods - Public"

   Public Function ArgsCount() As Integer
      ArgsCount = My.Application.CommandLineArgs.Count
   End Function

   Public Function ArgsGetByNameS(ByVal sParam As String) As String

      Try
         For i As Integer = 1 To ArgsDelimiter.Length
            For j As Integer = 1 To ParamDelimiter.Length
               For Each s As String In My.Application.CommandLineArgs
                  If s.ToLower.StartsWith(Mid$(ArgsDelimiter, i, 1) & sParam & Mid$(ParamDelimiter, j, 1)) Then
                     Return s.Remove(0, Len(Mid$(ArgsDelimiter, i, 1) & sParam & Mid$(ParamDelimiter, j, 1)))
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

   Public Function Init(ByVal sCmd As String) As Boolean
      '------------------------------------------------------------------------------
      'Purpose  : !!! This method must be called first !!!
      '           Initialze the object with the passed command line parameters
      '
      'Prereq.  : -
      'Parameter: sCmd  -  Command line passed to the calling application
      'Returns  : True  -  Success of parsing
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim i As Long, lCount As Long
      Dim asParams() As String
      Dim avntValues As Object

      On Error GoTo InitError

      asParams() = Split(sCmd, ParamDelimiter)

      ' Enthielt CLI Werte?
      If UBound(asParams) > -1 Then
         ReDim avntValues(0 To UBound(asParams)) As Variant
   Else
         Init = True
         Exit Function
      End If

      For i = 0 To UBound(asParams)
         'Nur wenn auch ein Parameter da ist...
         If Len(Trim$(asParams(i))) > 0 Then
            'Parameter in der Art /User=Knuth.
            '"User" ist der Parameter, "Knuth" ist der Wert
            If InStr(asParams(i), ValueDelimiter) > 0 Then
               avntValues(i) = Mid(asParams(i), InStr(asParams(i), ValueDelimiter) + 1)
               asParams(i) = Trim$(Left$(asParams(i), InStr(asParams(i), ValueDelimiter) - 1))
            Else
               'Parameter in der Art /Quit.
               '"Quit" ist der Parameter, als Wert wird "True" angenommen
               avntValues(i) = True
               asParams(i) = Trim$(asParams(i))
            End If
            ParamCount = ParamCount + 1
         End If
      Next i

      'Arrays "umschaufeln" damit wir ein 1-basiertes Array bekommen
      ReDim masParams(1 To ParamCount)
      ReDim mavntValues(1 To ParamCount)

      lCount = 1
      For i = 0 To UBound(asParams)
         If Len(Trim$(asParams(i))) > 0 Then
            masParams(lCount) = asParams(i)
            mavntValues(lCount) = avntValues(i)
            lCount = lCount + 1
         End If
      Next i

      Init = True

InitExit:
      On Error GoTo 0
      Exit Function

InitError:
      Init = False
      'Arrays "aufräumen"
      Erase masParams, mavntValues
      Resume InitExit

   End Function

#End Region

#Region "Constructor/Deconstructor"

   Public Sub New()

      MyBase.New
      Init System.Environment.CommandLine

   End Sub

#End Region

End Class
