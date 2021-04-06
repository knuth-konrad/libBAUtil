'------------------------------------------------------------------------------
'   Author: Knuth Konrad 2021-04-05
'  Changed: -
'------------------------------------------------------------------------------
Imports libBAUtil.StringUtil

''' <summary>
''' Commandline parameter handling
''' See https://docs.microsoft.com/en-us/archive/msdn-magazine/2019/march/net-parse-the-command-line-with-system-commandline
''' See https://github.com/commandlineparser/commandline
''' </summary>
Public Class CmdArgs

#Region "Declarations"

   ' Default arguments and key/value delimiter
   Private Const DELIMITER_ARGS As String = "/"
   Private Const DELIMITER_VALUE As String = "="

   Dim msDelimiterArgs As String = String.Empty    ' Arguments delimiter, typically "/"
   Dim msDelimiterValue As String = String.Empty   ' Key/value delimiter, typically "="
   Dim mlArgsCount As Int32                        ' # of arguments
   Dim masParams() As String                       ' Name of all arguments, e.g. /file = "file"
   Dim maoValues() As Object                       ' All values

   Private mcolKeyValues As List(Of KeyValue)

#End Region

#Region "Properties - Public"

   Public Property DelimiterArgs As String
      Get
         Return msDelimiterArgs
      End Get
      Set(value As String)
         msDelimiterArgs = value
      End Set
   End Property

   Public Property DelimiterValue As String
      Get
         Return msDelimiterValue
      End Get
      Set(value As String)
         msDelimiterValue = value
      End Set
   End Property

   Public ReadOnly Property ArgsCount As Int32
      Get
         Return mlArgsCount
      End Get
   End Property

   Public Property KeyValues As List(Of KeyValue)
      Get
         Return mcolKeyValues
      End Get
      Set(value As List(Of KeyValue))
         mcolKeyValues = value
      End Set
   End Property

#End Region

#Region "Methods - Private"

   Private Overloads Function ParseCmd(ByVal asArgs() As String, Optional ByVal startIndex As Int32 = 0) As Boolean
      ' Stub

      Dim sKey As String = String.Empty, sValue As String = String.Empty


      For i As Int32 = startIndex To asArgs.Length - 1
         Console.WriteLine(asArgs(i))
         Return ParseParam(asArgs(i))
      Next

      Return True
   End Function

   Private Overloads Function ParseCmd(ByVal sArgs As String) As Boolean
      ' Stub

      Return True
   End Function

   ''' <summary>
   ''' Parses a single key/pair combo into a matching <see cref="KeyValue"/> object
   ''' </summary>
   ''' <param name="sParam"></param>
   ''' <returns></returns>
   Private Function ParseParam(ByVal sParam As String) As Boolean

      If sParam.Length < 1 Then
         Return True
      End If

      With Me

         If sParam.Contains(.DelimiterValue) Then
            ' Parameter of the form /key=value

            Dim o As New KeyValue

            With o
               .KeyLong = Left(sParam, InStr(sParam, DelimiterValue) - 1).Trim
               ' Since we parse this from the command line, set both
               .KeyShort = .KeyLong
               .Value = Mid(sParam, InStr(sParam, DelimiterValue) + 1)
            End With

            .KeyValues.Add(o)

         Else
            ' Parameter of the form /Value.
            ' These are considered to be boolean parameters. If present, their value is 'True'

            Dim o As New KeyValue

            With o
               .KeyLong = sParam.Trim
               ' Since we parse this from the command line, set both
               .KeyShort = .KeyLong
               .Value = True
            End With

            .KeyValues.Add(o)

         End If

      End With

      Return True

   End Function


#End Region

#Region "Methods - Public"

   ''' <summary>
   ''' Initializes the object by parsing System.Environment.GetCommandLineArgs()
   ''' </summary>
   ''' <returns></returns>
   Public Function Initialize(Optional ByVal cmdLineArgs As String = "") As Boolean

      If cmdLineArgs.Length < 1 Then
         Dim asArgs() As String = System.Environment.GetCommandLineArgs()
         ' When using System.Environment.GetCommandLineArgs(), the 1st array element is the executables name
         Return ParseCmd(asArgs, 1)
      Else
         Return ParseCmd(cmdLineArgs)
      End If

   End Function


#End Region

#Region "Constructor/Dispose"

   Public Sub New()

      MyBase.New

      With Me
         .DelimiterArgs = DELIMITER_ARGS
         .DelimiterValue = DELIMITER_VALUE

         .KeyValues = New List(Of KeyValue)
      End With

   End Sub

   Public Sub New(Optional ByVal delimiterArgs As String = DELIMITER_ARGS, Optional ByVal delimiterValue As String = DELIMITER_VALUE)

      MyBase.New

      ' Safe guard
      If delimiterArgs.Length < 1 OrElse delimiterValue.Length < 1 Then
         Throw New ArgumentOutOfRangeException("Empty argument or key/value delimiter are not allowed.")
      End If

      With Me
         .DelimiterArgs = delimiterArgs
         .DelimiterValue = delimiterValue

         .KeyValues = New List(Of KeyValue)
      End With

   End Sub

#End Region

End Class

Public Class KeyValue

   Private msHelpText As String = String.Empty
   Private msKeyLong As String = String.Empty
   Private msKeyShort As String = String.Empty
   Private mbolMandatory As Boolean = False
   Private moValue As Object

   ''' <summary>
   ''' Help text for this parameter
   ''' </summary>
   Public Property HelpText As String
      Get
         Return msHelpText
      End Get
      Set(value As String)
         msHelpText = value
      End Set
   End Property

   ''' <summary>
   ''' Indicates that this parameter is mandatory
   ''' </summary>
   ''' <returns></returns>
   Public Property IsMandatory As Boolean
      Get
         Return mbolMandatory
      End Get
      Set(value As Boolean)
         mbolMandatory = value
      End Set
   End Property

   ''' <summary>
   ''' 'Outspoken' parameter name, e.g. /file
   ''' </summary>
   Public Property KeyLong As String
      Get
         Return msKeyLong
      End Get
      Set(value As String)
         msKeyLong = value
      End Set
   End Property

   ''' <summary>
   ''' Short parameter name, e.g. /f
   ''' </summary>
   Public Property KeyShort As String
      Get
         Return msKeyShort
      End Get
      Set(value As String)
         msKeyShort = value
      End Set
   End Property

   ''' <summary>
   ''' Parameter name, e.g. 'file' or 'f' from /file=MyFile.txt  or /f=MyFile.txt
   ''' </summary>
   Public ReadOnly Property Key As String
      Get
         ' KeyLong takes precedences
         With Me
            If .KeyLong.Length > 0 Then
               Return .KeyLong
            Else
               Return .KeyShort
            End If
         End With
      End Get
   End Property

   ''' <summary>
   ''' Parameter value, e.g. 'MyFile.txt' from /file=MyFile.txt
   ''' </summary>
   Public Property Value As Object
      Get
         Return moValue
      End Get
      Set(value As Object)
         moValue = value
      End Set
   End Property

#Region "Methods - Public"

   Public Overrides Function ToString() As String

      With Me
         Dim sText As String = .Key
         If .IsMandatory = True Then
            sText &= " (mandatory)"
         End If
         If .HelpText.Length < 1 Then
            Return sText
         Else
            Return sText & ": " & .HelpText
         End If
      End With

   End Function

#End Region

   Public Sub New(Optional ByVal keyShort As String = "", Optional ByVal keyLong As String = "",
                  Optional ByVal value As Object = Nothing, Optional ByVal isMandatory As Boolean = False,
                  Optional ByVal helpText As String = "")

      MyBase.New

      With Me
         .HelpText = helpText
         .IsMandatory = isMandatory
         .KeyLong = keyLong
         .KeyShort = keyShort
         .Value = value
      End With

   End Sub

End Class
