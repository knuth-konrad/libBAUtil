''' <summary>
''' Command line arguments parser.
''' </summary>
Public Class CmdLineArgs

   ' Defaults
   Private Const PARAM_DELIMITER As String = "/"
   Private Const VALUE_DELIMITER As String = "="

   Private Structure cmdLineCollection
      Public Index As Int32
      Public Parameter As String
      Public Value As Object
   End Structure

   Private msParamDelimiter As String = PARAM_DELIMITER
   Private mlParamCount As Int32

   Private msValueDelimiter As String = VALUE_DELIMITER

   Private masParams As List(Of cmdLineCollection)        'Alle Parameternamen


   Public Property ParamDelimiter As String
      Get
         Return msParamDelimiter
      End Get
      Set(value As String)
         msParamDelimiter = value
      End Set
   End Property

   Public ReadOnly Property ParamCount() As Int32
      Get
         Return mlParamCount
      End Get
   End Property

   Public Property ValueDelimiter As String
      Get
         Return msParamDelimiter
      End Get
      Set(value As String)
         msParamDelimiter = value
      End Set
   End Property

   ''' <summary>
   ''' Returns the parameter's name, e.g. for /MyParam=MyValue, "MyParam" is returned.
   ''' </summary>
   ''' <param name="paramIndex">
   ''' 0-based index for which the parameter name should be returned.
   ''' </param>
   ''' <returns>
   ''' Name of the parameter
   ''' </returns>
   Public Function GetParamNameByIndex(Optional ByVal paramIndex As Int32 = 0) As String

      ' Safe guard
      If (paramIndex < masParams.Count - 1) Or (paramIndex > masParams.Count - 1) Then
         Throw New ArgumentOutOfRangeException("paramIndex")
      End If

      For Each o As cmdLineCollection In masParams
         If o.Index = paramIndex Then
            Return o.Parameter
         End If
      Next

      Return String.Empty

   End Function

   Public Function Init(ByVal cmdLine As String) As Boolean
      '------------------------------------------------------------------------------
      'Purpose  : Diese Prozedur muß els erste Prozedur aufgerufen werden. Sie übernimmt
      '           die Kommandozeile und wertet sie aus.
      '
      'Prereq.  : -
      'Parameter: sCmd  -  Per COMMAND$ empfangene Kommandozeile
      'Returns  : True  -  Kommandozeile konnte ausgewertet werden
      '           False -  Ein Fehler ist aufgetreten
      'Note     : -
      '
      '   Author: Knuth Konrad 27.09.2000
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim i As Long, lCount As Long
      Dim asParams() As String
      Dim avntValues As Variant

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


   Public Sub New(Optional ByVal prmDelimiter As String = PARAM_DELIMITER,
                  Optional ByVal valDelimiter As String = VALUE_DELIMITER)

      MyBase.New

      With Me
         .ParamDelimiter = prmDelimiter
         .ValueDelimiter = valDelimiter
      End With

   End Sub

End Class
