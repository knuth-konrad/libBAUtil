Imports System.ComponentModel

'Imports System.IO
'Imports System.IO.File
'Imports System.IO.Path

'Imports System.Xml
'Imports System.Xml.Serialization

''' <summary>
''' General purpose helper methods
''' </summary>
Public Class baUtil

#Region "Declares"

   ' Mimic Microsoft.VisualBasic.TriState
   Public Enum TriState
      [False] = 0
      [True] = -1
      [UseDefault] = -2
   End Enum

#End Region

#Region "Formatting (Strings / Numbers)"

   'Public Shared Function FormatNumber(ByVal Expression As Object,
   '                                    Optional ByVal NumDigitsAfterDecimal As Int32 = -1,
   '                                    Optional ByVal IncludeLeadingDigit As TriState = TriState.UseDefault,
   '                                    Optional ByVal UseParensForNegativeNumbers As TriState = TriState.UseDefault,
   '                                    Optional ByVal GroupDigits As TriState = TriState.UseDefault) As String

   '   Dim sFormatMask As String = String.Empty
   '   Dim curCulture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture

   '   Console.WriteLine(Expression.ToString)

   '   If NumDigitsAfterDecimal = -1 Then
   '      ' -1 = use system defaults

   '   End If

   'End Function

#End Region

End Class

Public Class MathUtil
   '------------------------------------------------------------------------------
   'File:    : baUtil.vb
   'Purpose  : General math helpers
   '
   'Prereq.  : -
   'Note     : -
   '
   '   Author: Knuth Konrad
   '     Date: 18.04.2019
   '   Source: -
   '  Changed: -
   '------------------------------------------------------------------------------

   Public Shared Function Percent(ByVal part As Double, ByVal total As Double) As Double
      '------------------------------------------------------------------------------
      'Purpose  : Returns the % of Total given by Part,
      '           e.g. Total = 200, Part = 50 = 25(%)
      'Param    : part  - Part to expressed as a precent value
      '           total - = 100%
      '
      'Prereq.  : -
      'Returns  : -
      'Note     : -
      '
      '   Author: Knuth Konrad
      '     Date: 15.08.2017
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------

      If total = 0 Then
         Throw New System.ArgumentOutOfRangeException("total", "Value can't be zero.")
      End If

      Return (part / total) * 100

   End Function

End Class

<DefaultProperty("Value")>
<Serializable()> Public Class FixedLengthString

   Private Const PADDING_CHAR As Char = " "c

   Private miLength As UInt16
   Private miMinLength As UInt16
   Private mcPaddingChar As Char = PADDING_CHAR
   Private msValue As String

   Public Property Length As UInt16
      Get
         Return miLength
      End Get
      Set(value As UInt16)
         miLength = value
      End Set
   End Property

   Public Property MinLength As UInt16
      Get
         Return miMinLength
      End Get
      Set(value As UInt16)
         miMinLength = value
      End Set
   End Property

   Public Property Value As String
      Get
         Return msValue
      End Get
      Set(value As String)
         With Me
            If (.MinLength > 0) And (value.Length < .MinLength) Then
               Throw New ArgumentOutOfRangeException("Value", "Length < MinLength")
            End If

            If value.Length < .MinLength Then
               msValue = value & New String(CType(" ", Char), .MinLength - value.Length)
            Else
               msValue = value.Substring(0, .MinLength)
            End If
         End With
      End Set
   End Property

   Public Property PaddingChar As Char
      Get
         Return mcPaddingChar
      End Get
      Set(value As Char)
         mcPaddingChar = value
      End Set
   End Property

   Public Function ToLower() As String
      Return Me.Value.ToLower
   End Function

   Public Function ToUpper() As String
      Return Me.Value.ToUpper
   End Function

   Public Sub New(ByVal stringLength As UInt16)

      MyBase.New
      With Me
         .Length = stringLength
         .Value = New String(.PaddingChar, stringLength)
      End With
   End Sub

   Public Sub New(ByVal stringLength As UInt16, ByVal stringValue As String)

      MyBase.New
      With Me
         .Length = stringLength
         .Value = stringValue
         If stringValue.Length < .Length Then
            .Value = stringValue.PadRight(.Length, .PaddingChar)
         End If
      End With

   End Sub

   Public Sub New(ByVal stringLength As UInt16, ByVal stringMinLength As UInt16)

      MyBase.New
      If (stringMinLength > 0) And (stringLength < stringMinLength) Then
         Throw New ArgumentOutOfRangeException("Length", "Length can't be < MinLength")
      End If

      With Me
         .Length = stringLength
         .MinLength = stringMinLength
         .Value = New String(.PaddingChar, stringLength)
      End With

   End Sub

   Public Sub New(ByVal stringLength As UInt16, ByVal stringValue As String, ByVal stringMinLength As UInt16)

      MyBase.New
      If (stringMinLength > 0) And (stringLength < stringMinLength) Then
         Throw New ArgumentOutOfRangeException("Length", "Length can't be < MinLength")
      End If

      With Me
         .Length = stringLength
         .MinLength = stringMinLength
         .Value = stringValue
         If stringValue.Length < .Length Then
            .Value = stringValue.PadRight(.Length, .PaddingChar)
         End If
      End With

   End Sub

End Class
