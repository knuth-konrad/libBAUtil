Imports System.Globalization

''' <summary>
''' General purpose string handling/formatting helpers
''' </summary>
Public Class StringUtil
   '------------------------------------------------------------------------------
   'Prereq.  : -
   '
   '   Author: Knuth Konrad
   '     Date: 13.11.2018
   '   Source: -
   '  Changed: -
   '------------------------------------------------------------------------------

#Region "Declares"

   ' Bytes to <unit> - Function Bytes2FormattedString()
   Enum eSizeUnits As Long
      B = 1024L
      KB = B * B
      MB = KB * B
      GB = MB * B
      TB = GB * B
   End Enum

#End Region

   ''' <summary>
   ''' Creates a formatted string representing the size in its proper 'spelled out' unit
   ''' (Bytes, KB etc.)
   ''' </summary>
   ''' <param name="uintBytes">Number of bytes to transform</param>
   ''' <returns>
   ''' Spelled out size, e.g. 1030 -> '1KB'
   ''' </returns>
   Public Overloads Shared Function Bytes2FormattedString(ByVal uintBytes As UInt64) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      'Note     : -
      '
      '   Author:  dbasnett
      '     Date: 07.06.2018
      '   Source: http://www.vbforums.com/showthread.php?634675-RESOLVED-Bytes-to-MB-etc
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim dblInUnits As Double
      Dim sUnits As String = String.Empty, szAsStr As String = String.Empty

      If uintBytes < eSizeUnits.B Then
         szAsStr = uintBytes.ToString("n0")
         sUnits = "Bytes"
      ElseIf uintBytes <= eSizeUnits.KB Then
         dblInUnits = uintBytes / eSizeUnits.B
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "KB"
      ElseIf uintBytes <= eSizeUnits.MB Then
         dblInUnits = uintBytes / eSizeUnits.KB
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "MB"
      ElseIf uintBytes <= eSizeUnits.GB Then
         dblInUnits = uintBytes / eSizeUnits.MB
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "GB"
      Else
         dblInUnits = uintBytes / eSizeUnits.GB
         szAsStr = dblInUnits.ToString("n1")
         sUnits = "TB"
      End If

      Return String.Format("{0} {1}", szAsStr, sUnits)

   End Function

   ''' <summary>
   ''' Creates a formatted string representing the size in its proper spelled out unit
   ''' (Bytes, KB etc.)
   ''' </summary>
   ''' <param name="uintBytes">Number of bytes to transform</param>
   ''' <param name="largestUnitOnly"><see langword="true"/>: return only the largest part, e.g. 5.5 GB -> 5GB</param>
   ''' <returns>
   ''' Spelled out size, e.g. 1030 -> '1KB'
   ''' </returns>
   Public Overloads Shared Function Bytes2FormattedString(ByVal uintBytes As UInt64,
                                                          Optional ByVal largestUnitOnly As Boolean = True) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 07.06.2018
      '   Source: -
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim uintDivisor As UInt64
      Dim sUnits As String = String.Empty, szAsStr As String = String.Empty

      If largestUnitOnly = True Then
         Return Bytes2FormattedString(uintBytes)
      End If

      Do While uintBytes > 0

         If (uintBytes \ CType(1024 ^ 4, UInt64)) > 0 Then
            ' TB
            uintDivisor = uintBytes \ CType(1024 ^ 4, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 4, UInt64))
            sUnits = "GB "
         ElseIf uintBytes \ CType(1024 ^ 3, UInt64) > 0 Then
            ' GB
            uintDivisor = uintBytes \ CType(1024 ^ 3, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 3, UInt64))
            sUnits = "GB "
         ElseIf uintBytes \ CType(1024 ^ 2, UInt64) > 0 Then
            ' MB
            uintDivisor = uintBytes \ CType(1024 ^ 2, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 2, UInt64))
            sUnits = "MB "
         ElseIf uintBytes \ CType(1024 ^ 1, UInt64) > 0 Then
            ' KB
            uintDivisor = uintBytes \ CType(1024 ^ 1, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 1, UInt64))
            sUnits = "KB "
         Else
            ' B
            uintDivisor = uintBytes \ CType(1024 ^ 0, UInt64)
            uintBytes = uintBytes - (uintDivisor * CType(1024 ^ 0, UInt64))
            sUnits = "B "
         End If

         szAsStr &= uintDivisor.ToString("n0") & sUnits

      Loop

      Return szAsStr.TrimEnd

   End Function

   ''' <summary>
   ''' Capitalize the first letter of a string.
   ''' </summary>
   ''' <param name="sText">Source string</param>
   ''' <param name="sCulture">Specific culture string e.g. "en-US"</param>
   ''' <returns>
   ''' <paramref name="sText"/> with the first letter capitalized.
   ''' </returns>
   Public Shared Function MCase(ByVal sText As String, Optional ByVal sCulture As String = "") As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 13.11.2018
      '   Source: https://social.msdn.microsoft.com/Forums/vstudio/en-US/c0872f6d-2975-43e6-872a-d2ba7901ed0e/convert-first-letter-of-string-to-capital?forum=csharpgeneral
      '  Changed: -
      '------------------------------------------------------------------------------
      Dim ti As TextInfo

      Try
         If sCulture.Length > 0 Then
            ti = New CultureInfo(sCulture, False).TextInfo
            Return ti.ToTitleCase(sText)
         Else
            Return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(sText)
         End If
      Catch
         Return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(sText)
      End Try

   End Function

   ''' <summary>
   ''' Implements VB6's Left$() functionality.
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="leftChars">Number of characters to return</param>
   ''' <returns>
   ''' For leftChars ...
   '''    &gt; source.Length: source
   '''    = 0: Empty string
   '''    &lt; 0: Position from the end of source, e.g. Left("1234567890", -2) -> "12345678"
   ''' </returns>
   Public Shared Function Left(ByVal source As String, ByVal leftChars As Integer) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 19.03.2019
      '   Source: Developed from https://stackoverflow.com/questions/844059/net-equivalent-of-the-old-vb-leftstring-length-function/12481156
      '  Changed: -
      '------------------------------------------------------------------------------

      If String.IsNullOrEmpty(source) OrElse leftChars = 0 Then
         Return String.Empty
      ElseIf leftChars > source.Length Then
         Return source
      ElseIf leftChars < 0 Then
         Return source.Substring(0, source.Length + leftChars)
      Else
         Return source.Substring(0, Math.Min(leftChars, CType(source.Length, Integer)))
      End If

   End Function

   ''' <summary>
   ''' Implements VB's/PB's Right$() functionality.
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="rightChars">Number of characters to return</param>
   ''' <returns>
   ''' For rightChars ...
   '''    &gt; source.Length: source
   '''    = 0: Empty string
   '''    &lt; 0: Position from the start of source, e.g. Right("1234567890", -2) -&gt; "34567890"
   ''' </returns>
   Public Shared Function Right(ByVal source As String, ByVal rightChars As Integer) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 19.03.2019
      '   Source: Developed from https://stackoverflow.com/questions/844059/net-equivalent-of-the-old-vb-leftstring-length-function/12481156
      '  Changed: 09.04.2019
      '           - Add exceptional case handling
      '------------------------------------------------------------------------------

      If String.IsNullOrEmpty(source) OrElse rightChars = 0 Then
         Return String.Empty
      ElseIf rightChars > source.Length Then
         Return source
      ElseIf rightChars < 0 Then
         Return source.Substring(Math.Abs(rightChars))
      Else
         Return source.Substring(source.Length - rightChars, rightChars)
      End If

   End Function

   ''' <summary>
   ''' Implements VB's/PB's Mid$() functionality, as .NET's String.SubString() 
   ''' differs in its behavior that it raises an exception if startIndex > source.Length, 
   ''' whereas Mid$() returns an empty string in such a case.
   ''' </summary>
   ''' <param name="source">Source string</param>
   ''' <param name="startIndex">(0-based) start</param>
   ''' <param name="length">Number of chars to return</param>
   ''' <returns>
   ''' For startIndex > source.Length: String.Empty
   ''' For length > startIndex + source.Length: all of source from startIndex
   ''' </returns>
   Public Shared Function Mid(ByVal source As String, ByVal startIndex As Integer, Optional ByVal length As Integer = 0) As String
      '------------------------------------------------------------------------------
      'Prereq.  : -
      '
      '   Author: Knuth Konrad
      '     Date: 26.04.2019
      '   Source: Developed from https://stackoverflow.com/questions/844059/net-equivalent-of-the-old-vb-leftstring-length-function/12481156
      '  Changed: -
      '------------------------------------------------------------------------------

      ' Safe guards
      If String.IsNullOrEmpty(source) OrElse (startIndex > source.Length - 1) Then
         Return String.Empty
      End If
      If startIndex < 0 Then
         Throw New ArgumentOutOfRangeException("startIndex")
      End If
      If length < 0 Then
         Throw New ArgumentOutOfRangeException("length")
      End If

      ' Adjust length, if needed
      Try
         If startIndex + length > source.Length - 1 OrElse length = 0 Then
            Return source.Substring(startIndex)
         Else
            Return source.Substring(startIndex, length)
         End If
      Catch ex As ArgumentOutOfRangeException
         Return String.Empty
      End Try

   End Function

   Public Overloads Shared Function Space(ByVal number As UInt32) As String
      Return New String(" "c, CType(number, Integer))
   End Function

   Public Overloads Shared Function Space(ByVal number As Int32) As String
      Return New String(" "c, CType(number, Integer))
   End Function

   ' ** Replacements for various handy VB6 string constants
   Public Shared Function vbNewLine() As String
      Return Environment.NewLine
   End Function

   Public Shared Function vbTab() As String
      Return System.Convert.ToChar(9).ToString
   End Function

End Class
