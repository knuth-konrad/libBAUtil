Imports libBAUtil.StringUtil

''' <summary>
''' General purpose console application helpers
''' </summary>
Public Class ConsoleUtil

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="appName">Name of the application</param>
   ''' <param name="versionMajor">Major version</param>
   ''' <param name="versionMinor">Minor version</param>
   ''' <param name="versionRevision">Revision</param>
   ''' <param name="versionBuild">Build</param>
   Public Overloads Shared Sub ConHeadline(ByVal appName As String, ByVal versionMajor As Integer,
                                           Optional ByVal versionMinor As Integer = 0,
                                           Optional ByVal versionRevision As Int32 = 0,
                                           Optional versionBuild As Int32 = 0)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & appName & " v" &
                        versionMajor.ToString & "." &
                        versionMinor.ToString & "." &
                        versionRevision.ToString & "." &
                        versionBuild.ToString &
                        " *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="appName">Name of the application</param>
   Public Overloads Shared Sub ConHeadline(ByVal appName As String)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & appName & " *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   ''' <summary>
   ''' Display an application intro
   ''' </summary>
   ''' <param name="appName">Name of the application</param>
   ''' <param name="versionMajor">Major version</param>
   Public Overloads Shared Sub ConHeadline(ByVal appName As String, ByVal versionMajor As Integer)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & appName & " v" & versionMajor.ToString & ".0 *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   ''' <summary>
   ''' Display a copyright notice.
   ''' </summary>
   Public Overloads Shared Sub ConCopyright()

      Console.WriteLine("Copyright " & Chr(169) & " " & DateTime.Now.Year.ToString & " by STA Travel GmbH. All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   ''' <summary>
   ''' Display a copyright notice.
   ''' <param name="year">Copyrighted in year</param>
   ''' <param name="companyName">Copyright owner</param>
   ''' </summary>
   Public Overloads Shared Sub ConCopyright(ByVal year As String, ByVal companyName As String)

      Console.WriteLine("Copyright " & Chr(169) & " " & year & " by " & companyName & ". All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   ''' <summary>
   ''' Display a copyright notice.
   ''' <param name="companyName">Copyright owner</param>
   ''' </summary>
   Public Overloads Shared Sub ConCopyright(ByVal companyName As String)

      Console.WriteLine("Copyright " & Chr(169) & " " & DateTime.Now.Year.ToString & " by " & companyName & ". All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   ''' <summary>
   ''' Pauses the program execution and waits for a key press
   ''' </summary>
   ''' <param name="waitMessage">Pause message</param>
   ''' <param name="blankLinesBefore">Number of blank lines before the message</param>
   ''' <param name="blankLinesAfter">Number of blank lines after the message</param>
   Public Shared Sub AnyKey(Optional ByVal waitMessage As String = "-- Press any key --",
                            Optional ByVal blankLinesBefore As Int32 = 0,
                            Optional ByVal blankLinesAfter As Int32 = 0)

      BlankLine(blankLinesBefore)
      Console.WriteLine(waitMessage)
      BlankLine(blankLinesAfter)
      Console.ReadLine()

   End Sub

   ''' <summary>
   ''' Insert a blank line at the current position.
   ''' </summary>
   ''' <param name="blankLines">Number of blank lines to insert.</param>
   Public Shared Sub BlankLine(Optional ByVal blankLines As Int32 = 1)

      ' Safe guard
      If blankLines < 1 Then
         Exit Sub
      End If

      For i As Int32 = 1 To blankLines
         Console.WriteLine("")
      Next

   End Sub

End Class
