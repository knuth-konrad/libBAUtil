Imports libBAUtil.StringUtil

''' <summary>
''' General purpose console application helpers
''' </summary>
Public Class ConsoleUtil


   Public Overloads Shared Sub ConHeadline(ByVal sAppName As String, ByVal iVersionMajor As Integer,
                                           Optional ByVal iVersionMinor As Integer = 0,
                                           Optional ByVal iVersionRevision As Int32 = 0,
                                           Optional iVersionBuild As Int32 = 0)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & sAppName & " v" &
                        iVersionMajor.ToString & "." &
                        iVersionMinor.ToString & "." &
                        iVersionRevision.ToString & "." &
                        iVersionBuild.ToString &
                        " *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   Public Overloads Shared Sub ConHeadline(ByVal sProgName As String)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & sProgName & " *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   Public Overloads Shared Sub ConHeadline(ByVal sProgName As String, ByVal iMajorVersion As Integer)

      Console.ForegroundColor = ConsoleColor.White
      Console.WriteLine("* " & sProgName & " v" & CType(iMajorVersion, String) & ".0 *")
      Console.ForegroundColor = ConsoleColor.Gray

   End Sub

   Public Overloads Shared Sub ConCopyright()

      Console.WriteLine("Copyright " & Chr(169) & " " & DateTime.Now.Year.ToString & " by STA Travel GmbH. All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   Public Overloads Shared Sub ConCopyright(ByVal sYear As String, ByVal sCompanyname As String)

      Console.WriteLine("Copyright " & Chr(169) & " " & sYear & " by " & sCompanyname & ". All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   Public Overloads Shared Sub ConCopyright(ByVal sCompanyname As String)

      Console.WriteLine("Copyright " & Chr(169) & " " & DateTime.Now.Year.ToString & " by " & sCompanyname & ". All rights reserved.")
      Console.WriteLine("Written by Knuth Konrad")

   End Sub

   Public Shared Sub AnyKey(Optional ByVal waitMessage As String = "Press any key.",
                            Optional ByVal blankLinesBefore As Int32 = 0,
                            Optional ByVal blankLinesAfter As Int32 = 0)

      BlankLine(blankLinesBefore)
      Console.WriteLine(waitMessage)
      BlankLine(blankLinesAfter)
      Console.ReadLine()

   End Sub


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
