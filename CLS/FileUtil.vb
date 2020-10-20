Imports System.IO
Imports System.Text

''' <summary>
''' Simply text file manipulation.
''' </summary>
Public Class TextFileUtil

   Public Shared Function TxtReadFile(ByVal textFile As String) As String

      Dim file As String = String.Empty

      Using reader As New StreamReader(textFile)
         file = reader.ReadToEnd
      End Using

      Return file

   End Function

   Public Shared Function TxtReadLine(ByVal textFile As String) As String

      Dim line As String = String.Empty

      Using reader As New StreamReader(textFile)
         line = reader.ReadLine()
      End Using

      Return line

   End Function

   Public Shared Sub TxtWriteLine(ByVal textFile As String, ByVal textLine As String, Optional ByVal doAppend As Boolean = True)

      Using writer As New StreamWriter(textFile, doAppend)
         writer.WriteLine(textLine)
      End Using

   End Sub

End Class
