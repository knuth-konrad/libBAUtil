Imports System.IO
Imports System.Text

Public Class TextFileUtil

   Public Shared Sub TxtWriteLine(ByVal textFile As String, ByVal textLine As String, Optional ByVal doAppend As Boolean = True)

      Using writer As New StreamWriter(textFile, doAppend)
         writer.WriteLine(textLine)
      End Using

   End Sub

   Public Shared Function TxtReadLine(ByVal textFile As String) As String

      Dim line As String

      Using reader As New StreamReader(textFile)
         line = reader.ReadLine()
      End Using

      Return line

   End Function

End Class
