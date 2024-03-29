Imports System.Linq
Imports System.Reflection
Imports libBAUtil.StringHelper

''' <summary>
''' (NLog) logging helper methods
''' </summary>
Public Class LoggingHelper


  ''' <summary>
  ''' Return a list of parameters info passed to a method.
  ''' </summary>
  ''' <param name="mb">The current method.</param>
  ''' <param name="param">Parameters passed to the above method.</param>
  ''' <returns>Parameters</returns>
  ''' <remarks>
  ''' When passing the actual parameters in <paramref name="param"/>, ideally these 
  ''' should be passed in the order they appear in the method prototype.
  ''' </remarks>
  ''' <example>
  ''' Console.WriteLine(GetMethodParameters(System.Reflection.MethodBase.GetCurrentMethod()))
  ''' </example>
  Public Function GetMethodParameters(ByVal mb As MethodBase, ParamArray param As Object()) As String

    Dim pis As ParameterInfo() = mb.GetParameters()
    Dim result As String = "Parameters:" & vbNewLine()

    Try
      For Each pi As ParameterInfo In pis
        If pi.IsOut Then
          result &= String.Format(" - {0} (out)" & vbNewLine(), pi.Name)
        Else
          result &= String.Format(" - {0}" & vbNewLine(), pi.Name)
        End If

        result &= String.Format("  Type    : {0}" & vbNewLine(), pi.ParameterType)
        result &= String.Format("  Position: {0}" & vbNewLine(), pi.Position.ToString())

        If pi.HasDefaultValue Then
          result &= String.Format("  Default : {0}" & vbNewLine(), pi.DefaultValue.ToString())
        End If

        If (param IsNot Nothing) AndAlso param.GetLength(0) > 0 Then
          result &= "Values:" & vbNewLine()
          For Each o As Object In param
            result &= String.Format("  {0}" & vbNewLine(), o.ToString())
          Next
        End If
      Next
    Catch ex As Exception
      Return String.Empty
    End Try

    Return result

  End Function

  ''' <summary>
  ''' Retrieve the current (text) log file's name and path
  ''' </summary>
  ''' <returns>
  ''' >NLog's log file incl. full path
  ''' </returns>
  ''' <remarks>
  ''' Source: https://stackoverflow.com/questions/7332393/how-can-i-query-the-path-to-an-nlog-log-file
  ''' </remarks>
  Public Function GetNLogCurrentLogFile(Optional fileTargetName As String = "logfile") As String

    ' https://stackoverflow.com/questions/13393180/log-with-nlog-to-available-drive
    ' https://www.appsloveworld.com/linq/100/25/vb-net-linq-query-on-listof-object-source-code
    ' https://www.aspsnippets.com/questions/149948/Filter-List-Using-Linq-in-C-and-VBNet/
    ' https://www.tutlane.com/tutorial/linq/linq-to-lists-collections


    For Each ft As NLog.Targets.FileTarget In NLog.LogManager.Configuration.AllTargets
      If ft.Name = fileTargetName Then
        Return ft.FileName.Render(New NLog.LogEventInfo)
      End If
    Next

  End Function

End Class
