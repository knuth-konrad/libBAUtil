''' <summary>
''' General Date/Time helper methods
''' </summary>
Public Class DateTimeUtil

   ''' <summary>
   ''' Return the last day in a month
   ''' </summary>
   ''' <param name="month">Last day of this month</param>
   ''' <param name="year">Month in this year</param>
   ''' <returns></returns>
   Public Overloads Shared Function GetLastDayInMonth(ByVal month As Int32, ByVal year As Int32) As DateTime

      If month = 12 Then
         month = 1
      Else
         month += 1
      End If

      Dim dtm As New DateTime(year, month, 1)

      Dim tsp As New TimeSpan(1, 0, 0, 0)
      Return dtm.Subtract(tsp)

   End Function

End Class
