Imports System.Data.SqlClient

Namespace DBTools.SQL

   Public Class SqlDBUtil

#Region "Methods SqlClient"

      Public Shared Function GetDBString(ByVal dr As SqlDataReader, ByVal sColumn As String) As String

         Try
            With dr
               If .IsDBNull(.GetOrdinal(sColumn)) Then
                  Return String.Empty
               Else
                  Return CType(.Item(sColumn), String)
               End If
            End With
         Catch ex As Exception
            Return String.Empty
         End Try

      End Function

      Public Shared Function GetDBInteger(ByVal dr As SqlDataReader, ByVal sColumn As String) As Integer

         Try
            With dr
               If .IsDBNull(.GetOrdinal(sColumn)) Then
                  Return 0
               Else
                  Return CType(.Item(sColumn), Integer)
               End If
            End With
         Catch ex As Exception
            Return 0
         End Try

      End Function

      Public Shared Function GetDBByte(ByVal dr As SqlDataReader, ByVal sColumn As String) As Byte

         Try
            With dr
               If .IsDBNull(.GetOrdinal(sColumn)) Then
                  Return 0
               Else
                  Return CType(.Item(sColumn), Byte)
               End If
            End With
         Catch ex As Exception
            Return 0
         End Try

      End Function

      Public Shared Function GetDBDate(ByVal dr As SqlDataReader, ByVal sColumn As String) As Date

         Try
            With dr
               If .IsDBNull(.GetOrdinal(sColumn)) Then
                  Return Nothing
               Else
                  Return CType(.Item(sColumn), Date)
               End If
            End With
         Catch ex As Exception
            Return CDate(Nothing)
         End Try

      End Function

      Public Shared Function GetDBGuid(ByVal dr As SqlDataReader, ByVal sColumn As String) As Guid

         Try
            With dr
               If .IsDBNull(.GetOrdinal(sColumn)) Then
                  Return Nothing
               Else
                  Return CType(.Item(sColumn), Guid)
               End If
            End With
         Catch ex As Exception
            Return Nothing
         End Try

      End Function

      ''' <summary>
      ''' Reads fieldName from Data Reader. If fieldName is DbNull, returns String.Empty.
      ''' </summary>
      ''' <returns>Safely returns a string. No need to check for DbNull.</returns>
      Public Overloads Shared Function ReadNullAsEmptyString(ByVal reader As IDataReader, ByVal fieldName As String) As String
         ' Source: https://stackoverflow.com/questions/20862628/how-to-deal-with-sqldatareader-null-values-in-vb-net
         With reader
            If .IsDBNull(.GetOrdinal(fieldName)) Then
               Return String.Empty
            Else
               Return .Item(fieldName).ToString
            End If
         End With
      End Function

      ''' <summary>
      ''' Reads fieldOrdinal from Data Reader. If fieldOrdinal is DbNull, returns String.Empty.
      ''' </summary>
      ''' <returns>Safely returns a string. No need to check for DbNull.</returns>
      Public Overloads Shared Function ReadNullAsEmptyString(ByVal reader As IDataReader, ByVal fieldOrdinal As Integer) As String
         ' Source: https://stackoverflow.com/questions/20862628/how-to-deal-with-sqldatareader-null-values-in-vb-net
         With reader
            If .IsDBNull(fieldOrdinal) Then
               Return String.Empty
            Else
               Return reader.Item(fieldOrdinal).ToString
            End If
         End With
      End Function


#End Region

      Public Sub New()
         MyBase.New()
      End Sub

   End Class

   Public Class cObjectDescription
   End Class

End Namespace
