Imports System.Data.SqlClient

Namespace DBTools.SQL

   ''' <summary>
   ''' ADO.NET helper for SqlClient provider.
   ''' </summary>
   Public Class SqlDBUtil

#Region "Methods SqlClient"

      ''' <summary>
      ''' Retrieve the value of a <see cref="SqlDbType.TinyInt"/> column.
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnOrdinal">Column ordinal</param>
      ''' <returns>Column value as <see cref="Byte"/></returns>
      Public Overloads Shared Function GetDBByte(ByVal dr As SqlDataReader, ByVal columnOrdinal As Int32) As Byte

         Try
            With dr
               If .IsDBNull(columnOrdinal) Then
                  Return 0
               Else
                  Return CType(.Item(columnOrdinal), Byte)
               End If
            End With
         Catch ex As Exception
            Return 0
         End Try

      End Function

      ''' <summary>
      ''' Retrieve the value of a <see cref="SqlDbType.TinyInt"/> column.
      ''' </summary>
      ''' <param name="dr">Instantiated SqlDataReader</param>
      ''' <param name="columName">Column name</param>
      ''' <returns>Column value as <see cref="Byte"/></returns>
      Public Overloads Shared Function GetDBByte(ByVal dr As SqlDataReader, ByVal columName As String) As Byte
         Return GetDBByte(dr, dr.GetOrdinal(columName))
      End Function

      ''' <summary>
      ''' Retrieve the value of a <see cref="SqlDbType.DateTime"/> column.
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnOrdinal">Column ordinal</param>
      ''' <returns>Column value as <see cref="DateTime"/></returns>
      Public Overloads Shared Function GetDBDate(ByVal dr As SqlDataReader, ByVal columnOrdinal As Int32) As DateTime

         Try
            With dr
               If .IsDBNull(columnOrdinal) Then
                  Return Nothing
               Else
                  Return CType(.Item(columnOrdinal), DateTime)
               End If
            End With
         Catch ex As Exception
            Return CType(Nothing, DateTime)
         End Try

      End Function

      ''' <summary>
      ''' Retrieve the value of a <see cref="SqlDbType.DateTime"/> column.
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columName">Column name</param>
      ''' <returns>Column value as <see cref="DateTime"/></returns>
      Public Overloads Shared Function GetDBDate(ByVal dr As SqlDataReader, ByVal columName As String) As DateTime
         Return GetDBDate(dr, dr.GetOrdinal(columName))
      End Function

      ''' <summary>
      ''' Retrieve a GUID stored in a integer type column.
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnOrdinal">Column ordinal</param>
      ''' <returns>Column value as <see cref="Guid"/></returns>
      Public Overloads Shared Function GetDBGuid(ByVal dr As SqlDataReader, ByVal columnOrdinal As Int32) As Guid

         Try
            With dr
               If .IsDBNull(columnOrdinal) Then
                  Return Nothing
               Else
                  Return CType(.Item(columnOrdinal), Guid)
               End If
            End With
         Catch ex As Exception
            Return Nothing
         End Try

      End Function

      ''' <summary>
      ''' Retrieve a GUID stored in a integer type column.
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnName">Column name</param>
      ''' <returns>Column value as <see cref="Guid"/></returns>
      Public Overloads Shared Function GetDBGuid(ByVal dr As SqlDataReader, ByVal columnName As String) As Guid
         Return GetDBGuid(dr, dr.GetOrdinal(columnName))
      End Function

      ''' <summary>
      ''' Retrieve the value of an integer type column with a column size 
      ''' less than or equal to <see cref="SqlDbType.Int"/>
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnOrdinal">Column ordinal</param>
      ''' <returns>Column value as <see cref="Int32"/></returns>
      Public Overloads Shared Function GetDBInteger(ByVal dr As SqlDataReader, ByVal columnOrdinal As Int32) As Int32

         Try
            With dr
               If .IsDBNull(columnOrdinal) Then
                  Return 0
               Else
                  Return CType(.Item(columnOrdinal), Int32)
               End If
            End With
         Catch ex As Exception
            Return 0
         End Try

      End Function

      ''' <summary>
      ''' Retrieve the value of an integer type column with a column size 
      ''' less than or equal to <see cref="SqlDbType.Int"/>
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columName">Column name</param>
      ''' <returns>Column value as <see cref="Int32"/></returns>
      Public Overloads Shared Function GetDBInteger(ByVal dr As SqlDataReader, ByVal columName As String) As Int32
         Return GetDBInteger(dr, dr.GetOrdinal(columName))
      End Function

      ''' <summary>
      ''' Retrieve the value of an integer type column with a column size 
      ''' less than or equal to <see cref="SqlDbType.Int"/>
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnOrdinal">Column ordinal</param>
      ''' <returns>Column value as <see cref="Int32"/></returns>
      Public Overloads Shared Function GetDBInteger32(ByVal dr As SqlDataReader, ByVal columnOrdinal As Int32) As Int32
         Return GetDBInteger(dr, columnOrdinal)
      End Function

      ''' <summary>
      ''' Retrieve the value of an integer type column with a column size 
      ''' less than or equal to <see cref="SqlDbType.Int"/>
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columName">Column name</param>
      ''' <returns>Column value as <see cref="Int32"/></returns>
      Public Overloads Shared Function GetDBInteger32(ByVal dr As SqlDataReader, ByVal columName As String) As Int32
         Return GetDBInteger(dr, dr.GetOrdinal(columName))
      End Function

      ''' <summary>
      ''' Retrieve the value of an integer type column with a column size 
      ''' less than or equal to <see cref="SqlDbType.BigInt"/>
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columnOrdinal">Column ordinal</param>
      ''' <returns>Column value as <see cref="Int64"/></returns>
      Public Overloads Shared Function GetDBInteger64(ByVal dr As SqlDataReader, ByVal columnOrdinal As Int32) As Int64

         Try
            With dr
               If .IsDBNull(columnOrdinal) Then
                  Return 0
               Else
                  Return CType(.Item(columnOrdinal), Int64)
               End If
            End With
         Catch ex As Exception
            Return 0
         End Try

      End Function


      ''' <summary>
      ''' Retrieve the value of an integer type column with a column size 
      ''' less than or equal to <see cref="SqlDbType.BigInt"/>
      ''' </summary>
      ''' <param name="dr">Instantiated <see cref="SqlDataReader"/></param>
      ''' <param name="columName">Column name</param>
      ''' <returns>Column value as <see cref="Int64"/></returns>
      Public Overloads Shared Function GetDBInteger64(ByVal dr As SqlDataReader, ByVal columName As String) As Int64
         Return GetDBInteger64(dr, dr.GetOrdinal(columName))
      End Function

      Public Overloads Shared Function GetDBString(ByVal dr As SqlDataReader, ByVal columOrdinal As Int32) As String

         Try
            With dr
               If .IsDBNull(columOrdinal) Then
                  Return String.Empty
               Else
                  Return CType(.Item(columOrdinal), String)
               End If
            End With
         Catch ex As Exception
            Return String.Empty
         End Try

      End Function

      Public Overloads Shared Function GetDBString(ByVal dr As SqlDataReader, ByVal columName As String) As String
         Return GetDBString(dr, dr.GetOrdinal(columName))
      End Function

      Public Overloads Shared Function GetDBVarBinaryAsString(ByVal dr As SqlDataReader, ByVal columOrdinal As Int32) As String

         Try
            With dr
               If .IsDBNull(columOrdinal) Then
                  Return String.Empty
               Else
                  Dim lLength As Int32 = CType(.GetBytes(columOrdinal, 0, Nothing, 0, 255), Int32)
                  Dim b(lLength - 1) As Byte
                  Dim lRet As Int32 = CType(.GetBytes(columOrdinal, 0, b, 0, lLength), Int32)
                  Return System.Text.Encoding.Unicode.GetString(b)
               End If
            End With
         Catch ex As Exception
            Return String.Empty
         End Try

      End Function

      Public Overloads Shared Function GetDBVarBinaryAsString(ByVal dr As SqlDataReader, ByVal columName As String) As String
         Return GetDBVarBinaryAsString(dr, dr.GetOrdinal(columName))
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
