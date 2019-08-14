Option Explicit On
Option Strict On

Imports System.Data.OleDb

Namespace DBTools.Ole

   Public Class OleDBUtil

#Region "Methods OleDb"

      Public Function GetDBString(ByVal dr As OleDbDataReader, ByVal sColumn As String) As String

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

      Public Function GetDBInteger(ByVal dr As OleDbDataReader, ByVal sColumn As String) As Integer

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

      Public Function GetDBByte(ByVal dr As OleDbDataReader, ByVal sColumn As String) As Byte

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

      Public Function GetDBDate(ByVal dr As OleDbDataReader, ByVal sColumn As String) As Date

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

#End Region

      Public Sub New()
         MyBase.New()
      End Sub

   End Class

End Namespace
