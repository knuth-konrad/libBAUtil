Imports System
Imports System.Runtime.InteropServices
Imports ComTypes = System.Runtime.InteropServices.ComTypes

''' <summary>
''' COM objects helper class.
''' </summary>
Public Class ComHelper
	''' <summary>
	''' Returns a string value representing the type name of the specified COM object.
	''' </summary>
	''' <param name="comObj">A COM object the type name of which to return.</param>
	''' <returns>A string containing the type name.</returns>
	''' <remarks>
	''' Source: https://www.add-in-express.com/creating-addins-blog/2011/12/20/type-name-system-comobject/
	''' </remarks>
	Public Shared Function GetTypeName(ByVal comObj As Object) As String

		If comObj Is Nothing Then
			Return String.Empty
		End If

		' The specified object is not a COM object
		If Not Marshal.IsComObject(comObj) Then
			Return String.Empty
		End If

		Dim dispatch As IDispatch = TryCast(comObj, IDispatch)
		' The specified COM object doesn't support getting type information
		If dispatch Is Nothing Then
			Return String.Empty
		End If

		Dim typeInfo As ComTypes.ITypeInfo = Nothing
		Try
			Try
				' Obtain the ITypeInfo interface from the object
				dispatch.GetTypeInfo(0, 0, typeInfo)
			Catch ex As Exception
				' Cannot get the ITypeInfo interface for the specified COM object
				Return String.Empty
			End Try

			Dim typeName As String = String.Empty
			Dim documentation As String = String.Empty, helpFile As String = String.Empty
			Dim helpContext As Int32 = -1

			Try
				' Retrieves the documentation string for the specified type description 
				typeInfo.GetDocumentation(-1, typeName, documentation, helpContext, helpFile)
			Catch ex As Exception
				' Cannot extract ITypeInfo information
				Return String.Empty
			End Try
			Return typeName
		Catch ex As Exception
			' Unexpected error
			Return String.Empty
		Finally
			If typeInfo IsNot Nothing Then Marshal.ReleaseComObject(typeInfo)
		End Try
	End Function
End Class

''' <summary>
''' Exposes objects, methods and properties to programming tools and other
''' applications that support Automation.
''' </summary>
<ComImport()> <Guid("00020400-0000-0000-C000-000000000046")> <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IDispatch
	<PreserveSig>
	Function GetTypeInfoCount(<Out> ByRef Count As Integer) As Integer

	<PreserveSig>
	Function GetTypeInfo(
		<MarshalAs(UnmanagedType.U4)> ByVal iTInfo As Integer,
		<MarshalAs(UnmanagedType.U4)> ByVal lcid As Integer,
		<Out> ByRef typeInfo As ComTypes.ITypeInfo
		) As Integer

	<PreserveSig>
	Function GetIDsOfNames(
		ByRef riid As Guid,
		<MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.LPWStr)> ByVal rgsNames As String(),
		ByVal cNames As Integer,
		ByVal lcid As Integer,
		<MarshalAs(UnmanagedType.LPArray)> ByVal rgDispId As Integer()
		) As Integer

	<PreserveSig>
	Function Invoke(
		ByVal dispIdMember As Integer,
		ByRef riid As Guid,
		ByVal lcid As UInteger,
		ByVal wFlags As UShort,
		ByRef pDispParams As ComTypes.DISPPARAMS,
		<Out> ByRef pVarResult As Object,
		ByRef pExcepInfo As ComTypes.EXCEPINFO,
		ByVal pArgErr As IntPtr()
		) As Integer

End Interface
