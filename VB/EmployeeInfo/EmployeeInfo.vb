Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Drawing
Imports System.IO
Imports System.Xml.Serialization

Namespace EmployeeInfo
	Public Module EmployeesInfo
		Public Function GetData() As IEnumerable(Of EmployeeInfo)
			Return DataSourceHelper.GetDataSouresFromXml(Of EmployeeInfo)("EmployeeInfo.data.xml", "ArrayOfEmployeeInfo")
		End Function
	End Module

	Public Module DataSourceHelper
		Public Function GetDataSouresFromXml(Of T As Class)(ByVal resName As String, ByVal attribute As String) As List(Of T)
			Using stream As Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resName)
				Dim s As New XmlSerializer(GetType(List(Of T)), New XmlRootAttribute(attribute))
				Return DirectCast(s.Deserialize(stream), List(Of T))
			End Using
		End Function
	End Module

	Public Class EmployeeInfo
		Public Property LastName() As String
		Public Property FirstName() As String
		Public Property Title() As String
		Public Property TitleOfCourtesy() As String
		Public Property BirthDate() As Date
		Public Property HireDate() As Date
		Public Property Address() As String
		Public Property City() As String
		Public Property Region() As String
		Public Property PostalCode() As String
		Public Property Country() As String
		Public Property HomePhone() As String
		Public Property Extension() As String
		<XmlIgnoreAttribute()>
		Public Property Photo() As Bitmap
		<XmlElementAttribute("Photo")>
		Public Property PictureByteArray() As Byte()
			Get
				If Photo IsNot Nothing Then
					Dim BitmapConverter As TypeConverter = TypeDescriptor.GetConverter(Photo.GetType())
					Return DirectCast(BitmapConverter.ConvertTo(Photo, GetType(Byte())), Byte())
				Else
					Return Nothing
				End If
			End Get
			Set(ByVal value As Byte())
				If value IsNot Nothing Then
					Photo = New Bitmap(New MemoryStream(value))
				Else
					Photo = Nothing
				End If
			End Set
		End Property
		Public Property Notes() As String
	End Class

End Namespace
