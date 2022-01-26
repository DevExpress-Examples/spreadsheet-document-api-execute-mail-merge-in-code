Imports DevExpress.Spreadsheet
Imports System

Namespace MailMergeExample
    Friend NotInheritable Class Program

        Private Sub New()
        End Sub

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        <STAThread> _
        Shared Sub Main()
'			#Region "#main"
			Using workbook As New Workbook()
				workbook.Unit = DevExpress.Office.DocumentUnit.Inch
				workbook.BeginUpdate()
				' Create a mail merge template.
				Dim template As Worksheet = workbook.Worksheets(0)
				template.Rows(1).RowHeight = 1.5
				template.Columns(1).ColumnWidth = 1.0
				template.Columns(1).Alignment.Vertical = SpreadsheetVerticalAlignment.Center
				template.Columns(2).ColumnWidth = 2.5
				template.Columns(2).Alignment.WrapText = True
				template.Cells("C2").Formula = "FIELDPICTURE(""Photo"", ""range"", C2, FALSE, 50)"
				template.Cells("C3").Formula = "=FIELD(""FirstName"")&"" ""&FIELD(""LastName"")"
				template.Cells("B4").Value = "Position:"
				template.Cells("C4").Formula = "FIELD(""Title"")"
				template.Cells("B5").Value = "Birth Date:"
				template.Cells("C5").Formula = "FIELD(""BirthDate"")"
				template.Cells("C5").NumberFormat = "M/d/yyyy"
				template.Cells("B6").Value = "Hire Date:"
				template.Cells("C6").Formula = "FIELD(""HireDate"")"
				template.Cells("C6").NumberFormat = "dddd MMMM dd, yyyy"
				template.Cells("B7").Value = "Home Phone:"
				template.Cells("C7").Formula = "FIELD(""HomePhone"")"
				template.Cells("B8").Value = "Address:"
				template.Cells("C8").Formula = "=FIELD(""Address"")&"" ""&FIELD(""City"")"
				template.Cells("B9").Value = "About:"
				template.Cells("C9").Formula = "FIELD(""Notes"")"
				workbook.EndUpdate()

				' Set a detail range in the template.
				Dim detail As CellRange = template.Range("C1:C9")
				detail.Name = "DETAILRANGE"

				' Set a header range in the template.
				Dim header As CellRange = template.Range("B1:B9")
				header.Name = "HEADERRANGE"

				' Switch the mail merge mode to "Multiple Sheets".
				workbook.DefinedNames.Add("MAILMERGEMODE", "=""Worksheets""")
				' Switch the mail merge mode to "Multiple Documents".
				'workbook.DefinedNames.GetDefinedName("MAILMERGEMODE").RefersTo = "\"Documents\"";
				' Switch the mail merge mode to "Single Sheet".
				'workbook.DefinedNames.GetDefinedName("MAILMERGEMODE").RefersTo = "\"OneWorksheet\"";

				' Set vertical document orientation.
				workbook.DefinedNames.Add("HORIZONTALMODE", "=TRUE")

				' Perform mail merge.
				workbook.MailMergeDataSource = EmployeeInfo.EmployeesInfo.GetData()
				Dim result = workbook.GenerateMailMergeDocuments()
				result(0).SaveDocument("result.xlsx")
				System.Diagnostics.Process.Start("result.xlsx")
			End Using
			'			#End Region ' #main
        End Sub
    End Class
End Namespace
