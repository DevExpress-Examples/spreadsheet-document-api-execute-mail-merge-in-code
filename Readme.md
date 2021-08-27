<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/128612985/16.2.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T515791)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
<!-- default file list -->
*Files to look at*:

* [Program.cs](./CS/MailMergeExample/Program.cs) (VB: [Program.vb](./VB/MailMergeExample/Program.vb))
<!-- default file list end -->
# Spreadsheet Document API - How to create a spreadsheet template in code and perform mail merge


This example demonstrates how to use the <a href="https://documentation.devexpress.com/OfficeFileAPI/118749/Spreadsheet-Document-API/Mail-Merge">Spreadsheet Mail Merge</a> functionality to automatically generate a document based on data retrieved from an object data source. <br>The <a href="https://documentation.devexpress.com/OfficeFileAPI/118747/Spreadsheet-Document-API/Mail-Merge/Template-Document">mail merge template</a> is created in code when the application starts. The data source is specified using the <a href="https://documentation.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook.MailMergeDataSource.property">Workbook.MailMergeDataSource</a> property. The <a href="https://documentation.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Workbook.GenerateMailMergeDocuments.method">Workbook.GenerateMailMergeDocuments</a> method accomplishes mail merge and returns the resulting workbook. Since the mail merge mode is set to <strong>Multiple Sheets</strong>, all worksheet created by the <strong>GenerateMailMergeDocuments</strong> method are contained in a single workbook.<br><br><img src="https://raw.githubusercontent.com/DevExpress-Examples/document-server-how-to-create-a-spreadsheet-template-in-code-and-perform-mail-merge-t515791/16.2.3+/media/299de11c-3b10-11e7-80c0-00155d624807.png">
<br/>
