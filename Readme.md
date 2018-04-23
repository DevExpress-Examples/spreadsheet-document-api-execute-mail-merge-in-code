# Document Server- How to create a spreadsheet template in code and perform mail merge


This example demonstrates how to use the <a href="http://help.devexpress.com/#DocumentServer/CustomDocument118749">Spreadsheet Mail Merge</a> functionality to automatically generate a document based on data retrieved from an object data source. <br>The <a href="http://help.devexpress.com/#DocumentServer/CustomDocument118747">mail merge template</a> is created in code when the application starts. The data source is specified using the <a href="http://help.devexpress.com/#DocumentServer/DevExpressSpreadsheetWorkbook_MailMergeDataSourcetopic">Workbook.MailMergeDataSource</a> property. The <a href="http://help.devexpress.com/#DocumentServer/DevExpressSpreadsheetWorkbook_GenerateMailMergeDocumentstopic">Workbook.GenerateMailMergeDocuments</a> method accomplishes mail merge and returns the resulting workbook. Since the mail merge mode is set to <strong>Multiple Sheets</strong>, all worksheet created by the <strong>GenerateMailMergeDocuments</strong> method are contained in a single workbook.<br><br><img src="https://raw.githubusercontent.com/DevExpress-Examples/document-server-how-to-create-a-spreadsheet-template-in-code-and-perform-mail-merge-t515791/16.2.3+/media/299de11c-3b10-11e7-80c0-00155d624807.png">

<br/>


