<div align="center">

## Completely Reusable SQL Server/XML Generator \!


</div>

### Description

This is a really simple way of grabbing the column details for all the user tables in a SQL Server database, and converting those details to XML. If you wanted to get a little creative, you could re-use this with XSL to create SQL and ASP templates for your apps.
 
### More Info
 
The only real 'condition' is that you should have a default database set up in your connection.

I know this works in SQL Server 2000 and version 7.0, but I won't guarantee anything before that (I don't have copies of the system tables to check).

XML data outlining the column information for the tables in a database.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brandon McPherson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brandon-mcpherson.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brandon-mcpherson-completely-reusable-sql-server-xml-generator__4-6504/archive/master.zip)





### Source Code

```
<%
 Set Cnn1 = Server.CreateObject("ADODB.Connection")
 Set adoCmd = Server.CreateObject("ADODB.Command")
 Dim sXML
' Set the dsn up to a database
 Cnn1.Open "DSN=nwind;UID=sa;PWD="
 Set adoCmd.ActiveConnection = Cnn1
' I'm sure there's a much more graceful way to do this....
 adoCmd.CommandText = "select TableName=t1.[name], ColumnName=c1.[name], datatype=(select a1.[name] from dbo.systypes as a1 where a1.xusertype = c1.xusertype), c1.isnullable, c1.length, c1.colid from dbo.syscolumns as c1 inner join dbo.sysobjects as t1 on c1.id = t1.id where t1.xtype = 'U'"
 Set tmpRST = adoCmd.Execute
 sXML = "<?xml version=""1.0""?><tables>"
 Dim sTable
 Dim iColCount
 Dim intIterator
 iColCount = tmpRST.Fields.Count - 1
 sTable = tmpRST.Fields("TableName")
 sXML = sXML & "<table><name>" & tmpRST("TableName") & "</name>"
 Do While Not tmpRST.EOF
  If tmpRST.Fields("TableName") <> sTable Then
   sXML = sXML & "</table><table><name>" & tmpRST.Fields("TableName") & "</name>"
  End If
  sXML = sXML & "<column>"
  For intIterator = 1 To iColCount
   sXML = sXML & "<" & tmpRST.Fields(intIterator).Name & ">" & tmpRST.Fields(intIterator) & "</" & tmpRST.Fields(intIterator).Name & ">"
  Next
  sXML = sXML & "</column>"
  sTable = tmpRST.Fields("TableName")
  tmpRST.MoveNext
 Loop
 sXML = sXML & "</table></tables>"
 Response.Write sXML
%>
```

