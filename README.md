<div align="center">

## Multiple RecordSets


</div>

### Description

with this procedure we can handle the queries or stored procedures which returns multiple recordsets - through ADO's
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Murty Yellapantula](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/murty-yellapantula.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/murty-yellapantula-multiple-recordsets__1-6490/archive/master.zip)





### Source Code

```
Public Sub MultipleRecordSets()
Dim AdoConn As Object
Dim AdoRs As Object
Dim I As Integer
Set AdoConn = CreateObject("ADODB.Connection")
AdoConn.Open ConnectionString
'stored procedure which returns multiple record sets
ssql = "StoredProcedure Parameter1, Parameter2, ... "
Set AdoRs = AdoConn.Execute(ssql)
Do Until AdoRs Is Nothing
  While Not AdoRs.EOF
    For I = 0 To AdoRs.Fields.Count - 1
      Debug.Print AdoRs.Fields(I)
    Next I
    AdoRs.MoveNext
  Wend
  Set AdoRs = AdoRs.NextRecordset
Loop
End Sub
```

