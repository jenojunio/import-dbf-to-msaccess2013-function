'
'You call it like this:
'
'Private Sub Command0_Click()
'    ImportDBF "C:\CustomDBFTablesDirectory", "DB_TABLENAME"
'End Sub
'
'It will create a new local access table with name same as your inputted dbftablename
'
'You can run this without referencing ADO Library
'


Function ImportDBF(ByVal dbfFileDir As String, _
                    ByVal dbfTableName As String)

  
    dbfFileDir = dbfFileDir & "\\"

    Dim dbfCn As Object
    
    Dim dbfRs As Object
    
    Dim dbfStrSql As String
    
    Dim dbfStrConnection As String
    
    Set dbfCn = CreateObject("ADODB.Connection")
    
    dbfStrConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & dbfFileDir & ";" & _
        "Extended Properties=dBase IV"
        
    dbfStrSql = "SELECT * FROM " & dbfTableName & ""
        
    dbfCn.Open dbfStrConnection
    
    Set dbfRs = dbfCn.Execute(dbfStrSql)

    Dim fieldIndex As Integer
    
    Dim ddlNewAccessTable As String
    Dim ddlColumns As String
    
    Dim dmlInsert As String
    Dim dmlColumns As String
    Dim dmlValues As String
    
    dmlColumns = "("
    ddlColumns = "("
    
    For fieldIndex = 0 To dbfRs.Fields.Count - 1
        
        dmlColumns = dmlColumns & dbfRs.Fields(fieldIndex).Name & ","
        
        Select Case dbfRs.Fields(fieldIndex).Type
            
            Case 202
               ddlColumns = ddlColumns & dbfRs.Fields(fieldIndex).Name & " " & _
                            "TEXT,"
            Case 203
                ddlColumns = ddlColumns & dbfRs.Fields(fieldIndex).Name & " " & _
                            "MEMO,"
            Case 5
                ddlColumns = ddlColumns & dbfRs.Fields(fieldIndex).Name & " " & _
                            "DOUBLE,"
            Case 7
                ddlColumns = ddlColumns & dbfRs.Fields(fieldIndex).Name & " " & _
                            "DATETIME,"
            Case 11
                ddlColumns = ddlColumns & dbfRs.Fields(fieldIndex).Name & " " & _
                            "YESNO,"
            Case Else
                ddlColumns = ddlColumns & dbfRs.Fields(fieldIndex).Name & " " & _
                            "TEXT,"
        End Select
        
    Next fieldIndex
    
    dmlColumns = Left(dmlColumns, Len(dmlColumns) - 1) & ")"
           
    ddlColumns = Left(ddlColumns, Len(ddlColumns) - 1) & ")"
    
    ddlNewAccessTable = "CREATE TABLE " & dbfTableName & " " & ddlColumns & ";"
    
    Dim myDb As Database
    Set myDb = CurrentDb()
    myDb.Execute ddlNewAccessTable
         
   
    
    Dim fieldIndex2 As Integer
    
    While Not dbfRs.EOF
    dmlInsert = ""
    dmlValues = "("
    
        For fieldIndex2 = 0 To dbfRs.Fields.Count - 1
            Select Case dbfRs(fieldIndex2).Type
                Case 202
                     dmlValues = dmlValues & "'" & dbfRs(fieldIndex2).Value & "',"
                Case 203
                    dmlValues = dmlValues & "'" & dbfRs(fieldIndex2).Value & "',"
                Case 5
                    dmlValues = dmlValues & dbfRs(fieldIndex2).Value & ","
                Case 11
                    dmlValues = dmlValues & dbfRs(fieldIndex2).Value & ","
                Case 7
                    If IsDate(dbfRs(fieldIndex2).Value) Then
                        dmlValues = dmlValues & "#" & dbfRs(fieldIndex2).Value & "#,"
                    Else
                        dmlValues = dmlValues & "NULL,"
                    End If
                Case Else
                    dmlValues = dmlValues & "'" & dbfRs(fieldIndex2).Value & "',"
                End Select
        Next fieldIndex2
                
        dmlValues = Left(dmlValues, Len(dmlValues) - 1) & ")"
        
        dmlInsert = "INSERT INTO " & dbfTableName & dmlColumns & " VALUES" & dmlValues
        
        myDb.Execute dmlInsert
        
        dbfRs.MoveNext
    Wend
    
    
    MsgBox "Finished! " & Now
    
End Function
