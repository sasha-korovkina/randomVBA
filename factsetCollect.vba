Sub TransferDataToSQLWithNulls()
    Dim strSQL As String
    Dim i As Long
    Dim lastRow As Long

    Dim conn As New ADODB.Connection
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    conn.Open "Provider=SQLOLEDB; Data Source=CMDS-SQL02.CMDS.local; Initial Catalog=SND; Integrated Security=SSPI;"


    ' Loop through the rows in Excel
    For i = 14 To 50
        Dim rankValue As Variant
        Dim holderVal As String
        Dim instCode As String
        Dim fundCode As String
        Dim entityId As String
        Dim vlookupDB As String

        rankValue = ThisWorkbook.Sheets("Public").Cells(i, "A").Value
        holderVal = ThisWorkbook.Sheets("Public").Cells(i, "D").Value
        instCode = ThisWorkbook.Sheets("Public").Cells(i, "E").Value
        fundCode = ThisWorkbook.Sheets("Public").Cells(i, "F").Value
        entityId = ThisWorkbook.Sheets("Public").Cells(i, "G").Value
        vlookupDB = ThisWorkbook.Sheets("Public").Cells(i, "H").Value
        
        If rankValue = "" Then
            strSQL = "INSERT INTO dbo.publicEquities (Rank, HolderID, FundCode, InstCode, EntityID) VALUES (NULL, '" & holderVal & "', '" & fundCode & "', '" & instCode & "', '" & entityId & "')"
        Else
            strSQL = "INSERT INTO dbo.publicEquities (Rank, HolderID, FundCode, InstCode, EntityID) VALUES (" & rankValue & ", '" & holderVal & "', '" & fundCode & "', '" & instCode & "', '" & entityId & "')"
        End If
    
        conn.Execute strSQL
    Next i
'
'        ' Build the SQL INSERT statement
'        strSQL = "INSERT INTO dbo.publicEquities ([Rank]) VALUES ("
'
'
''        ' Add the VlookupDB value
'        If IsEmpty(vlookupValue) Then
'            strSQL = strSQL & "NULL)"
'        Else
'            strSQL = strSQL & "'" & vlookupValue & "')"
'        End If
    
        ' Execute the SQL statement
        'conn.Execute strSQL
    
        ' Check for errors
'        If Err.Number <> 0 Then
'            MsgBox "Error on line " & i & vbCrLf & "Error Description: " & Err.Description
'            Err.Clear
'        End If
    
    On Error GoTo 0



    ' Close the connection
    conn.Close

    ' Clean up
    Set conn = Nothing
End Sub

