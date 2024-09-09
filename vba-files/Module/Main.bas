Option Explicit

Sub Main()
    

    
    Dim connection As New ADODB.connection
    
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Filename & _
        ";Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
    Dim query As String
    query = "Select * From [tsbd_data$]"
    
    Dim rs As New ADODB.Recordset
    
    rs.Open query, connection
    
    Const rowCount = rs.RecordCount
    
    connection.Close
    
    For x = 0 To rowCount - 1
        For y = 0 To colCount - 1
        
            TSBD (rowCount)
            
        Next y
    Next x
    
End Sub

Function TSBD(rowCount As Integer)
    
    Dim appIE As InternetExplorerMedium
    
    Set appIE = New InternetExplorerMedium
    
    
    strURL = "file:///C:/Users/admin/OneDrive/M%C3%A1y%20t%C3%ADnh/tsbd/steps_code/step%201/H%E1%BB%87%20th%E1%BB%91ng%20Qu%E1%BA%A3n%20l%C3%BD%20T%C3%A0i%20s%E1%BA%A3n%20b%E1%BA%A3o%20%C4%91%E1%BA%A3m%20v%C3%A0%20H%E1%BA%A1n%20m%E1%BB%A9c%20-%20Time%20left_%2059%20minute(s)%2057%20second(s).html"
    With appIE
        .Navigate strURL
        .Visible = True
    End With
    
    Do While appIE.Busy Or appIE.ReadyState <> 4
        DoEvents
    Loop
    
    Set doc = appIE.Document
    
    Set ul = doc.getElementById("horizonalUL")
    
    
End Function

Sub test2()


    'Dim appIE As InternetExplorerMedium
    Dim appIE As Object
    Dim strURL As String
    Dim doc As HTMLDocument
    
    'Set appIE = New InternetExplorerMedium
    Set appIE = CreateObject("InternetExplorer.Application")
    
    
    strURL = "https://www.w3schools.com/cssref/tryit.php?filename=trycss_sel_hover"
    With appIE
        .Navigate strURL
        .Visible = True
    End With
    
    Do While appIE.Busy Or appIE.ReadyState <> 4
        DoEvents
    Loop
    
    'Do While appIE.ReadyState <> 4
    '    Application.Wait DateAdd("s", 1, Now)
    'Loop
    
    
    Set doc = appIE.Document
    
    doc.getElementsByTagName("a")(3).hover
    doc.getElementsByTagName("a")(4).hover
    

End Sub




' Attribute VB_Name = "Main"
' Sub Main()
'     Dim a
'     Dim b
    
'     a = Range("A2").Value
'     b = Range("B2").Value
    
'     Dim appIE As InternetExplorerMedium
    
'     Set appIE = New InternetExplorerMedium
    
    
'     strURL = "10.53.4.188:9098/cms/"
'     With appIE
'         .Navigate strURL
'         .Visible = True
'     End With
    
'     Do While appIE.Busy Or appIE.ReadyState <> 4
'         DoEvents
'     Loop
    
'     Set doc = appIE.Document
    
'     doc.getElementById("fname").Value = a
'     doc.getElementById("lname").Value = b
    
'     For Each l In doc.getElementsByTagName("input")
'        l.Click
'     Next


'     Set appIE = Nothing
' End Sub
