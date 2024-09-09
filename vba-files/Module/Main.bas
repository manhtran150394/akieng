Attribute VB_Name = "Main"
Sub Main()
    Dim a
    Dim b
    
    a = Range("A2").Value
    b = Range("B2").Value
    
    Dim appIE As InternetExplorerMedium
    
    Set appIE = New InternetExplorerMedium
    
    
    strURL = "10.53.4.188:9098/cms/"
    With appIE
        .Navigate strURL
        .Visible = True
    End With
    
    Do While appIE.Busy Or appIE.ReadyState <> 4
        DoEvents
    Loop
    
    Set doc = appIE.Document
    
    doc.getElementById("fname").Value = a
    doc.getElementById("lname").Value = b
    
    For Each l In doc.getElementsByTagName("input")
       l.Click
    Next

    Set appIE = Nothing
End Sub
