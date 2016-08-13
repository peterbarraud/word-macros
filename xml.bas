Attribute VB_Name = "Module1"
Sub HandleXml()
'
' next_line Macro
'
'
    Const FILEPATH = ""
    
    Dim xXML As MSXML2.DOMDocument60
    Dim xNodeList As IXMLDOMNodeList
    Dim xElement As IXMLDOMElement
    Set xXML = New MSXML2.DOMDocument60
    ' ignore the DTD definition. If any
    xXML.setProperty "ProhibitDTD", False
    xXML.resolveExternals = True
    xXML.validateOnParse = True
    ' make sure the Xml is valid
    If xXML.Load(FILEPATH) Then
        Set xNodeList = xXML.SelectNodes("<some path>")
        For Each xElement In xNodeList
            Debug.Print xElement.getAttribute("<attribute name>")
        Next
        ' and if you need
        xXML.Save "<save path>"
    Else
        Debug.Print xXML.parseError.reason
    End If

    
End Sub

