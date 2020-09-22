Attribute VB_Name = "modXML"
'Entire code by Didier Aeschimann
'Copyright Didier Aeschimann
'Please do not use without permission of author



Public Function ParseXML(XML As String, XMLStartKey As String) As String
  
  Dim XML2 As String
  Dim StartPos As Long
  Dim EndPos As Long
  Dim XMLEndKey As String
  
  XMLEndKey = UCase("</" & XMLStartKey & ">")
  
  XMLStartKey = UCase("<" & XMLStartKey & ">")
  XML2 = UCase(XML)
  
  StartPos = InStr(1, XML2, XMLStartKey)
  
  If StartPos = 0 Then
      ParseXML = "NONE"
    Exit Function
  End If
  
  EndPos = InStr(StartPos, XML2, XMLEndKey)
  
  If StartPos = 0 Or EndPos = 0 Then
      ParseXML = "NONE"
    Exit Function
  Else
      StartPos = StartPos + Len(XMLStartKey)
  End If
  
  ParseXML = Trim(Mid(XML, StartPos, EndPos - StartPos))

End Function
