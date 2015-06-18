DEFAULT_DEPTH_LEVEL = 5
DEBUG_PATH= "C:\safs\Project\tfsmdebuglog.txt"
DEBUG_FILE_OBJECT = "Nothing"
DDG_TC_REGULAR_STR_PREFIX = "Sys."

Function DesignTimeExecute
  UserForms.TFSM.defaultLevel.Caption = "1. Default depthLevel=" & DEFAULT_DEPTH_LEVEL
  UserForms.TFSM.cxLabelDebug.Caption = "2. " & DEBUG_PATH
  UserForms.TFSM.ShowModal
End Function

Sub TFSM_btnHighLight_OnClick(Sender)
    Dim Form
    Dim WinRs, CompRS
    Dim wObj, cObj
    Dim methodName
         
    methodName = "TFSM_btnHighLight_OnClick() "
    
    Set Form = UserForms.TFSM
       
    winRS = Form.winRS.text
    compRS = Form.compRS.text
  
    Form.btnHighLight.Enabled = False
    StartTime = Timer()
    If InStr(winRS,DDG_TC_REGULAR_STR_PREFIX) <> 1 Then
      Set wObj = DDGGetFindWindow(winRS)
    Else 
      Set wObj = Eval(winRS)
    End If
    StopTime = Timer() 
    
    If TypeName(wObj) <> "Nothing" Then
     Form.winStatus.Style.Font.Color = &hFF0000  
     Form.winStatus.Caption = FormatNumber(StopTime - StartTime,2)
     Debug methodName & "windows object : " & wObj.FullName
     'wObj.click
    Else 
     Form.winStatus.Style.Font.Color = &h0000FF
     Form.winStatus.Caption = "Object not found"
     Form.btnHighLight.Enabled = True
     UserForms.TFSM.SetFocus
     Exit Sub
    End If
         
    StartTime = Timer()    
    If InStr(compRS,DDG_TC_REGULAR_STR_PREFIX) <> 1 Then
      Set cObj = DDGGetFindComponent(wObj,compRS)
    Else
      Set cObj = Eval(compRS)
    End If
    StopTime = Timer()
        
    If compRS <> "" and TypeName(cObj) <> "Nothing" Then
      Form.CompStatus.Style.Font.Color = &hFF0000
      Form.CompStatus.Caption = FormatNumber(StopTime - StartTime,2)
      Form.btnHighLight.Enabled = True
      Debug methodName & "component object : " & cObj.FullName
      Sys.HighlightObject cObj, 12, GetRGBColor(245, 97, 0)
      UserForms.TFSM.SetFocus     
    Else 
      Form.CompStatus.Style.Font.Color = &h0000FF
      Form.CompStatus.Caption = "Object not found"
      Form.btnHighLight.Enabled = True
      UserForms.TFSM.SetFocus
      Exit Sub
    End If    
End Sub

Sub TFSM_cxDebug_OnClick(Sender)
  Dim fs,f
  Dim objFile
  
  Set fs = CreateObject("Scripting.FileSystemObject")  
  If UserForms.TFSM.cxDebug.Checked = True Then
    If Not fs.FileExists(DEBUG_PATH) Then      
      Set objFile = fs.CreateTextFile(DEBUG_PATH)
      objFile.Close      
    End If
    Set f = fs.OpenTextFile(DEBUG_PATH,2, 0)
    Set DEBUG_FILE_OBJECT = f
  Else 
    DEBUG_FILE_OBJECT = "Nothing"
  End If

End Sub

Sub TFSM_ObjectPickerWin_OnObjectPicked(Sender)
  Dim WinRs: Set WinRs = Eval(UserForms.TFSM.ObjectPickerWin.PickedObjectName)
  UserForms.TFSM.winRS.Text = WinRs.FullName
End Sub

Sub TFSM_ObjectPickerComp_OnObjectPicked(Sender)
  Dim CompRs: Set CompRs = Eval(UserForms.TFSM.ObjectPickerComp.PickedObjectName)
  UserForms.TFSM.compRS.Text = CompRs.FullName
End Sub

Sub TFSM_RectObjectPickerWin_OnObjectPicked(Sender)
  Dim WinRs: Set WinRs = Eval(UserForms.TFSM.RectObjectPickerWin.PickedObjectName)
  UserForms.TFSM.WinRS.Text = WinRs.FullName
End Sub

Sub TFSM_RectObjectPickerComp_OnObjectPicked(Sender)
  Dim CompRs: Set CompRs = Eval(UserForms.TFSM.RectObjectPickerComp.PickedObjectName)
  UserForms.TFSM.compRS.Text = CompRs.FullName
End Sub

Function DDGGetFindWindow (windowRecStr)  
  Dim windowPropNames()
  Dim windowPropValues()
  Dim root, processName, processValue, counter, length, methodName,i,j
  Dim newWindowPropNames()
  Dim newWindowPropValues()
  Dim recStrArray
  Dim depthLevelStr, depthLevel
  Dim dbugStr
  Dim win
  
  methodName = "DDGUIUtilites.DDGGetFindWindow() "  
  Set win = Nothing
  Set DDGGetFindWindow = Nothing
  processName = "processName" 
  depthLevel = DEFAULT_DEPTH_LEVEL 
  depthLevelStr = "depthLevel"  ' set level to speed up find process 

  On Error Resume next  
   Debug "Calling method " & methodName
   recStrArray = Split(windowRecStr,";\;",-1,1)  

  For j = 0 To UBound(recStrArray)
    
    DDGConvertStringToProArray recStrArray(j), windowPropNames,windowPropValues
     
    length = Ubound(windowPropNames)
    
    If (InStr(Ucase(recStrArray(j)),Ucase(processName)) > 0 ) Then      
        length = length - 1
    End If 
            
    If (InStr(Ucase(recStrArray(j)),Ucase(depthLevelStr)) > 0 ) Then      
        length = length - 1
    End If
        
    ReDim newWindowPropNames(length) 
    ReDim newWindowPropValues(length) 
    counter =0

    For i = 0 To UBound(windowPropNames)
      
       If (StrComp(Ucase(windowPropNames(i)),Ucase(processName),1) = 0) or _
           (StrComp(windowPropNames(i),depthLevelStr,1) = 0) Then
           
          If (StrComp(Ucase(windowPropNames(i)),Ucase(processName),1) = 0) Then
              processValue = windowPropValues(i) 
              Debug methodName & "window recognition string: " _ 
                  & processName& "=" &processValue 
                         
              Set root = Sys.Find(processName,processValue)                    
              If root.Exists <> True Then
                Debug methodName & "root object not found" 
                Exit Function
              Else
                 Debug methodName & "root object found" 
              End If
          End If
                                    
          If (StrComp(windowPropNames(i),depthLevelStr,1) = 0) Then
              depthLevel = windowPropValues(i)              
          End If
       Else 
            newWindowPropNames(counter) = windowPropNames(i) 
            newWindowPropValues(counter) = windowPropValues(i)               
            counter = counter + 1
            dbugStr = dbugStr + windowPropNames(i) & "="  & windowPropValues(i) & ";"            
       End If       
    Next    
    
    If Len(Join(newWindowPropNames)) <> 0  Then
                  
      Set win = root.FindChild(newWindowPropNames, newWindowPropValues, depthLevel, True)
      If win.Exists = False Then
        Debug methodName & "Windows object not found"  
        Exit Function
      End If
      
      dbugStr = dbugStr + "depthLevel="  & depthLevel
      Debug  methodName & "window recognition string: "& dbugStr
      dbugStr = "" 
      Debug methodName & "Windows object found"             
      set root = win
            
    End If        
  Next

    Set DDGGetFindWindow = root
  
End Function

Function DDGGetFindComponent (windowObject, cmpRecStr)
  Dim methodName, depthLevel, counter, i, j, depthLevelStr, length
  Dim cmpPropNames()
  Dim cmpPropValues()
  Dim newCmpPropNames()
  Dim newCmpPropValues()
  Dim recStrArray
  Dim dbugStr
  Dim comp
  
  Set comp = Nothing
  Set DDGGetFindComponent = Nothing
  depthLevel = DEFAULT_DEPTH_LEVEL
  depthLevelStr = "depthLevel"  ' set level to speed up find process
  methodName = "DDGUIUtilites.DDGGetFindComponent() "  
 
  Debug "Calling method " & methodName

  On Error Resume next
  recStrArray = Split(cmpRecStr,";\;",-1,1)  

  For j = 0 To UBound(recStrArray)
  
    DDGConvertStringToProArray recStrArray(j), cmpPropNames,cmpPropValues
  
    If (InStr(Ucase(recStrArray(j)),Ucase(depthLevelStr)) > 0 ) Then
      length = Ubound(cmpPropNames)- 1
     Else 
      length = Ubound(cmpPropNames)
    End If

    ReDim newCmpPropNames(length) 
    ReDim newCmpPropValues(length)    
    counter =0

    For i = 0 To UBound(cmpPropNames)
                    
        If (StrComp(cmpPropNames(i),depthLevelStr,1) = 0) Then
            depthLevel = cmpPropValues(i)                
        Else 
            newCmpPropNames(counter) = cmpPropNames(i) 
            newCmpPropValues(counter) = cmpPropValues(i)                           
            counter = counter + 1       
            dbugStr = dbugStr + cmpPropNames(i) & "="  & cmpPropValues(i) & ";"  
        End If         
    Next  
  
    dbugStr = dbugStr + "depthLevel=" & depthLevel
    Debug methodName & "component recognition string: " & dbugStr 
    dbugStr = ""
                       
    Set comp = windowObject.FindChild(newCmpPropNames, newCmpPropValues, depthLevel, True)   
    If comp.Exists = False Then
      Debug methodName & "Component object not found" 
      Exit Function
    Else
      Debug methodName & "Component object found"
    End If
 
    Set windowObject = comp       
  Next 
  
    Set DDGGetFindComponent = windowObject
  
End Function

Sub DDGConvertStringToProArray (windowRecStr,propNames(),propValues())
    Dim propArray
    Dim propNameValueArray
    Dim cleanStr, length,i
    Dim DDG_TC_FIND_SEARCH_MODE
    DDG_TC_FIND_SEARCH_MODE="TFSM"
  
    If InStr(windowRecStr,DDG_TC_FIND_SEARCH_MODE) = 1 Then
        cleanStr = Replace(windowRecStr,DDG_TC_FIND_SEARCH_MODE,"",1,1)
    Else 
        cleanStr = windowRecStr      
    End If
    propArray = Split(cleanStr,";",-1,1)
    length = Ubound(propArray)
    ReDim propNames(length)
    ReDim propValues(length)
  
    For i = 0 to Ubound(propArray)
      'split the string by '='
      propNameValueArray = Split(propArray(i),"=",-1,1)
      propNames(i) = propNameValueArray(0)
      propValues(i) = propNameValueArray(1)    
    Next

End Sub

Function GetRGBColor(r, g, b)
  GetRGBColor = r + g * 256 + b * 256 * 256
End Function

Sub Debug (dbugText)  
  If UserForms.TFSM.cxDebug.Checked = True Then
    DEBUG_FILE_OBJECT.Writeline(FormatDateTime(Now())& " : " & dbugText & vbCrlf) 
  End If
End Sub

