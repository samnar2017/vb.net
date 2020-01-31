Attribute VB_Name = "Module1"
Option Explicit

'   Google Fill Travel Times/Distance with Vias


Const strUnits = "imperial" ' imperial/metric (miles/km)
Const strTransportMode = "driving" ' alternative = 'walking'
Const strDelimeter = "|"    'for lists of via points
Const MAX_GOOGLE_RETRIES = 10




Function GetGoogleTravelTimeByRange(ByRef rngSource As Range, ByRef strTravelTime As String, ByRef strDistance As String, ByRef strError As String)
'Returns the total journey times for all cells in the range rngSource

Dim strList As String
Dim strWaypoints As String
Dim lngColumnCount As Long ' Total Columns in Range
Dim lngRowCount As Long ' Total Columns in Range
Dim lngRow As Long ' Current Row
Dim lngCol As Long ' Current Col
Dim strCellValue As String

 
'Convert the range into a List
With rngSource
    lngColumnCount = .Columns.Count 'Number of columns in the selected range
    lngRowCount = .Rows.Count 'Number of rows in the selected range
     
    For lngRow = 1 To lngRowCount   'for each row in the selected range
        For lngCol = 1 To lngColumnCount 'for each column in the selected range
            strCellValue = .Cells(lngRow, lngCol)
            
            If strCellValue <> "" And Len(strCellValue) > 0 Then  ' if the cell is not empty
                 strList = strList & IIf(Len(strList) > 0, strDelimeter, "") & strCellValue
            End If
        Next
    Next
End With

GetGoogleTravelTimeByRange = GetGoogleTravelTimeByList(strList, strTravelTime, strDistance, strError)
 
End Function


Function GetGoogleTravelTimeByList(ByVal strList As String, ByRef strTravelTime As String, ByRef strDistance As String, ByRef strError As String)
'Returns the travel times for a list of addresses (seperated by strDelimeter constant defined above, |)
Dim arrList
Dim strFrom As String
Dim strTo As String
Dim strPrefix As String
Dim i As Long


arrList = Split(strList, strDelimeter)

Select Case UBound(arrList)
Case Is <= 0:
    'Empty List or only 1 item
    strTravelTime = "00:00"
    strDistance = "00:00"
    strError = ""
    Exit Function
    
Case 1:
    'Simple From/To
    strFrom = arrList(0)
    strTo = arrList(1)
    
Case Is > 1:
    'Create waypoints
    strFrom = arrList(0)
    
    strPrefix = "&waypoints="
    For i = 1 To UBound(arrList) - 1
        strFrom = strFrom & strPrefix & arrList(i)
        strPrefix = "|"
    Next
    strTo = arrList(UBound(arrList))
End Select

GetGoogleTravelTimeByList = gglDirectionsResponse(strFrom, strTo, strTravelTime, strDistance, strError)

End Function




Function FillTravelTimes()
' Example function
' This function looks for postcodes/addresses in Columns A to D, and returns total distances/travel times in columns E & F

' This function works by calling GetGoogleTravelTimeByRange with the range A:D for each row,
'  in turn, GetGoogleTravelTimeByRange creates a list of addresses which is passed to GetGoogleTravelTimeByList
'   GetGoogleTravelTimeByList which seperates them list items into waypoints which are then submitted to gglDirectionsResponse
'    gglDirectionsResponse adds up the legs of the journey and returns the total distance/travel time back to this FillTravelTimeFunction.


Dim lngLastRow As Long
Dim lngCurrRow As Long
Dim strFrom As String
Dim strTo As String
Dim strDistance As String
Dim strTravelTime As String
Dim blnOverLimit As Boolean
Dim lngStartTimer As Long
Dim lngQueryCount As Long
Dim lngQueryPauses As Long
Dim strInstructions As String
Dim strError As String
Dim lngRetries As Long


lngStartTimer = Timer
lngQueryCount = 0
lngRetries = 0

Application.DisplayStatusBar = True

With ActiveSheet
    lngLastRow = .UsedRange.Rows.Count  'gets the last row of the sheet that is used
    
    For lngCurrRow = 2 To lngLastRow    'This loops through the rows, starting at Row 2, until the last row
        'Try to work out the TravelTime / Distance
        
        If (CStr(.Range("E" & lngCurrRow)) = "") And (CStr(.Range("F" & lngCurrRow)) = "") Then
        
            Do  ' The Do/Loop will spot OVER_QUERY_LIMIT problems and will keep trying until a good result is found
            
                blnOverLimit = False
                strFrom = "A" & lngCurrRow  'From in Column A
                strTo = "D" & lngCurrRow    'To in Column D
                
                If Not GetGoogleTravelTimeByRange(.Range(strFrom & ":" & strTo), strTravelTime, strDistance, strError) Then
                    strDistance = strError
                    strTravelTime = strError
                    lngRetries = 0
                End If
                                
                If (strDistance = "OVER_QUERY_LIMIT") Or (strTravelTime = "OVER_QUERY_LIMIT") Then
                    ' Google has maxed out, wait a couple of seconds and try again.
                    Application.StatusBar = "Waiting 2 second for Google overload"
                    Application.Wait Now + TimeValue("00:00:02")  ' pause 2 seconds
                    Application.StatusBar = "Try again"
                    
                    lngQueryPauses = lngQueryPauses + 1
                    blnOverLimit = True
                    lngRetries = lngRetries + 1
                Else
                    If (strError = "") And (Val(strDistance) > 0) Then
                        Application.StatusBar = "Processed " & lngCurrRow & "/" & lngLastRow
                        lngQueryCount = lngQueryCount + 1
                    End If
                End If
                
                If lngRetries > MAX_GOOGLE_RETRIES Then
                    ' the Google per day allowance hase been reached
                    GoTo GoogleTooManyQueries
                End If
                
            Loop Until Not blnOverLimit  ' Over Limit either means too many queries too fast, or that the per day allowance has been reached
                    
            ' If the results are ok then populate columns E and F with Distance/Time respectively
            If (strDistance <> "INVALID_REQUEST") And (strTravelTime <> "INVALID_REQUEST") Then
                .Range("E" & lngCurrRow) = strDistance
                .Range("F" & lngCurrRow) = strTravelTime
            End If
            
        End If
    Next

End With

CleanExit:
    Application.StatusBar = "Finished"
    MsgBox "Finished: " & lngQueryCount & " records in processed in " & Round(Timer - lngStartTimer) & " seconds (Total Pauses:" & lngQueryPauses & ")"
    Exit Function

GoogleTooManyQueries:
    MsgBox "Sorry, Google limit of 2000 queries per day has been reached. This may take upto 24 hours to reset", vbCritical
    
    Exit Function

ErrorHandler:
    MsgBox "Error :" & Err.Description, vbCritical
    Exit Function

End Function




Function gglDirectionsResponse(ByVal strStartLocation, ByVal strEndLocation, ByRef strTravelTime, ByRef strDistance, Optional ByRef strError = "") As Boolean
On Error GoTo ErrorHandler
' Helper function to request and process XML generated by Google Maps.
 
Dim strURL As String
Dim objXMLHttp As Object
Dim objDOMDocument As Object
Dim nodeRoute As Object
Dim lngDistance As Long
Dim strThisLegDuration As String
Dim legRoute
Dim lngSeconds As Long


Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
Set objDOMDocument = CreateObject("MSXML2.DOMDocument.6.0")
  
strStartLocation = Replace(strStartLocation, " ", "+")
strEndLocation = Replace(strEndLocation, " ", "+")
strTravelTime = "00:00"
  
strURL = "https://maps.googleapis.com/maps/api/directions/xml" & _
            "?origin=" & strStartLocation & _
            "&destination=" & strEndLocation & _
            "&key=entergooglemapsAPIkey here" & _
            "&sensor=false" & _
            "&units=" & strUnits & _
            "&nocache=" & Now() 'Sensor field is required by google and indicates whether a Geo-sensor is being used by the device making the request
  
'Send XML request
With objXMLHttp
    .Open "GET", strURL, False
    .setRequestHeader "Content-Type", "application/x-www-form-URLEncoded"
    .Send
    objDOMDocument.LoadXML .ResponseText
End With

With objDOMDocument
    If .SelectSingleNode("//status").Text = "OK" Then
        'Get Distance
        
        'Iterate through each leg
        
        For Each legRoute In .SelectSingleNode("//route").ChildNodes
            If legRoute.BaseName = "leg" Then 'SelectSingleNode("/distance/value").Text
                  For Each nodeRoute In legRoute.ChildNodes
                    If nodeRoute.BaseName = "step" Then
                       lngDistance = lngDistance + nodeRoute.SelectSingleNode("distance/value").Text    ' Retrieves distance in meters
                       lngSeconds = lngSeconds + Val(nodeRoute.SelectSingleNode("duration/value").Text)
                    End If
                  Next
            End If
        Next
        
        strTravelTime = formatGoogleTime(lngSeconds)    ' Retrieves distance in meters
        
        
        Select Case strUnits
            Case "imperial": strDistance = Round(lngDistance * 0.00062137, 1)  'Convert meters to miles
            Case "metric": strDistance = Round(lngDistance / 1000, 1) 'Convert meters to miles
        End Select
              
    Else
        strError = .SelectSingleNode("//status").Text
        GoTo ErrorHandler
    End If
End With
  
gglDirectionsResponse = True
GoTo CleanExit
  
ErrorHandler:
    If strError = "" Then strError = Err.Description
    strDistance = -1
    strTravelTime = "00:00"
    gglDirectionsResponse = False
  
CleanExit:
    Set objDOMDocument = Nothing
    Set objXMLHttp = Nothing
  
End Function




Public Function formatGoogleTime(ByVal lngSeconds As Double)
'Helper function. Google returns the time in seconds, so this converts it into time format hh:mm
 
Dim lngMinutes As Long
Dim lngHours As Long
 
lngMinutes = Fix(lngSeconds / 60)
lngHours = Fix(lngMinutes / 60)
lngMinutes = lngMinutes - (lngHours * 60)
 
formatGoogleTime = Format(lngHours, "00") & ":" & Format(lngMinutes, "00")

End Function
