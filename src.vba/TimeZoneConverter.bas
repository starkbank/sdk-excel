Option Explicit

Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT& = 2
Dim dteStart As Date, dteFinish As Date
Dim dteStopped As Date, dteElapsed As Date
Dim boolStopPressed As Boolean, boolResetPressed As Boolean


Private Type SYSTEMTIME
    wyear As Integer
    wmonth As Integer
    wdayofweek As Integer
    whour As Integer
    wminute As Integer
    wsecond As Integer
    wmilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    bias As Long
    Standardname(1 To 63) As Byte
    standarddate As SYSTEMTIME
    standardbias As Long
    Daylightname(0 To 63) As Byte
    daylightdate As SYSTEMTIME
    daylightbias As Long
End Type
Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (IpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Public resetMe As Boolean
Public myVal As Variant


Public Function UtcToBrt(utcDate As Date) As String
    Dim tzi As TIME_ZONE_INFORMATION
    Dim brt As Date
    Dim dwbias As Long
    Dim tmp As String
    Select Case GetTimeZoneInformation(tzi)
        Case TIME_ZONE_ID_DAYLIGHT
            dwbias = tzi.bias + tzi.daylightbias
        Case Else
            dwbias = tzi.bias + tzi.standardbias
    End Select
    
    brt = DateAdd("n", -dwbias, utcDate)
    tmp = Format(brt, "dd/mm/yyyy hh:mm:ss")
    
    UtcToBrt = tmp
End Function