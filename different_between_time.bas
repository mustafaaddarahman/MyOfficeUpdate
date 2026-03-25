Attribute VB_Name = "different_between_time"
Function TimeDifferenceInMinutesAndSeconds(startTime As String, EndTime As String) As String
    Dim startTimeValue As Date
    Dim endTimeValue As Date
    Dim timeDifference As Double
    Dim minutes As Long
    Dim seconds As Long
    
    ' Õ”«» «·«Êﬁ«  «·‰’ÌÂ «·Ï  «—ÌŒ ÊÊﬁ 
    startTimeValue = TimeValue(startTime)
    endTimeValue = TimeValue(EndTime)
    
    ' Õ”«» «·›—ﬁ »Ì‰ «·Êﬁ Ì‰
    timeDifference = endTimeValue - startTimeValue
    
    ' Õ”«» «·œﬁ«∆ﬁ Ê«·ÀÊ«‰Ì
    minutes = Int(timeDifference * 24 * 60)
    seconds = Round((timeDifference - (minutes / 1440)) * 86400)
    
    ' «—Ã«⁄ «·‰ ÌÃ… ﬂœﬁ«∆ﬁ ÊÀÊ«‰Ì
    TimeDifferenceInMinutesAndSeconds = minutes & ":" & seconds
End Function
