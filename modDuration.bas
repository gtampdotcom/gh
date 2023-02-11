Attribute VB_Name = "modDuration"
Option Explicit

' Duration
' By Quentin Zervaas [zervaas@strangeness.org]
' Use at will! Works as well as (better!) than mIRC's
' $duration function
' since it gives you multiple output options
' It's pretty easy to add more of your own
' depending on how you want it.
'
' Just add new cases, and define the strings accordingly
'(remember to allow spaces too.

Public Function Duration(TotalSeconds As Long, UpFormat As _
    Integer) As String
 
  ' Format = 0, 1, 2
  ' This determines the format of the time to be returned
  ' Type 0: 1d 4h 15m 47s
  ' Type 1: 1 day, 4:15:47
  ' Type 2: 1 day 4hrs 15mins 47secs
  ' Type else: Defaults to type 0
  
  Dim Seconds
  Dim Minutes
  Dim Hours
  Dim Days
  Dim Years
  
  Dim SecondString As String
  Dim MinuteString As String
  Dim HourString As String
  Dim DayString As String
  
  Dim DayFormat As String
  Dim HourFormat As String
  Dim MinuteFormat As String
  Dim SecondFormat As String
  
  Seconds = Int(TotalSeconds Mod 60)
  Minutes = Int(TotalSeconds \ 60 Mod 60)
  Hours = Int(TotalSeconds \ 3600 Mod 24)
  Days = Int(TotalSeconds \ 3600 \ 24)

  Select Case UpFormat
    Case 0
      DayString = "d "
      HourString = "h "
      MinuteString = "m "
      SecondString = "s"
    Case 1
      If Days = 1 Then DayString = " day, " _
      Else: DayString = " days, "
      HourString = ":"
      MinuteString = ":"
      SecondString = vbNullString
    Case 2
      If Days = 1 Then DayString = " day " _
      Else: DayString = " days, "
      If Hours = 1 Then HourString = " hour " _
      Else: HourString = " hours "
      If Minutes = 1 Then MinuteString = " minute " _
      Else: MinuteString = " minutes "
      If Seconds = 1 Then SecondString = " second " _
      Else: SecondString = " seconds"
    Case Else
      DayString = "d "
      HourString = "h "
      MinuteString = "m "
      SecondString = "s"
  End Select
  
  'Only display a unit if it's more than zero
  If Days Then DayFormat = Format(Days, "0") & DayString
  If Hours Then HourFormat = Format(Hours, "0") & HourString
  If Minutes Then MinuteFormat = Format(Minutes, "0") & MinuteString
  If Seconds Then SecondFormat = Format(Seconds, "0") & SecondString
  
  Duration = DayFormat & HourFormat & MinuteFormat & SecondFormat
                 
End Function
