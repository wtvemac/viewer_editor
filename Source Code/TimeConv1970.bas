Attribute VB_Name = "eMacTimeConv1970"
' ---
' This code is ripped because I'm too lazy to write it myself
' - eMac
' ---

Option Explicit

Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Public Type TIME_ZONE_INFORMATION
  Bias As Long                      'minutes between UTC/GMT and local time
  StandardName(0 To 31) As Integer  'Name of the Standard time zone (each item is a letter)
  StandardDate As SYSTEMTIME        'The current time/date settings for Standard Time
  StandardBias As Long              'The minutes between UTC and standard local time
  DaylightName(0 To 31) As Integer  'Name of the Daylight time zone
  DaylightDate As SYSTEMTIME        'The current time/date settings for Daylight Time
  DaylightBias As Long              'The # of minutes between UTC and Daylight local time
End Type

Public Declare Function GetTimeFormat Lib "kernel32.dll" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As Any, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Public Declare Function GetTimeZoneInformation Lib "kernel32.dll" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Function FormatELTime(EvtSeconds As Long) As String
Dim Seconds As Long, Minutes As Long, Hours As Long, Days As Long
Dim ControlDate As Date, RealDate As Date
Dim RealTime As String
Dim LocaleTimeDiff As Integer

Seconds = EvtSeconds

'We need to get the current time difference between the systems time zone and UTC (or GMT).
'Here we will also calculate Daylight Time, if its applicable.
LocaleTimeDiff = GetTimeDiff()

Minutes = Seconds \ 60
If LocaleTimeDiff > 0 Then              'here we either add or subtract the time difference,
    Minutes = Minutes - LocaleTimeDiff  'as all time zones are either plus (+) or minus (-)
Else                                    'GMT (Greenwich Mean Time).
    Minutes = Minutes + LocaleTimeDiff
End If
Seconds = Seconds Mod 60   ' reset seconds to # of secs left after subtr. of minutes above.

Hours = Minutes \ 60
Minutes = Minutes Mod 60    'reset minutes

Days = Hours \ 24
Hours = Hours Mod 24        'reset hours

ControlDate = #1/1/1970#    'This is our Control Date, or the date from which time starts
                            '(for our uses, time starts then)

'Here we add our Days calculated to the ControlDate variable, which is a Date data-type var.
'set to 1/1/1970.  This gives us our exact date.  The remaining Hours, Minutes and Seconds is
'the exact time of day.
RealDate = ControlDate + Days

Hours = Hours + 4

'Now the exact date and time has been calculated.  The next step is to convert the output for
'the local time zone and locale formatting properties.  There is no conversion necessary for
'the RealDate variable, as it is a Date data-type, which Windows automatically formats
'according to locale properties.
RealTime = FormatTime(CInt(Hours), CInt(Minutes), CInt(Seconds))

'Return the output as "<date> <time>"
FormatELTime = RealDate & " " & RealTime

End Function

Public Function FormatTime(Hours As Integer, Minutes As Integer, Seconds As Integer) As String

Dim sTime As String                 'This will be our formatted time for the locale
Dim TimeAsSysTime As SYSTEMTIME     'This we will pass as data to the function
Dim lStrLen As Long                 'The length of the returned string

'Load the passed time into a SYSTEMTIME data structure, we only need to load Hours, Minutes
'and Seconds because that's all we're calculating
With TimeAsSysTime
    .wHour = CLng(Hours)
    .wMinute = CLng(Minutes)
    .wSecond = CLng(Seconds)
End With

'pad with spaces for a buffer to pass
sTime = Space(255)

'get the time in the format of the current system locale
lStrLen = GetTimeFormat(0, CLng(0), TimeAsSysTime, CLng(0), sTime, Len(sTime))

'trim off any null chars.
sTime = Left(sTime, lStrLen)

'return the output
FormatTime = sTime

End Function

Public Function GetTimeDiff() As Integer
'This function returns the exact number of minutes we need to add (or subtract) to get the
'time correct for the running systems locale.

Dim TZInfo As TIME_ZONE_INFORMATION
Dim lRet As Long                        'return value
Dim iTimeDiff As Integer                'calculated time difference

lRet = GetTimeZoneInformation(TZInfo)
'If lRet <> 1 Then
'    MsgBox "Error Getting this systems Time Zone Information.", vbApplicationModal + vbExclamation + vbOKOnly, "GetTimeZoneInformation()"
'    Exit Function
'End If

iTimeDiff = CInt(TZInfo.Bias + TZInfo.StandardBias + TZInfo.DaylightBias)

GetTimeDiff = iTimeDiff

End Function
