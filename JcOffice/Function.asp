<%
' ###################################################
'
' 公共函数
'
' DPRS - ASP Digital Press Retrieval System
' Copyright (c) 2009 Windsfly.CN
' Author : 孙峰
' Contact : QQ( 105857855,729028242 ) 
' E-mail : sf0223cn@163.com
' Support : http://www.windsfly.cn
' Author Blog : http://www.windsfly.cn/blog
'
' ###################################################

' *************************************
' 跳转 By sunfeng 2009-3-13
' @param msg 提示信息
' @param uri 链接地址
' @param flag 标识
' *************************************
Function goURL(msg, uri, flag)
	If flag = 0 Then
		Response.Write "<script>window.location.href='"&uri&"'</script>"
	ElseIf flag = 1 Then
		Response.Write "<script>window.alert('"&msg&"');window.location.href='"&uri&"';</script>"
	ElseIf flag = 2 Then
		Response.Write "<script>window.alert('"&msg&"');window.history.back();</script>"
	End If
End Function

' *************************************
' 日期转换函数
' *************************************
Function DateToStr(DateTime, ShowType)
    Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond
    Dim FullWeekday, shortWeekday, Fullmonth, Shortmonth, TimeZone1, TimeZone2
    TimeZone1 = "+0800"
    TimeZone2 = "+08:00"
    FullWeekday = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    shortWeekday = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    Fullmonth = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Shortmonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    DateWeek = Weekday(DateTime)
    DateSecond = Second(DateTime)
    If Len(DateMonth)<2 Then DateMonth = "0"&DateMonth
    If Len(DateDay)<2 Then DateDay = "0"&DateDay
    If Len(DateMinute)<2 Then DateMinute = "0"&DateMinute
    Select Case ShowType
        Case "Y-m-d"
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay
        Case "Y-m-d H:I A"
            Dim DateAMPM
            If DateHour>12 Then
                DateHour = DateHour -12
                DateAMPM = "PM"
            Else
                DateHour = DateHour
                DateAMPM = "AM"
            End If
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
        Case "Y-m-d H:I:S"
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
        Case "YmdHIS"
            DateSecond = Second(DateTime)
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond
        Case "ym"
            DateToStr = Right(Year(DateTime), 2)&DateMonth
        Case "d"
            DateToStr = DateDay
        Case "ymd"
            DateToStr = Right(Year(DateTime), 4)&DateMonth&DateDay
        Case "mdy"
            Dim DayEnd
            Select Case DateDay
            Case 1
                DayEnd = "st"
            Case 2
                DayEnd = "nd"
            Case 3
                DayEnd = "rd"
            Case Else
                DayEnd = "th"
        End Select
        DateToStr = Fullmonth(DateMonth -1)&" "&DateDay&DayEnd&" "&Right(Year(DateTime), 4)
        Case "w,d m y H:I:S"
            DateSecond = Second(DateTime)
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = shortWeekday(DateWeek -1)&","&DateDay&" "& Left(Fullmonth(DateMonth -1), 3) &" "&Right(Year(DateTime), 4)&" "&DateHour&":"&DateMinute&":"&DateSecond&" "&TimeZone1
        Case "y-m-dTH:I:S"
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&TimeZone2
        Case Else
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
    End Select
End Function

' *************************************
' 查询报社名称
' @param officeid 报刊ID
' 2009-3-23 By Sunfeng
' *************************************
Function FindOffice(ByVal officeid)
	'Dim officeid
	Dim SqlOffice,RsOffice,Tb_OfficeName
	If officeid <> "" Then
		SqlOffice = "Select Tb_OfficeName From Tb_Office Where Tb_OfficeID=" & Officeid
		Set RsOffice = Jcs.Execute(SqlOffice)
		Tb_OfficeName = RsOffice("Tb_OfficeName")
		RsOffice.Close:Set RsOffice = Nothing
	End If
	FindOffice = Tb_OfficeName
End Function 
%>