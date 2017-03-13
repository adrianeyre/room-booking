<%

' ****************************************************
' *              confirmbooking.asp                  *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 05/12/2012                 *
' *             Version : 1.0.0                      *
' *                                                  *
' ****************************************************
%>

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 5px;
	margin-right: 0px;
	margin-bottom: 0px;
}
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
}
.style1 {font-size: x-large}
.style2 {color: #FFFFFF}
-->
</style>
<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/PCRoomConnection.asp" -->

<%
Dim AdminDetails
Dim AdminDetails_numRows
Dim Administrators
Dim Admins(20)
Dim Username
Dim TempUsername

TempUsername = Server.HTMLEncode(Request.ServerVariables("AUTH_USER"))
RealUsername = ""

num = 0
for a = 1 to len(TempUsername)
	num = num + 1
	if mid(TempUsername,a,1) = "\" then numon = num + 1
next
for a = numon to len(TempUsername)
	RealUsername = RealUsername + lcase(mid(TempUsername,a,1))
next

Set AdminDetails = Server.CreateObject("ADODB.Recordset")
AdminDetails.ActiveConnection = MM_PCRoomConnection_STRING
AdminDetails.Source = "SELECT * FROM dbo.AdminDetails"
AdminDetails.CursorType = 0
AdminDetails.CursorLocation = 2
AdminDetails.LockType = 1
AdminDetails.Open()

AdminDetails_numRows = 0
Administrators = 0

While (Not AdminDetails.EOF)
	Administrators = Administrators + 1
	Admins(Administrators) = AdminDetails.Fields.Item("Username").Value
    AdminDetails.MoveNext
Wend

Admin = false
for a = 1 to Administrators
	if Admins(a) = RealUsername then Admin = true
next


Dim BookDate 'as string

BookDate =  request.querystring("Date")
'BookDate = Cstr(BookDate)
Room =  request.querystring("Room")
Row =  request.querystring("Row")
Col =  request.querystring("Col")
Username =  request.querystring("Name")
DateOn =  request.querystring("DateOn")
PreBook = request.querystring("Prebook")
showpre = request.querystring("showpre")
SelectDay = request.querystring("SelectDay")
SelectWeek = request.querystring("SelectWeek")
Action =  request.querystring("Action")
Usage =  request.querystring("Usage")

if Admin = false then
	if Username <> RealUsername then Action = "adminerror"
end if

Dim CheckRecord
Dim CheckRecord_numRows

if Action = "delete" then
	Set CheckRecord = Server.CreateObject("ADODB.Recordset")
	CheckRecord.ActiveConnection = MM_PCRoomConnection_STRING
	CheckRecord.Source = "SELECT * FROM dbo.RoomBookings WHERE BookingDate LIKE '" + Replace(BookDate, "'", "''") + "' AND RoomNumber=" + Replace(Room, "'", "''") + " AND Row=" + Replace(Row, "'", "''") + " AND Col= "+ Replace(Col, "'", "''") + " AND Name LIKE '" + Replace(Username, "'", "''") +  "'"
	CheckRecord.CursorType = 0
	CheckRecord.CursorLocation = 2
	CheckRecord.LockType = 1
	CheckRecord.Open()

	CheckRecord_numRows = 0

	While (Not CheckRecord.EOF)
		CheckRecord_numRows = CheckRecord_numRows + 1
		'Temp = CheckRecord.Fields.Item("Username").Value
	    CheckRecord.MoveNext
	Wend
	if CheckRecord_numRows = 0 and Admin = false then Action = "prebook"

	CheckRecord.Close()
	Set CheckRecord = Nothing
end if



if Action = "add" then
	if PreBook = "checkbox" and Admin = true then
		set Command1 = Server.CreateObject("ADODB.Command")
		Command1.ActiveConnection = MM_PCRoomConnection_STRING
		Command1.CommandText = "INSERT INTO dbo.PreBooking (RoomNumber, Row, Col, Name, Usage)  VALUES ('" & Room & "' , '" & Row & "' , '" & Col & "' , '" & Username & "' , '" & Usage & "' ) "
		Command1.CommandType = 1
		Command1.CommandTimeout = 0
		Command1.Prepared = true
		Command1.Execute()
	else
		set Command1 = Server.CreateObject("ADODB.Command")
		Command1.ActiveConnection = MM_PCRoomConnection_STRING
		Command1.CommandText = "INSERT INTO dbo.RoomBookings (BookingDate, RoomNumber, Row, Col, Name, Usage)  VALUES ('" & BookDate & "' , '" & Room & "' , '" & Row & "' , '" & Col & "' , '" & Username & "' , '" & Usage & "' ) "
		Command1.CommandType = 1
		Command1.CommandTimeout = 0
		Command1.Prepared = true
		Command1.Execute()
	end if
end if

if Action = "delete" then
'	if PreBook = "checkbox" and Admin = true then
	if CheckRecord_numRows = 0 and Admin = true then
		set Command1 = Server.CreateObject("ADODB.Command")
		Command1.ActiveConnection = MM_PCRoomConnection_STRING
		Command1.CommandText = "DELETE FROM dbo.PreBooking WHERE RoomNumber=" + Replace(Room, "'", "''") + " AND Row=" + Replace(Row, "'", "''") + " AND Col= "+ Replace(Col, "'", "''")
		Command1.CommandType = 1
		Command1.CommandTimeout = 0
		Command1.Prepared = true
		Command1.Execute()
	else
		set Command1 = Server.CreateObject("ADODB.Command")
		Command1.ActiveConnection = MM_PCRoomConnection_STRING
		Command1.CommandText = "DELETE FROM dbo.RoomBookings WHERE BookingDate LIKE '" + Replace(BookDate, "'", "''") + "' AND RoomNumber=" + Replace(Room, "'", "''") + " AND Row=" + Replace(Row, "'", "''") + " AND Col= "+ Replace(Col, "'", "''") + " AND Name LIKE '" + Replace(Username, "'", "''") +  "'"
		Command1.CommandType = 1
		Command1.CommandTimeout = 0
		Command1.Prepared = true
		Command1.Execute()
	end if
		'response.write(BookDate)
end if

if (Action = "prebook" or Action = "adminerror") and Admin = false then
%>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td background="../images/backcaritas.png" bgcolor="#FF0000"><div align="center" class="style1 style2">Error!</div></td>
  </tr>
  <tr>
    <td height="50"><div align="center">
<%
		select case Action
			case "prebook": message = "This is a pre-booking, please contact an Administrator to delete the booking"
			case "adminerror": message = "Only Administrator can delete other bookings!"
		end select
		response.write(message)
%>
	</div></td>
  </tr>
  <tr>
    <td height="50"><div align="center">
      <form action="booking.asp?dateon=<%response.write(DateOn)%>&RoomName=<%response.write(Room)%>&newname=<%response.write(Username)%>" method="post" name="form1" target="_self" id="form1">
        <label>
        <input type="submit" name="Submit" value="Submit" />
        </label>
                        </form>
      </div></td>
  </tr>
</table>
<%
else
	if SelectDay <> "" then
		Page = "booking.asp?dateon="+DateOn+"&newname="+Username+"&prebook="+PreBook+"&showpre="+showpre+"&SelectWeek="+SelectWeek+"&SelectDay="+SelectDay
	else
		Page = "booking.asp?dateon="+DateOn+"&RoomName="+Room+"&newname="+Username+"&prebook="+PreBook+"&showpre="+showpre
	end if
	'response.write(Page)
	Response.Redirect(Page)
end if
%>
