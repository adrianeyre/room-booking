<%
' ****************************************************
' *                    booking.asp                   *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 05/12/2012                 *
' *             Version : 1.0.1                      *
' *                                                  *
' ****************************************************
%>

<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/PCRoomConnection.asp" -->
<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
}
body {
	margin-left: 0px;
	margin-top: 5px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {
	font-size: x-large;
	color: #FFFFFF;
}
#MainLayer {
	position:absolute;
	width:710px;
	height:1960px;
	z-index:1;
	left: 0px;
	top: 366px;
}
.style4 {font-size: x-large}
.style5 {color: #FFFFFF}
a:link {
	color: #FFFF00;
	text-decoration: none;
}
a:visited {
	color: #FFFF00;
	text-decoration: none;
}
a:hover {
	color: #FFFF00;
	text-decoration: none;
}
a:active {
	color: #FFFF00;
	text-decoration: none;
}
.style8 {font-size: 18px; font-weight: bold; color: #FF0000; }
.style9 {font-size: 18px; font-weight: bold; color: #FFFFFF; }
.style10 {color: #FFFFFF; font-weight: bold; }
.style11 {color: #FF0000; font-weight: bold;}
.style13 {color: #FFFFFF; font-weight: bold; font-size: large; }
.style15 {color: #FFFFFF; font-size: 18px; }
.style16 {
	color: #FF0000;
	font-size: large;
}
-->
</style>
<script type="text/vbscript">

Sub myFunction(BD, Room, rows, cols, AddUsername, DateOn, Prebook, showpre, SelectWeek, SelectDay)
	Dim MyURL

	if LocalWeb = 0 then
		Usage=InputBox("Enter class or room usage", "PC Room Booking")
	end if
	if SelectWeek <> 0 then
		MyURL = "confirmbooking.asp?Date="&BD&"&Room="&Room&"&Row="&rows&"&Col="&cols&"&Name="&AddUsername&"&DateOn="&DateOn&"&Prebook="&Prebook&"&showpre="&showpre&"&Action=add&Usage="&Usage&"&SelectWeek="&SelectWeek&"&SelectDay="&SelectDay
	else
		MyURL = "confirmbooking.asp?Date="&BD&"&Room="&Room&"&Row="&rows&"&Col="&cols&"&Name="&AddUsername&"&DateOn="&DateOn&"&Prebook="&Prebook&"&showpre="&showpre&"&Action=add&Usage="&Usage
	end if
	location.href = MyURL
End Sub

</script>
</head>

<body>

<%

response.cookies("link")="index.asp>links.asp>staff\staff.asp>staff\booking.asp"
response.cookies("linktext")="Home>Links>Staff Portal>PC Room Booking"

Dim RoomName(20)

Dim SelectWeek
Dim SelectDay
Dim DateOn
SelectDay = 0
SelectWeek = 0
DateOn = 0
LocalWeb = 0
SelectWeek = Request.Form("week")
if SelectWeek = "" then SelectWeek = request.querystring("SelectWeek")
if SelectWeek <> 0 then DateOn = SelectWeek
if SelectWeek = 0 then SelectWeek = 1
SelectDay = Request.Form("day")
if SelectDay = 0 then SelectDay = request.querystring("SelectDay")
If SelectDay <> 0 then Room = 0
if SelectDay <> 0 and DateOn = 0 then DateOn = 1
if lcase(Request.ServerVariables("SERVER_NAME")) = "www.allsaints.notts.sch.uk" then LocalWeb = 1

Dim AddUsername
AddUsername = Request.Form("newname")
if AddUsername = "" then AddUsername = request.querystring("newname")

Dim PreBook
PreBook = Request.Form("prebook")
if PreBook = "" then PreBook = request.querystring("prebook")

Dim showpre
showpre = Request.Form("showpre")
if showpre = "" then showpre = request.querystring("showpre")

Dim RoomNames
Dim RoomNames_numRows

Set RoomNames = Server.CreateObject("ADODB.Recordset")
RoomNames.ActiveConnection = MM_PCRoomConnection_STRING
RoomNames.Source = "SELECT * FROM dbo.RoomDetails"
RoomNames.CursorType = 0
RoomNames.CursorLocation = 2
RoomNames.LockType = 1
RoomNames.Open()

RoomNames_numRows = 0

Dim Computers(20)

While (Not RoomNames.EOF)
	RoomNames_numRows = RoomNames_numRows + 1
	Computers(RoomNames_numRows) = RoomNames.Fields.Item("Computers").Value
    RoomNames.MoveNext
Wend

Set RoomNames = Server.CreateObject("ADODB.Recordset")
RoomNames.ActiveConnection = MM_PCRoomConnection_STRING
RoomNames.Source = "SELECT * FROM dbo.RoomDetails"
RoomNames.CursorType = 0
RoomNames.CursorLocation = 2
RoomNames.LockType = 1
RoomNames.Open()

RoomNames_numRows = 0

Dim Username
Dim RoomData(20,9,9)
Dim AltText(20,9,9)
Dim Title(9,2)
Dim Room
Dim Administrators
Dim Admins(10)
Dim Admin

Room =  request.querystring("RoomName")
Dim Recordset1__MMColParam
Recordset1__MMColParam = Room

if Room = "" then
	Room = Request.Form("RoomName")
	Recordset1__MMColParam = "1"
	If (Request.Form("RoomName") <> "") Then
	  Recordset1__MMColParam = Request.Form("RoomName")
	End If
end if

Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_PCRoomConnection_STRING
Recordset1.Source = "SELECT RoomName, Computers FROM dbo.RoomDetails WHERE ID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0

TempUsername = Server.HTMLEncode(Request.ServerVariables("AUTH_USER"))
Username = ""

num = 0
for a = 1 to len(TempUsername)
	num = num + 1
	if mid(TempUsername,a,1) = "\" then numon = num + 1
next
for a = numon to len(TempUsername)
	Username = Username + lcase(mid(TempUsername,a,1))
next

if AddUsername = "" then AddUsername = Username

Dim AdminDetails
Dim AdminDetails_numRows

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
	if Admins(a) = Username then Admin = true
next

Title(1,1) = "Before"
Title(1,2) = "School"
Title(2,1) = "Period"
Title(2,2) = "1"
Title(3,1) = "Period"
Title(3,2) = "2"
Title(4,1) = "Tutor"
Title(4,2) = "Time"
Title(5,1) = "Period"
Title(5,2) = "3"
Title(6,1) = "Dinner"
Title(6,2) = "Time"
Title(7,1) = "Period"
Title(7,2) = "4"
Title(8,1) = "Period"
Title(8,2) = "5"
Title(9,1) = "After"
Title(9,2) = "School"

Dim GetDates
Dim GetDates_numRows
Dim AmountofDates
Dim Dates(50)
Dim NextDate
Dim PrevDate

Set GetDates = Server.CreateObject("ADODB.Recordset")
GetDates.ActiveConnection = MM_PCRoomConnection_STRING
GetDates.Source = "SELECT * FROM dbo.Dates"
GetDates.CursorType = 0
GetDates.CursorLocation = 2
GetDates.LockType = 1
GetDates.Open()

GetDates_numRows = 0
AmountofDates = 0

if DateOn = 0 then DateOn = request.querystring("DateOn")

ThisDate = Date - 6

While (Not GetDates.EOF)
	TempDate = GetDates.Fields.Item("Dates").Value
	If DateDiff("d", TempDate, ThisDate) < 1 Then
		AmountofDates = AmountofDates + 1
		Dates(AmountofDates) = TempDate
	end if
    GetDates.MoveNext
Wend

NextDate = DateOn + 1
if NextDate > AmountofDates then NextDate = AmountofDates

PrevDate = DateOn - 1
if PrevDate < 1 then PrevDate = 1

Dim PreBookings
Dim PreBookings_numRows

Set PreBookings = Server.CreateObject("ADODB.Recordset")
PreBookings.ActiveConnection = MM_PCRoomConnection_STRING
PreBookings.Source = "SELECT * FROM dbo.PreBooking"
PreBookings.CursorType = 0
PreBookings.CursorLocation = 2
PreBookings.LockType = 1
PreBookings.Open()

PreBookings_numRows = 0

if Admin = true and showpre = "checkbox" then

else
	While (Not PreBookings.EOF)
		temproom = PreBookings.Fields.Item("RoomNumber").Value
		temprow = PreBookings.Fields.Item("Row").Value
		tempcol = PreBookings.Fields.Item("Col").Value
		tempname = PreBookings.Fields.Item("Name").Value
		tempalt = PreBookings.Fields.Item("Usage").Value
		RoomData(temproom,temprow,tempcol) = tempname
		AltText(temproom,temprow,tempcol) = tempalt
	    PreBookings.MoveNext
	Wend
end if

Dim RoomBookings
Dim RoomBookings_numRows

Set RoomBookings = Server.CreateObject("ADODB.Recordset")
RoomBookings.ActiveConnection = MM_PCRoomConnection_STRING
RoomBookings.Source = "SELECT * FROM dbo.RoomBookings"
RoomBookings.CursorType = 0
RoomBookings.CursorLocation = 2
RoomBookings.LockType = 1
RoomBookings.Open()

RoomBookings_numRows = 0

While (Not RoomBookings.EOF)
	tempdate = RoomBookings.Fields.Item("BookingDate").Value
	temproom = RoomBookings.Fields.Item("RoomNumber").Value
	temprow = RoomBookings.Fields.Item("Row").Value
	tempcol = RoomBookings.Fields.Item("Col").Value
	tempname = RoomBookings.Fields.Item("Name").Value
	tempalt = RoomBookings.Fields.Item("Usage").Value
	if Dates(DateOn) = tempdate then
		RoomData(temproom,temprow,tempcol) = tempname
		AltText(temproom,temprow,tempcol) = tempalt
	end if
    RoomBookings.MoveNext
Wend

%>

<table width="100%" height="45" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="100%" height="40" align="left" background="../images/backdefault.png" bgcolor="#192F68"><div align="center" class="style1 style4">PC Room Booking</div></td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#999999"><form action="booking.asp?DateOn=1" name="form1" method="post">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="20" colspan="4" align="center" valign="top" bgcolor="#003399"><label></label>            <label><span class="style15">Search by Room </span></label></td>
          </tr>
        <tr>
          <td width="145" height="20" align="center" valign="middle">&nbsp;</td>
          <td width="150">&nbsp;</td>
          <td width="150">&nbsp;</td>
          <td width="270">&nbsp;</td>
        </tr>
        <tr>
          <td height="20" align="center" valign="middle">&nbsp;</td>
          <td>Please select room </td>
          <td><select name="RoomName" id="RoomName">
		  	<option value="0">Select</option>
            <%
			AmountofRooms = 0
While (NOT RoomNames.EOF)
			AmountofRooms = AmountofRooms + 1
			RoomName(AmountofRooms) = RoomNames.Fields.Item("RoomName").Value
%>
            <option value="<%=(RoomNames.Fields.Item("ID").Value)%>"><%response.write(RoomName(AmountofRooms))%></option>
            <%
  RoomNames.MoveNext()
Wend
If (RoomNames.CursorType > 0) Then
  RoomNames.MoveFirst
Else
  RoomNames.Requery
End If
%>
          </select></td>
          <td><input type="submit" name="Submit" value="Submit"></td>
        </tr>
      </table>
        </form>    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#666666"><form name="form3" method="post" action="booking.asp">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="4" bgcolor="#003399"><div align="center"><span class="style15">Search by Date </span></div></td>
          </tr>
        <tr>
          <td width="145" height="25">&nbsp;</td>
          <td width="150" height="25">&nbsp;</td>
          <td width="150" height="25"><label></label></td>
          <td width="270" height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Day of Week </td>
          <td height="25"><label>
            <select name="day" id="day">
              <option value="0">Select</option>
              <option value="1">Monday</option>
              <option value="2">Tuesday</option>
              <option value="3">Wednesday</option>
              <option value="4">Thursday</option>
              <option value="5">Friday</option>
            </select>
          </label></td>
          <td height="25"><label></label></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Week Commencing</td>
          <td height="25"><select name="week" id="week">
            <option value="0" selected>Select</option>
            <% for a = 1 to AmountofDates%>
            <option value="<%response.write(a)%>">
            <%response.write(Dates(a))%>
            </option>
            <% next %>
          </select></td>
          <td height="25"><input type="submit" name="Submit3" value="Submit"></td>
        </tr>
      </table>
        </form>    </td>
  </tr>
</table>
<%
if Admin = true then
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
  	<%
		link = "booking.asp?dateon="&DateOn&"&RoomName="&Room
		if SelectDay <> "" and SelectDay <> 0 then link = link & "&SelectWeek="&SelectWeek&"&SelectDay="&SelectDay
	%>
    <td><form name="form2" method="post" action="<%response.write(link)%>">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="3" valign="bottom" bgcolor="#003399"><label><div align="center"><span class="style15">Administrator Pannel</span></div></label></td>
          </tr>
        <tr>
          <td width="149" bgcolor="#999999">&nbsp;</td>
          <td width="236" bgcolor="#999999">&nbsp;</td>
          <td width="330" bgcolor="#999999"><label></label></td>
          </tr>
        <tr>
          <td bgcolor="#999999">&nbsp;</td>
          <td bgcolor="#999999">Username</td>
          <td bgcolor="#999999"><input name="newname" type="text" id="newname" value="<%if AddUsername <> "" then response.write(AddUsername)%>" size="30">
            <input type="submit" name="Submit2" value="Submit"></td>
        </tr>
        <tr>
          <td bgcolor="#999999">&nbsp;</td>
          <td bgcolor="#999999">Pre Booking (Yes/No) </td>
          <td bgcolor="#999999"><input name="prebook" type="checkbox" id="prebook" value="checkbox" <%if PreBook = "checkbox" then response.write("checked") %>></td>
        </tr>
        <tr>
          <td bgcolor="#999999">&nbsp;</td>
          <td bgcolor="#999999">Remove Pre Bookings View (Yes/No) </td>
          <td bgcolor="#999999"><label>
            <input name="showpre" type="checkbox" id="showpre" value="checkbox" <%if showpre = "checkbox" then response.write("checked") %>>
          </label></td>
        </tr>
        <tr>
          <td bgcolor="#999999">&nbsp;</td>
          <td bgcolor="#999999">&nbsp;</td>
          <td bgcolor="#999999">&nbsp;</td>
        </tr>
      </table>
        </form>    </td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<%
end if

if Room <> 0 then %>

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#192F68">
  <tr>
    <td height="30" colspan="3" background="../images/backdefault.png"><div align="center" class="style4"><span class="style10">Room:</span><span class="style11"> <%=(RecordSet1.Fields.Item("RoomName").Value)%></span></div></td>
  </tr>
  <tr>
    <td width="22" height="30" background="../images/backdefault.png">&nbsp;</td>
    <td width="475" height="30" background="../images/backdefault.png"><span class="style9">Week Commencing</span><span class="style8">
      <%response.write(Dates(DateOn))%>
    </span></td>
    <td width="150" height="30" background="../images/backdefault.png"><span class="style9">Computers </span><span class="style8"><%=(RecordSet1.Fields.Item("Computers").Value)%></span></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#999999"><table width="100%" border="2" cellspacing="0" cellpadding="0">
      <tr>
        <td width="58">&nbsp;</td>
        <td width="168"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=1&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self">Monday</a></strong></div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
          </tr>
        </table></td>
        <td width="168"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=2&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self">Tuesday</a></strong></div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
          </tr>
        </table></td>
        <td width="168"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=3&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self">Wednesday</a></strong></div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
          </tr>
        </table></td>
        <td width="168"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=4&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self">Thursday</a></strong></div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
          </tr>
        </table></td>
        <td width="168"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=5&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self">Friday</a></strong></div></td>
          </tr>
          <tr>
            <td><div align="center"></div></td>
          </tr>
        </table></td>
      </tr>
    </table>
<% for rows = 1 to 9 %>
      <table width="100%" border="2" cellspacing="0" cellpadding="0">
        <tr>
          <td width="58"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><div align="center">
                  <%response.write(Title(rows,1))%>
              </div></td>
            </tr>
            <tr>
              <td><div align="center">
                  <%response.write(Title(rows,2))%>
              </div></td>
            </tr>
          </table></td>
<%
	for cols = 1 to 5
%>
        <td width="168"><table width="100%" border="0" cellspacing="0" cellpadding="0" <%
			select case lcase(RoomData(Room,rows,cols))
				Case Username: response.write("bgcolor='#0000FF'")
				Case "closed", "inset", "pshe","intervention","concert","interviews","interview","show","production","event", "revision", "dofe" : response.write("bgcolor='#FF00FF'")
				case "upgrade","exams","exam": response.write("bgcolor='#FF6600'")
				Case ""
				case else: response.write("bgcolor='#FF0000'")
			end select
%>>
          <tr>
            <td><div align="center" class="style5"><%response.write(RoomData(Room,rows,cols))
			if RoomData(Room,rows,cols) <> "" and AltText(Room,rows,cols) <> "" then %>
              <img src="../images/infoicon.jpg" alt="<%response.write(AltText(Room,rows,cols))%>" width="10" height="10" border="0"></div></td>
			<% end if %>
          </tr>
          <tr>
            <td><div align="center">
			<% 	if RoomData(Room,rows,cols) = Username or (Admin = true and  RoomData(Room,rows,cols) <> "") then response.write("<a href='confirmbooking.asp?Date="&Dates(Dateon)&"&Room="&Room&"&Row="&rows&"&Col="&cols&"&Name="&RoomData(Room,rows,cols)&"&DateOn="&DateOn&"&Prebook="&Prebook&"&showpre="&showpre&"&Action=delete' target='_self'>Cancel</a>") else response.write("&nbsp;")
				'if RoomData(Room,rows,cols) = "" then response.write("<a href='confirmbooking.asp?Date="&Dates(Dateon)&"&Room="&Room&"&Row="&rows&"&Col="&cols&"&Name="&AddUsername&"&DateOn="&DateOn&"&Prebook="&Prebook&"&showpre="&showpre&"&Action=add' target='_self'>Available</a>")
				if RoomData(Room,rows,cols) = "" then %>
				<button onClick="myFunction '<%response.write(Dates(Dateon))%>',<%response.write(Room)%>,<%response.write(rows)%>,<%response.write(cols)%>,'<%response.write(AddUsername)%>',<%response.write(DateOn)%>,'<%response.write(Prebook)%>','<%response.write(showpre)%>',0,0">Available</button>
				<%
				end if
				%>
			</div></td>
          </tr>
        </table></td>
<%
	next
%>
        </tr>
      </table>
<%
	next
%>    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40"><div align="center"><a href="booking.asp?dateon=<%response.write(PrevDate)%>&RoomName=<%response.write(Room)%>&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>"><img src="../images/previousbutton.png" width="140" height="29" border="0"></a></div></td>
    <td height="40"><div align="center"><a href="booking.asp?dateon=<%response.write(NextDate)%>&RoomName=<%response.write(Room)%>&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/nextbutton.png" width="140" height="29" border="0"></a></div></td>
  </tr>
</table>
<% else
	if SelectDay <> 0 then
		DisplayDay = cint(mid(Dates(DateOn),1,2))
		DisplayDay = DisplayDay + SelectDay - 1
		'DisplayDate = cstr(DisplayDay) + mid(Dates(DateOn),3,8)
		DisplayDate = DateAdd("d",(SelectDay-1),Dates(DateOn))
%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#192F68">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td background="../images/backdefault.png"><div align="center"><span class="style13">Date: </span><span class="style16">
        <%response.write(DisplayDate)%></span></div></td>
      </tr>
      <tr>
        <td background="../images/backdefault.png"><div align="center"><span class="style13">Day: </span><span class="style16">
        <%
		select case SelectDay
			case 1: message = "Monday"
			case 2: message = "Tuesday"
			case 3: message = "Wednesday"
			case 4: message = "Thursday"
			case 5: message = "Friday"
		end select
		response.write(message)

		%></span> </div></td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#999999"><table width="100%" border="2" cellspacing="0" cellpadding="0">
      <tr>
        <td width="58">&nbsp;</td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Before</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>School</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Period</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>1</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Period</strong></div></td>
          </tr>
          <tr>
            <td height="17"><div align="center"><strong>2</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Tutor</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>Time</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Period</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>3</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Dinner</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>Time</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Period</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>4</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>Period</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>5</strong></div></td>
          </tr>
        </table></td>
        <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><div align="center"><strong>After</strong></div></td>
          </tr>
          <tr>
            <td><div align="center"><strong>School</strong></div></td>
          </tr>
        </table></td>
      </tr>
    </table>
        <%

		for RoomOn = 1 to AmountofRooms %>
        <table width="100%" border="2" cellspacing="0" cellpadding="0">
          <tr>
            <td width="58"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><div align="center"><font size="2"><a href="booking.asp?dateon=<%response.write(DateOn)%>&RoomName=<%response.write(RoomOn)%>&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><%response.write(RoomName(RoomOn))%></a></font>
                  </div>                    <div align="center"></div></td>
                </tr>
                <tr>
                  <td><div align="center"><%response.write(Computers(RoomOn))%></div></td>
                </tr>

            </table></td>
            <%
			Period = 0
	for cols = 1 to 9
			'select case Period
			'	case 0 : Period = 2
			'	case 2 : Period = 3
			'	case 3 : Period = 5
			'	case 5 : Period = 7
			'	case 7 : Period = 8
			'end select
			Period = Period + 1
%>
            <td width="93"><table width="100%" border="0" cellspacing="0" cellpadding="0"
<%
			select case lcase(RoomData(RoomOn,Period,SelectDay))
				Case Username: response.write("bgcolor='#0000FF'")
				Case "closed", "inset", "pshe","intervention","concert","interviews","interview","show","production","event", "revision", "dofe": response.write("bgcolor='#FF00FF'")
				case "upgrade","exams","exam": response.write("bgcolor='#FF6600'")
				Case ""
				case else: response.write("bgcolor='#FF0000'")
			end select
%>>
                <tr>
                  <td><div align="center" class="style5">
					<%response.write(RoomData(RoomOn,Period,SelectDay))
					if RoomData(RoomOn,Period,SelectDay) <> "" and AltText(Room,rows,cols) <> "" then %>
              <img src="../images/infoicon.jpg" alt="<%response.write(AltText(RoomOn,Period,SelectDay))%>" width="10" height="10" border="0">
			<% end if %>
                  </div></td>
                </tr>
                <tr>
                  <td><div align="center">
                      <% 	if RoomData(RoomOn,Period,SelectDay) = Username or (Admin = true and  RoomData(RoomOn,Period,SelectDay) <> "") then response.write("<a href='confirmbooking.asp?Date="&Dates(SelectWeek)&"&Room="&RoomOn&"&Row="&Period&"&Col="&SelectDay&"&Name="&RoomData(RoomOn,Period,SelectDay)&"&DateOn="&SelectWeek&"&Prebook="&Prebook&"&showpre="&showpre&"&SelectWeek="&SelectWeek&"&SelectDay="&SelectDay&"&Action=delete' target='_self'>Cancel</a>") else response.write("&nbsp;")
				'if RoomData(RoomOn,Period,SelectDay) = "" then response.write("<a href='confirmbooking.asp?Date="&Dates(SelectWeek)&"&Room="&RoomOn&"&Row="&Period&"&Col="&SelectDay&"&Name="&AddUsername&"&DateOn="&SelectWeek&"&Prebook="&Prebook&"&showpre="&showpre&"&SelectWeek="&SelectWeek&"&SelectDay="&SelectDay&"&Action=add' target='_self'>Available</a>")
				if RoomData(RoomOn,Period,SelectDay) = "" then %>
				<button onClick="myFunction '<%response.write(Dates(SelectWeek))%>',<%response.write(RoomOn)%>,<%response.write(Period)%>,<%response.write(SelectDay)%>,'<%response.write(AddUsername)%>',<%response.write(SelectWeek)%>,'<%response.write(Prebook)%>','<%response.write(showpre)%>',<%response.write(SelectWeek)%>,<%response.write(SelectDay)%>">Available</button>
				<%
				end if
				%>
                  </div></td>
              </tr>
            </table></td>
            <%
	next
%>
          </tr>
      </table>
      <%
	next
%>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40"><div align="center"><a href="booking.asp?dateon=<%response.write(PrevDate)%>&SelectWeek=<%response.write(PrevDate)%>&SelectDay=<%response.write(SelectDay)%>&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/previousweek.png" alt="Previous Week" border="0"></a></div></td>
    <td height="40"><div align="center"><a href="booking.asp?dateon=<%response.write(NextDate)%>&SelectWeek=<%response.write(NextDate)%>&SelectDay=<%response.write(SelectDay)%>&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/nextweek.png" alt="Next Week" width="140" height="29" border="0"></a></div></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=1&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/monday.png" alt="Monday" width="140" height="29" border="0"></a></td>
    <td><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=2&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/tuesday.png" alt="Tuesday" width="140" height="29" border="0"></a></td>
    <td><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=3&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/wednesday.png" alt="Wednesday" width="140" height="29" border="0"></a></td>
    <td><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=4&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/thursday.png" alt="Thursday" width="140" height="29" border="0"></a></td>
    <td><a href="booking.asp?dateon=<%response.write(DateOn)%>&SelectWeek=<%response.write(DateOn)%>&SelectDay=5&newname=<%response.write(AddUsername)%>&prebook=<%response.write(PreBook)%>&showpre=<%response.write(showpre)%>" target="_self"><img src="../images/friday.png" alt="Friday" width="140" height="29" border="0"></a></td>
  </tr>
</table>
<p>
  <%
	end if
end if %>
</p>
</body>
</html>
<%
RoomNames.Close()
Set RoomNames = Nothing

GetDates.Close()
Set GetDates = Nothing

PreBookings.Close()
Set PreBookings = Nothing

Recordset1.Close()
Set Recordset1 = Nothing

RoomBookings.Close()
Set RoomBookings = Nothing

AdminDetails.Close()
Set AdminDetails = Nothing
%>
