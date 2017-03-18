# Room Booking
#### Technologies: ASP, VBScript, HTML, CSS, MS SQL Server
### [All Saints' Catholic Academy](http://www.allsaints.notts.sch.uk) - Built on 05/12/2012

## Index
* [Installation](#Install)
* [Usage](#Usage)
* [Screen Shots](#Shots)

## Challenege
A web-based PC Room Booking system coded in ASP/VBScript where uses can book an available room. Standard uses can only cancel their own bookings not other user’s bookings. Users can search for available rooms either by room or by date.

## <a name="Install">Installation</a>
* To clone the repo
```shell
$ git clone https://github.com/adrianeyre/room-booking
$ cd room-booking
```

* Set up a web framework such as MS IIS

* Add an ODBC connection to your SQL Server

* Update the file `Connections/PCRoomConnection.asp' with your connection, username and password
```shell
MM_PCRoomConnection_STRING = "dsn=<ODBC Connection>;uid=<USERNAME>;pwd=<PASSWORD>;"
```

## <a name="Shots">Screen Shots</a>
### Default Screen
[![Screenshot](https://raw.githubusercontent.com/adrianeyre/room-booking/master/images/screenshot1.png)](https://raw.githubusercontent.com/adrianeyre/room-booking/master/images/screenshot1.png "Screen Shot 1")

### Room Booking Screen
[![Screenshot](https://raw.githubusercontent.com/adrianeyre/room-booking/master/images/screenshot2.png)](https://raw.githubusercontent.com/adrianeyre/room-booking/master/images/screenshot2.png "Screen Shot 2")

