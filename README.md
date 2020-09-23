<div align="center">

## Display database results in 'pages'


</div>

### Description

Many a time I found myself needing to display results in 'page' form. And call me old fashioned, devoted to learning or just plain stoopid, but I'd much rather learn how to code something myself than just plug in variables and let a snotty frontpage wizard do it for me.

So, here it is, simple code to divide records into 'pages' of your choosing.
 
### More Info
 
I'm assuming that you have a database that will be used as an input. This will be discussed in the code.

There are a number of assumptions (basically page, table and database names), but they're noted in the code and are easy to customize.

A pretty paged list of entries :)

Drowsiness (or at least I was after taking the time to code it..)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tim Feeley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tim-feeley.md)
**Level**          |Intermediate
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__4-8.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tim-feeley-display-database-results-in-pages__4-6696/archive/master.zip)

### API Declarations

Free. Take it. Don't credit me. See if I care. Hrmpf.


### Source Code

```
<%
  'Initialize the data environment using ADO.
  set DBView = Server.CreateObject("ADODB.Recordset")
  'DSN smells. I use a connection on a folder in my webpage.
  'simply replace the connection string with your own path/DSN.
  DBView.ActiveConnection = "Driver={Microsoft Access Driver (*.mdb)};DBQ=c:/inetpub/wwwroot/fpdb/blarp.mdb"
  ' MUCHOS IMPORTANT! The main function (RecordSet) will not work
  ' if you omit this statement. Resist the urge to delete it!
  DBView.cursortype = 1
  ' Do not delete the above line or your recordset will always return -1.
  ' This is the SQL, replace it with your own
  ' (from is in brackets because i was a d0pe and made it my
  ' field name, so I needed to distinguish that i KNEW my syntax,
  ' and that it's only a field name.)
  DBView.Source = "SELECT [from], subject, date, message FROM messages ORDER BY Date DESC"
  DBView.Open
  ' When the user navigates using the page list, it makes the querystring page=1,2,etc.
  ' This will determine the page
  pgs = request.querystring("page")
  ' If it's blank, set it to one.
  if pgs = "" then pgs = 1 : rpg = 1
  ' RPG is basically a silly variable i came up with that
  ' is discontinued in this edition of my code. But you can keep it.
  rpg = pgs
  ' This is used to determine how far to move in the recordset.
  pgs = ((pgs - 1) * 5) +1
  ' We're only going to view 5 at a time (which is why 5 is there)
  ' You can change that number to whatever you want, just make sure to
  ' do it in the points ahead.
  ' Also, we're keeping a 'Curre' counter to determine how many records
  ' we've displayed. After it's reached 5 (or whatever you want), we're stopping.
  Curre = 0
  ' Okay, let's loop (if we're not at the end of the file.)
  if not DBView.EOF Then
  ' We're moving it to the record of the page we're starting from.
  DBView.Move (pgs - 1)
  ' Change this 5 to whatever you want. Remember 'Curre' is just
  ' how far we looped THIS time.
  While curre < 5 and not DBView.EOF
	'Basically, display the repeated results. I included my table here, but feel free to modify it with your own fields.
 %>
  <tr>
  <td width="100%" bgcolor="#800000"><b><font face="Arial" size="3"><%=(DBView("Subject"))%></font></b></td>
  </tr>
  <tr>
  <td width="100%" class="tbody" align="right" bgcolor="#800000">By: <%=(DBView("from"))%><br>Date: <%=(DBView("Date"))%></td>
  </tr>
  <tr>
  <td width="100%" class="tbody"><%=(DBView("Message"))%><br></td>
  </tr>
  <tr><td width="100%" height="20"></td></tr>
 <% DBView.MoveNext
  curre = curre +1
  wend
  ' Keep looping and moving records until we reach 5 records per page,
  ' or the end of the file. This End If ends the checking if we're at the End of the file.
 	end if
 %>
  < finish your HTML repeating reigon here, for example, close your table >
  <td width="100%" bgcolor="#800000" class="tbody">Messages: <b><%=pgs%>-<%=pgs+(curre-1)%> </b>of <b><%=(DBView.RecordCount)%></b></td></tr>
  <tr><td width="100%" class="tbody">
   <%
   ' Okay, a few things. The HTML line with ASP 'embedded' displays what we're looking at. We start at the pgs number (the beginning record)
   ' for example, 10 or 15. We're then adding how many records we could show to that number to get the end of the range. If all 5 cound be
   ' shown, we'll add 5, if not, we'll add just what we displayed.
			pqs = request.querystring("page")
			if pqs = "" then pqs = 1
   ' What page are we on again? The above code checks it. The code below this checks how many pages we'll need to fit all the data.
   ' Change all the 5's to whatever number you want.
			pages = int(dbview.recordcount \ 5)
			if dbview.recordcount mod 5 <> 0 then pages = pages + 1
   ' This code takes the numeric portion of the records divided by five, for example, if we had 7 records this would be 1.
   ' We then see if there's any left over data by using Mod and if so, add another page to accomodate it.
   ' Start displaying it.
			response.write("Pages: [")
   ' Loop through all the page numbers.
			For AI = 1 to pages
   ' I used a bunch of IF's when debugging. If it bothers you, add an Else :P
   ' We just check to see if pqs (the variable used above to determine what page is being viewed) matches the loop
   ' if so, don't add a hyperlink.
				if Cint(AI) = cint(pqs) then response.write "&nbsp;&nbsp;&nbsp;<B>" & ai & "</B>&nbsp;&nbsp;&nbsp;"
				if cint(AI) <> cint(pqs) then response.write "&nbsp;&nbsp;&nbsp;<a href=""blarp.asp?page=" & ai & """>" & ai & "</a>&nbsp;&nbsp;&nbsp;"
			Next
			response.write ("]")
   ' That's it. Note how in the link code, the page is blarp.asp. Change this with your own or you'll have a sad broken link.
%>
```

