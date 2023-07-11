<html>

<!-- #include file="_system.asp" -->

<head>
</head>

<body bgcolor="#FFFFFF">
 
 <%
 ge_ID = request.querystring("ge_ID")
 book_start_dato = request.querystring("book_dato")
 email = request.form("email")
 
 if email = "" then
  response.write "<h2>E-mail, hvis reservation slettes</h2>Der er ikke indtastet en e-mail adresse<br>" 
  response.write "<hr><a href=javaScript:history.go(-1)><img border=0 src='_images_knapper/book.gif'> G? tilbage til booking</a>" 
  response.end
 end if
 
 
 set conn = server.createObject("ADODB.connection")
 conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"    

 set rs = conn.execute("select re_ID, re_venteliste from reservationer where re_start_dato <= '" & DKdate2USdate(book_start_dato) & "' AND re_slut_dato >= '" & DKdate2USdate(book_start_dato) & "' AND re_genstand_ID = '" & ge_ID & "'")
  re_ID = rs("re_ID")
  re_venteliste = rs("re_venteliste")
  rs.close 
 set rs = nothing  

 if re_ID <> "" then

   conn.execute("update reservationer set re_venteliste = '" & re_venteliste & email & "?' where re_ID = " & re_ID) 

   response.write "<h2>E-mail, hvis reservation slettes</h2>" 
   response.write "Du f?r tilsendt en e-mail, hvis reservationen slettes for " & formatdatetime(book_start_dato, 1) 
       
 else
 
  response.write "<h2>Genstand ikke fundet ...</h2>"
  response.write "<hr><a href=javaScript:history.go(-1)><img border=0 src='_images_knapper/book.gif'> G? tilbage til booking</a>" 
  response.end
 
 end if 
 
  
 conn.close
 set conn = nothing 
 %>
  
</body>
</html>