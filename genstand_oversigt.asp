<html>

<!-- #include file="_system.asp" -->

<head>
</head>

<body bgcolor="#FFFFFF">

 <h2>Booking - Oversigt</h2>
 Tryk p? billedet for valg ...<br><br> 
 <%
 set conn = server.createObject("ADODB.connection")
  conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"    
  
 
 response.write "<table border='0' cellspacing='0' id='AutoNumber1' cellpadding='10' width='100%'>"

  set rs = conn.execute("select * from genstand order by ge_kategori") 

  do until rs.eof  
   response.write "<tr>"
  
   response.write "<td width='120'><a href='genstand_book_kalender.asp?ge_ID=" & rs("ge_ID") & "'><img width='100' border='0' src='_images_genstand/" & rs("ge_billede") & "'></a></td>"
   response.write "<td><a href='genstand_book_kalender.asp?ge_ID=" & rs("ge_ID") & "'><b>" & rs("ge_navn") & "</b></a></td>"
  
   response.write "</tr>"

   rs.movenext
  loop  

  rs.close 
  set rs = nothing  
  
 response.write "</table>"
 
 conn.close
 set conn = nothing 
 %>

</body>
</html>