<html>

<!-- #include file="_system.asp" -->

<head>
</head>

<body bgcolor="#FFFFFF">
 
 <%
 ge_ID = request.querystring("ge_ID")
 book_start_dato = request.querystring("book_dato")

 set conn = server.createObject("ADODB.connection")
 conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"


  set rs = conn.execute("select * from genstand where ge_ID = " & ge_ID) 

  if rs.bof AND rs.eof then
 
   rs.close 
   set rs = nothing  
   conn.close
   set conn = nothing 

   response.write "<h2>Genstand ikke fundet ...</h2>"
   response.write "<hr><a href=javaScript:history.go(-1)><img border=0 src='_images_knapper/kalender.gif'> G? tilbage til booking kalender</a>" 
   response.end
  
  end if
 
  ge_ID = rs("ge_ID") 
  ge_navn = rs("ge_navn")
  
 
  response.write "<table><tr><td><img height='50' border='0' hspace='10' src='_images_genstand/" & rs("ge_billede") & "'></td>"
  response.write "<td><h2>Booking - " & ge_navn & "</h2></td></tr></table>"

  response.write "<br><b>Prisliste</b><br><br>F?rste dag kr. " & rs("ge_pris_forste_dag") & ",-<br>Efterf?lgende dage kr. " & rs("ge_pris_efterfolgende_dag") & ",-"
  response.Write "<br>Forsikring p? " & lcase(ge_navn) & " kan v?lges til kr. " & rs("ge_pris_forsikring") & ",- pr. dag med en selvrisiko p? kr. " & rs("ge_selvrisiko") & ",-"
  
  if ge_ID < 1000 then 
   response.write "<br><br>Levering og afhentning kr. " & rs("ge_pris_opstilling") & ",- indtil 50 km fra Herning, derefter k?rsel efter statens takster, kr. " & km_sats / 100 & " pr. km. excl. moms"
  else
   depositum = rs("ge_depositum")
   if depositum > 0 then response.Write "<br><br>Depositum kr. " & depositum & ",-"
  end if 
 
 rs.close 
 set rs = nothing 
 
 if ge_ID < 1000 then
 
  response.Write "<br><br>Hvis du selv afhenter, s? kan du leje en trailer til transporten, hvis der er en ledig i den periode, du v?lger at leje."
  response.Write "<br>Pris p? trailer er. kr. 500,- incl. moms pr. dag. Traileren er en Selandia F-2036 HTD, totalv?gt 2.000 kg. Rampe og pallel?fter/s?kkevogn medf?lger."
  response.Write "<br><br>Du kan se den samlede pris p? n?ste side, f?r du godkender lejen."
 end if
 
 response.Write "<br><br>" 
 
 set rs = conn.execute("select * from reservationer where re_start_dato <= '" & DKdate2USdate(book_start_dato) & "' AND re_slut_dato >= '" & DKdate2USdate(book_start_dato) & "' AND re_genstand_ID = '" & ge_ID & "' order by re_start_dato")
 
  if NOT rs.eof then

   response.write "<h2>Status : RESERVERET</h2>Perioden " & formatdatetime(rs("re_start_dato"), 1) 
   if rs("re_start_dato") <> rs("re_slut_dato") then response.write " - " & formatdatetime(rs("re_slut_dato"), 1)
   response.write "<br><br>" 

   response.write "Du kan f? tilsendt en e-mail, hvis reservationen slettes.<br>Indtast din e-mail adresse herunder og tryk p? SEND.<br>"
   %>
   <br>
   <a title="Tilbage til kalender" href="javaScript:history.go(-1)"><img border="0" src="_images_knapper/kalender.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;
   <a title="Tilbage til oversigt" href="genstand_oversigt.asp"><img border="0" src="_images_knapper/oversigt.gif"></a>
  
   
   <form name="book_email" action="genstand_book_email.asp?ge_ID=<%response.write ge_ID %>&book_dato=<% response.write book_start_dato %>" method="POST" > 
    <input type="text" name="email" size="100">
    <br><br>
    <input type="submit" value="Send" name="B1"></p>
   </form>
   <%
 
 else
  re_status = "L"     
 end if 

 rs.close 
 set rs = nothing  
  
 if re_status = "L" then
  response.write "<h2>Status : LEDIG</h2>" 
   
  set rs = conn.execute("select * from reservationer where re_start_dato > '" & DKdate2USdate(book_start_dato) & "' AND re_genstand_ID = '" & ge_ID & "' order by re_start_dato")
  
   if NOT rs.eof then antal_dage_frem = datediff("d", book_start_dato, rs("re_start_dato")) - 1
   if (antal_dage_frem < 1) OR (antal_dage_frem > 14) then antal_dage_frem = 14 
  
  rs.close 
  set rs = nothing  

  response.write "Angiv herunder den periode, hvor du ?nsker at leje " & ge_navn & ".<br>Bem?rk, at du kun reserverer " & lcase(ge_navn) & ", indtil der er indbetalt mindst 50 % af det samlede lejebel?b, hvorefter reservationen skifter til udlejet.<br>"
  response.Write "<br>Lejeperioden starter p? lejedagen kl. 09:00 og slutter dagen efter 'til og med' dagen kl. 09:00<br>"
  
  %>
  <form name="book_reserver" action="genstand_book_reserver.asp?ge_ID=<%response.write ge_ID %>&book_dato=<% response.write book_start_dato %>" method="POST" > 
   Reserver fra og med <b><% response.write formatdatetime(book_start_dato, 1) %></b> til og med&nbsp;
   <b><select size="1" name="book_slut_dato">
    <option>samme dag</option>
    <%
    for dag_nr = 1 to antal_dage_frem
     naeste_dato = dateadd("d", dag_nr, book_start_dato)
     response.write "<option value='" & naeste_dato & "'>" & formatdatetime(naeste_dato, 1) & "</option>"  
    next
    %>
   </select></b>
   <br><br>

   Angiv navn og adresse herunder.<br>
   <textarea rows="4" name="ku_navn_adresse" cols="80"></textarea><br><br>

   E-mail adresse skal angives herunder, da vi kun sender faktura og oplysninger om betalingen p? e-mail.<br>   
   <input type="text" name="ku_email" size="80">
   
   <br /><br />?nsker du at tegne forsikring p? <% response.write lcase(ge_navn) %>, s? kryds af her
   <input type="checkbox" name="ge_forsikring" value="1">   
   
   <%if ge_ID < 1000 then  %>
     <br /><br />?nsker du <% response.write lcase(ge_navn) %> leveret og afhentet, s? kryds af her
     <input type="checkbox" name="ge_opstilling" value="1">
     og angiv antal km. fra Herning til din adresse <input type="text" name="ge_km" size="10" value="0">
     <br /><font color= red>ELLER</font><br />
     ?nsker du at leje en ledig trailer til transporten, s? kryds af her
     <input type="checkbox" name="ge_trailer" value="1"> og kryds af her <input type="checkbox" name="ge_trailer_forsikring" value="1"> hvis
     du ?nsker forsikring p? trailer 
   <%end if %>
   


   <br><br><input type="submit" value="Reserver" name="B1">
  </form>
  <%
 
 end if 

 conn.close
 set conn = nothing 
 %>
  
</body>
</html>