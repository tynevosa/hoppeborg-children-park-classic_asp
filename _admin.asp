<html>

<!-- #include file="_system.asp" -->

<head>
<title>Hoppeborg.nu - Administration</title>
</head>

<body>
 <h1>Administration</h1>

 <%
 if NOT PasswordGodkendt(session("gn_password")) then response.redirect "_logout.asp"

 mode = request.querystring("mode")
 re_faktura = request.querystring("re_faktura")
 re_ID = request.querystring("re_ID")
 
 set conn = server.createObject("ADODB.connection")
 conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"    



 if mode = "udlejet" then

 ' -------------------------------------------------------------------------
 '   SKIFT STATUS = udlejet
 ' -------------------------------------------------------------------------

  conn.execute("update reservationer set re_status = 'U' where re_faktura = " & re_faktura)
  response.redirect "_admin.asp"

 ' -------------------------------------------------------------------------
 '   SKIFT STATUS = reserveret
 ' -------------------------------------------------------------------------

 elseif mode = "reserveret" then

  conn.execute("update reservationer set re_status = 'R' where re_faktura = " & re_faktura)
  response.redirect "_admin.asp"

 ' -------------------------------------------------------------------------
 '  SLET reservation
 ' -------------------------------------------------------------------------

 elseif mode = "slet" then

  set rs = conn.execute("select re_venteliste, re_start_dato from reservationer where re_ID = " & re_ID) 

  if NOT rs.eof then 
   array_emails = Split(rs("re_venteliste"), "�") 
   for each str_email in array_emails
    if EmailGyldig(str_email) then test = SendHTMLEmail(str_email, "Besked fra Hoppeborg.nu", "", "Reservationen p� Hoppeborg.nu den " & rs("re_start_dato") & " er slettet")
   next
  end if

  rs.close
  set rs = nothing

  conn.execute "delete from reservationer where re_faktura = " & re_faktura

  response.redirect "_admin.asp"

 ' -------------------------------------------------------------------------
 '  VIS alle
 ' -------------------------------------------------------------------------

 else

  set rs = conn.execute("select * from reservationer where re_start_dato > '" & DKdate2USdate(now) & "' AND re_trailerID < 9999 order by re_genstand_ID, re_faktura") 

  genstand_ID = 0

  response.write "<table border='0' cellspacing='5' cellpadding='5'>"
   do until rs.eof  
   
    if rs("re_genstand_ID") <> genstand_ID then
     genstand_ID = rs("re_genstand_ID")
     set ge_rs = conn.execute("select * from genstand where ge_ID = " & genstand_ID ) 

      if NOT ge_rs.eof then
       genstand_navn = ge_rs("ge_navn")
      else
       genstand_navn = "UKENDT"
      end if    

      response.write "<tr><td colspan='13'></td></tr>"
      response.write "<tr bgcolor='#ff9900'><td colspan='13'><b>" & genstand_navn & "</b></td></tr>"
      response.write "<tr bgcolor='#ffff99'><td>Status</td><td>Slet</td><td>Faktura</td><td>Kundenr</td><td>Kunde</td><td>Status</td><td>Fra dato</td><td>Til dato</td><td>Pris</td><td>Depositum</td><td>Betalingsdato</td><td>Levering</td><td>Trailer</td></tr>"
    
     ge_rs.close
     set ge_rs = nothing
    end if
   
    if rs("re_status") = "R" then
     response.write "<tr><td><a title='Skift til udlejet' href='_admin.asp?mode=udlejet&re_faktura=" & rs("re_faktura") & "'><img border='0' src='_images_knapper/book.gif'></a></td>"    
    else
     response.write "<tr><td><a title='Skift til reserveret' href='_admin.asp?mode=reserveret&re_faktura=" & rs("re_faktura") & "'><img border='0' src='_images_knapper/logout.gif'></a></td>"        
    end if 

    response.write "<td><a title='Slet' href='_admin.asp?mode=slet&re_ID=" & rs("re_ID") & "&re_faktura=" & rs("re_faktura") & "'><img border='0' src='_images_knapper/slet.gif'></a></td>"        

    response.write "<td>" & rs("re_faktura") & "</td>"
    
    set ku_rs = conn.execute("select * from kunder where ku_ID = " & rs("re_kunde_ID")) 

     response.write "<td>" & ku_rs("ku_ID") & "</td>"
     response.write "<td>" & left(ku_rs("ku_navn_adresse"), 50) & " ...</td>"
    
    ku_rs.close
    set ku_rs = nothing 
    
    response.write "<td>" & rs("re_status") & "</td>"
    response.write "<td>" & rs("re_start_dato") & "</td>"
    response.write "<td>" & rs("re_slut_dato") & "</td>"
    response.write "<td>" & rs("re_samlet_pris") & "</td>"  
    response.write "<td>" & rs("re_depositum") & "</td>"      
    response.write "<td>" & rs("re_betalings_dato") & "</td>"
    
    re_opstilling = rs("re_opstilling")
    if re_opstilling = "1" then
     response.write "<td>JA</td>"
    else
      response.write "<td>NEJ</td>"
    end if    
    
    re_trailerID = rs("re_trailerID")
    if re_trailerID > 0 AND re_trailerID < 9999 then
     response.write "<td>JA</td>"
    else
      response.write "<td>NEJ</td>"
    end if    

    rs.movenext
   loop  
  response.write "</table>"

  rs.close
  set rs = nothing
  
 end if


 conn.close
 set conn = nothing 
 %>

</body>
</html>