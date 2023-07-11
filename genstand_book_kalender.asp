<html>

<!-- #include file="_system.asp" -->

<head>
</head>

<body bgcolor="#FFFFFF"> 
 <%
 ge_ID = request.querystring("ge_ID")
 start_dato = request.querystring("start_dato")
 book_dato = request.querystring("book_dato")

 set conn = server.createObject("ADODB.connection")
 conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"    

 set rs = conn.execute("select * from genstand where ge_ID = " & ge_ID) 

  if rs.bof AND rs.eof then
 
   rs.close 
   set rs = nothing  
   conn.close
   set conn = nothing 

   response.write "<h2>Genstand ikke fundet ...</h2>"
   response.write "<hr><a href=javaScript:history.go(-1)><img border=0 src='_images_knapper/tilbage.gif'> G? tilbage til oversigten</a>" 
   response.end
  
  end if
 
  response.write "<table><tr><td><img height='50' border='0' hspace='10' src='_images_genstand/" & rs("ge_billede") & "'></td>"
  response.write "<td><h2>Booking - Kalender for " & rs("ge_navn") & "</h2></td></tr></table><br>"
  response.write "<td><h3>Der er pt.et problem med bookingen"  & "</h3></td></tr></table>"
  response.write "<td><h3>Kalenderen vises korrekt, og reservationen bliver registreret (du modtager senere en mail), men processen afsluttes med at vise en fejl"

 rs.close 
 set rs = nothing  

 if start_dato = "" then
  denne_maaned = month(now)
  dette_aar = year(now) 
   start_dato = dette_aar & "-" & denne_maaned & "-01"
  end if
  
 slut_dato = DKdate2USdate(dateadd("d", -1, dateadd("m", 3, start_dato)))
 
 dim maaned(3)
 maaned(1) = month(start_dato)
 maaned(2) = month(dateadd("m", 1, start_dato))
 maaned(3) = month(dateadd("m", 2, start_dato))
  
 dim maaned_tekst(3) 
 maaned_tekst(1) = ucase(monthname(maaned(1)))
 maaned_tekst(2) = ucase(monthname(maaned(2)))
 maaned_tekst(3) = ucase(monthname(maaned(3)))
 
 dim antal_dage_i_maaned(3)
 antal_dage_i_maaned(1) = day(dateadd("d", -1, dateadd("m", 1, start_dato)))
 antal_dage_i_maaned(2) = day(dateadd("d", -1, dateadd("m", 2, start_dato)))
 antal_dage_i_maaned(3) = day(dateadd("d", -1, dateadd("m", 3, start_dato)))
 
 dim aar(3)
 aar(1) = year(start_dato)
 aar(2) = year(dateadd("m", 1, start_dato))
 aar(3) = year(dateadd("m", 2, start_dato))

 forrige_vises = (datediff("q", now, start_dato, now) > 0)
 if forrige_vises then forrige_start_dato = dateadd("m", -3, start_dato)

 naeste_start_dato = dateadd("m", 3, start_dato)

 forrige_URL = "genstand_book_kalender.asp?ge_ID=" & ge_ID & "&start_dato=" & DKdate2USdate(forrige_start_dato)
 naeste_URL = "genstand_book_kalender.asp?ge_ID=" & ge_ID & "&start_dato=" & DKdate2USdate(naeste_start_dato)
 
 dim maaned_status(3,31)
 
 sql_start_dato = start_dato
 sql_slut_dato = DKdate2USdate(dateadd("m", 1, start_dato))
  
 for maaned_nr = 1 to 3
  
  sql = "select * from reservationer where (('" & DKdate2USdate(sql_start_dato) & "' < re_slut_dato)" 
  sql = sql & " OR ('" & DKdate2USdate(sql_start_dato) & "' < re_start_dato AND re_start_dato < '" & DKdate2USdate(sql_slut_dato) & "'))"
  sql = sql & " AND re_genstand_ID = '" & ge_ID & "' order by re_start_dato" 

  set rs = conn.execute(sql)

   for dag_nr = 1 to antal_dage_i_maaned(maaned_nr)
   
    if NOT rs.eof then 

     dato_for_dag = aar(maaned_nr) & "-" & maaned(maaned_nr) & "-" & dag_nr

     re_start_dato = DKdate2USdate(rs("re_start_dato"))
     re_slut_dato = rs("re_slut_dato")
     
     if datediff("d", dato_for_dag, re_start_dato) > 0 then
 
       maaned_status(maaned_nr, dag_nr) = "L"

     else

       maaned_status(maaned_nr, dag_nr) = rs("re_status")
       
       if datediff("d", dato_for_dag, re_slut_dato) = 0 then rs.movenext   
      
     end if 
    
    else
     maaned_status(maaned_nr, dag_nr) = "L"
    
    end if 
   
   next ' dag_nr
 
  rs.close 
  set rs = nothing  

  sql_start_dato = DKdate2USdate(dateadd("m", 1, sql_start_dato))
  sql_slut_dato =  DKdate2USdate(dateadd("m", 1, sql_slut_dato))

 next ' maaned_nr
 
 %>

 <table border="0" id="AutoNumber1" cellpadding="2">
   <tr>
     <td>&nbsp;</td>
     <td>
      <% 
      if forrige_vises then
       response.write "<a title='Forrige periode' href=" & forrige_URL & "><img border='0' src='_images_knapper/forrige.gif'></a>"
      else
       response.write "&nbsp;"
      end if
      %>  
     </td>
     <td>
      <a title="N?ste periode" href="<% response.write naeste_URL %>"><img border="0" src="_images_knapper/naeste.gif"></a>
     </td>
     <td width="20">&nbsp;</td>
     <td><a title="Tilbage til oversigt" href="genstand_oversigt.asp">
     <img border="0" src="_images_knapper/oversigt.gif"></a></td>
     <td width="20">&nbsp;</td>
     <td width="20">
     <img border="0" src="_images_knapper/gron.gif" alt="Dagen er ledig - der kan bookes"></td>
     <td>Ledig</td>
     <td width="20">&nbsp;</td>
     <td width="20">
     <img border="0" src="_images_knapper/gul.gif" alt="Dagen er reserveret - du kan bestille en e-mail, hvis dagen bliver ledig"></td>
     <td>Reserveret</td>
     <td width="20">&nbsp;</td>
     <td width="20">
     <img border="0" src="_images_knapper/rod.gif" alt="Dagen er udlejet"></td>
     <td>Udlejet</td>
   </tr>
   </table>
 
 <br>
 
 <table border="1" cellspacing="0" id="kalender" cellpadding="5">
 
  <% 
  for maaned_nr = 1 to 3 

   response.write "<tr><td colspan='31' bgcolor='#FFFFCC'>" & maaned_tekst(maaned_nr) & " " & aar(maaned_nr)& "&nbsp;</td></tr><tr>"

   for dag_nr = 1 to antal_dage_i_maaned(maaned_nr) 
    
    dato = aar(maaned_nr) & "-" & maaned(maaned_nr) & "-" & dag_nr
    if weekday(dato, 1) = 7 then 
     farve = "#FFFF79"
    elseif weekday(dato, 1) = 1 then
     farve = "#FFCC99"
    else
     farve = "#FFFF99"
    end if  
        
    response.write "<td bgcolor='" & farve & "'>" & dag_nr & "</td>"
   next
   
   response.write "</tr><tr>"
   
   idag = DKdate2USdate(now)
   
   for dag_nr = 1 to antal_dage_i_maaned(maaned_nr)
    book_dato = aar(maaned_nr) & "-" & maaned(maaned_nr) & "-" & dag_nr
    
    if datediff("d", book_dato, idag) < 0 then
    
     select case maaned_status(maaned_nr, dag_nr)
     
      case "L" 
       response.write "<td><a title='Ledig - tryk for booking' href='genstand_book.asp?ge_ID=" & ge_ID & "&book_dato=" & book_dato & "'><img border='0' src='_images_knapper/gron.gif'></a></td>"
    
      case "R"
       response.write "<td><a title='Reserveret - tryk for at bestille en e-mail, hvis den bliver ledig' href='genstand_book.asp?ge_ID=" & ge_ID & "&book_dato=" & book_dato & "'><img border='0' src='_images_knapper/gul.gif'></a></td>"
      
      case "U" 
       response.write "<td><img border='0' src='_images_knapper/rod.gif' alt='Udlejet'></td>"

      case else
       response.write "<td>&nbsp;</td>"
       
     end select
    
    else
 
     response.write "<td><img border='0' src='_images_knapper/uncheck.gif'></td>"      
      
    end if 

   next
     
   response.write "</tr>"

  next 
  %> 

 </table>

 <% 
 conn.close
 set conn = nothing 
 %>
 
</body>
</html>