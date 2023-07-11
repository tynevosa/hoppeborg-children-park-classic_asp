<html>

<!-- #include file="_system.asp" -->

<head>
</head>

<body bgcolor="#FFFFFF"> 
 <%
 
 ' LEJET GENSTAND
 ge_ID = request.querystring("ge_ID")
 ge_book_start_dato = request.querystring("book_dato")
 
 ge_book_slut_dato = request.form("book_slut_dato")
 if NOT isdate(ge_book_slut_dato) then ge_book_slut_dato = ge_book_start_dato
 
 ge_opstilling = request.form("ge_opstilling")
 ge_forsikring = request.form("ge_forsikring")
 
 ge_km = request.form("ge_km")
 if NOT isnumeric(ge_km) then ge_km = 0
 
 
 ' EVENTUEL TILH?RENDE TRAILER 
 ge_trailer = request.form("ge_trailer")
 ge_trailer_forsikring = request.form("ge_trailer_forsikring")


 ' KUNDENS DATA
 ku_navn_adresse = request.form("ku_navn_adresse")
 ku_email = request.form("ku_email")
 
 
 ' TEST AF OVERF?RT DATA
 emailfejl = (ku_email = "") OR (instr(ku_email, "@") < 3) OR (instr(mid(ku_email, 5, 100), ".") = 0) OR (len(mid(ku_email, 5, 100)) < 3) OR (instr(right(ku_email, 4), ".") = 0)

 if ku_navn_adresse = "" OR emailfejl then
  response.write "<h2>Fejl i indtastningerne</h2>Der skal opgives navn og adresse, samt en e-mail adresse.<br><br>Hvis ikke e-mail adressen er gyldig, annulleres reservationen.<br><br>" 
  if emailfejl then response.write "Der er opgivet en forkert e-mail adresse<br><br>" 
  response.write "<hr><a href=javaScript:history.go(-1)><img border=0 src='_images_knapper/book.gif'> G? tilbage til booking</a>" 
  response.end
 end if

 if NOT GenstandLedig(ge_ID, ge_book_start_dato, ge_book_slut_dato) then
  response.write "<h2>Fejl i reservation</h2><b>Reservationen kan ikke gennemf?res</b><br><br>Det kan skyldes, at der er foretaget en reservation, mens du udfyldte forrige side, "
  response.Write "<br>reservationsperioden overlapper en udlejet periode<br>eller der kan v?re sket en fejl i systemet.<br><br>" 
  response.write "<hr><a href=genstand_book_kalender.asp?ge_ID=" & ge_ID & "><img border=0 src='_images_knapper/kalender.gif'> G? tilbage til kalender</a>"
  response.end 
 end if
 
 
 ' HENTER LEJET GENSTANDS DATA
 set conn = server.createObject("ADODB.connection")
 conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"    

 set rs = conn.execute("select * from genstand where ge_ID = " & ge_ID) 



  if rs.bof AND rs.eof then
 
   rs.close 
   set rs = nothing  
   conn.close
   set conn = nothing 

   response.write "<h2>Genstand ikke fundet ...</h2>"
   response.write "<hr><a href=javaScript:history.go(-1)><img border=0 src='_images_knapper/tilbage.gif'> G? tilbage til booking</a>" 
   response.end
  
  end if
 
  ge_navn = rs("ge_navn")
  ge_billede = rs("ge_billede")
  ge_pris_forste_dag = rs("ge_pris_forste_dag")
  ge_pris_efterfolgende_dag = rs("ge_pris_efterfolgende_dag")
  ge_pris_opstilling = rs("ge_pris_opstilling")
  ge_pris_forsikring = rs("ge_pris_forsikring")
  ge_selvrisiko = rs("ge_selvrisiko")
  ge_depositum = rs("ge_depositum")
  ge_storrelse = rs("ge_storrelse")
  
 rs.close 
 set rs = nothing  
 
 
 ' FIND LEDIG TRAILER, HVIS KUNDEN HAR ?NSKET DET
 trailer_ledig = false
 
 if ge_opstilling <> "1" AND ge_trailer = "1" then
  set rs2 = conn.execute("select * from genstand where ge_ID > 999 AND ge_storrelse >= " & ge_storrelse &  " ORDER BY ge_storrelse") 
 
   do until rs2.eof
    tr_ID = rs2("ge_ID")
     
    if GenstandLedig(tr_ID, ge_book_start_dato, ge_book_slut_dato) then
     trailer_ledig = true   
    
     tr_navn = rs2("ge_navn")
     tr_pris_forste_dag = rs2("ge_pris_forste_dag")
     tr_pris_efterfolgende_dag = rs2("ge_pris_efterfolgende_dag")
     tr_pris_forsikring = rs2("ge_pris_forsikring")
     tr_selvrisiko = rs2("ge_selvrisiko")
     tr_depositum = rs2("ge_depositum")
    
     exit do
    end if
   
    rs2.movenext
   loop
   
   rs2.close 
  set rs2 = nothing  
 end if
 
 ' OPRET KUNDEN MED KUNDENUMMER
 conn.execute("insert into kunder(ku_navn_adresse, ku_email) values ('" & ku_navn_adresse & "','" & ku_email & "')")
 set rs = conn.execute("select ku_ID from kunder order by ku_ID desc limit 1")
  ku_ID = rs("ku_ID")
 rs.close 
 set rs = nothing  
 
 
 ' BEKR?FTIGELSE 
 response.write "<table><tr><td><img height='50' border='0' hspace='10' src='_images_genstand/" & ge_billede & "'></td>"
 response.write "<td><h2>Booking - Bekr?ftigelse af booking for " & ge_navn & "</h2></td></tr></table><br>"

 response.write "Kundenummer : " & ku_ID & "<br>" & CrLf2BR(ku_navn_adresse) & "<br><br>"

 response.write "Bekr?ft reservation af " & lcase(ge_navn) & " i "

 if ge_book_start_dato = ge_book_slut_dato then
  antal_dage = 1
  response.write "en dag<br><br><b>" & formatdatetime(ge_book_start_dato, 1) & "</b><br><br>"
  
 else
  antal_dage = datediff("d", ge_book_start_dato, ge_book_slut_dato) + 1
  response.write antal_dage & " dage<br><br><b>fra og med " & formatdatetime(ge_book_start_dato, 1) & " til og med " &  formatdatetime(ge_book_slut_dato, 1) & "</b><br><br>"
  
 end if 
 
 pris_leje = ge_pris_forste_dag + ((antal_dage - 1) * ge_pris_efterfolgende_dag)

 if ge_forsikring = "1" then pris_forsikring = ge_pris_forsikring * antal_dage
 
 if ge_ID > 999 then pris_depositum = ge_depositum
 
 ' HVIS DET LEJEDE SKAL OPSTILLES
 if ge_opstilling = "1" then 
  pris_opstilling = ge_pris_opstilling
  km = ((ge_km - 50) * 4) * 1.25
  if km < 0 then km = 0
  pris_km = round((km * km_sats) / 100)

 ' HVIS DER ?NSKES EN TRAILER TIL TRANSPORTEN
 elseif ge_trailer = "1" then
  
  pris_trailer = tr_pris_forste_dag + ((antal_dage - 1) * tr_pris_efterfolgende_dag)
  
  if ge_trailer_forsikring = "1" then pris_trailer_forsikring = tr_pris_forsikring * antal_dage
  
  pris_trailer_depositum = tr_depositum  
    
 end if

 samlet_pris = pris_leje + pris_opstilling + pris_km + pris_forsikring + pris_trailer + pris_trailer_forsikring

 response.write "Lejepris for perioden er kr. " & samlet_pris & ",-, hvoraf du skal betale kr. " & round(samlet_pris / 2) & ",- forud senest 5 dage efter din reservation<br>"
 if pris_depositum > 0 then response.write "Ud over lejen skal du betale et depositum p? kr. " & pris_depositum & ",- senest ved afhentning af traileren<br>"
 response.write "<br>Det vil fremg? af fakturaen, som vi sender til e-mail adressen <b>" &  ku_email & "</b>, hvordan du indbetaler bel?bet.<br>"

 if ge_opstilling <> "1" AND ge_trailer = "1" then
  if trailer_ledig then
   response.Write "<br>Du har valgt at leje en trailer til transporten.<br>Der er en trailer ledig i perioden, s? den er reserveret og fremg?r af nedenst?ende oversigt.<br>"
   if pris_trailer_depositum > 0 then response.Write "For traileren skal der indbetales et depositum p? kr. " & pris_trailer_depositum & ",- senest ved afhentningen<br>"
  else
     response.Write "<br>Du har valgt at leje en trailer til transporten.<br>Der er desv?rre <b>ikke</b> en trailer ledig i perioden.<br>"
  end if
 end if

 response.Write "<br><b>Den samlede lejepris best?r af:</b><br><br>"
 
 response.write "<table>"
  response.Write "<tr><td>Leje af " & lcase(ge_navn) & "</td><td align='right'>kr. " & pris_leje & ",-</td></tr>"
  if ge_forsikring = "1" then response.Write "<tr><td>Forsikring af " & lcase(ge_navn) & "</td><td align='right'>kr. " & pris_forsikring & ",-</td><td> med en selvrisiko p? kr. </td><td>" & ge_selvrisiko & ",-</td></tr>"
 
  if ge_opstilling = "1" then 
   response.Write "<tr><td>Levering</td><td align='right'>kr. " & pris_opstilling & ",-<td></tr>"
   response.Write "<tr><td>K?rselsudgift</td><td align='right'>kr. " & pris_km & ",-</td></tr>"
  elseif ge_trailer = "1" AND trailer_ledig then
   response.Write "<tr><td>Leje af trailer (" & tr_navn & ")</td><td align='right'>kr. " & pris_trailer & ",-</td></tr>"
   if ge_trailer_forsikring = "1" then response.Write "<tr><td>Forsikring af trailer</td><td align='right'>kr. " & pris_trailer_forsikring & ",-</td><td> med en selvrisiko p? kr. </td><td>" & tr_selvrisiko & ",-</td></tr>"

  end if 
 response.Write "</table>"
 
 if ge_opstilling = "1" and ge_km > 0 then 
  response.write "<br>Du har valgt at f? " & lcase(ge_navn) & " leveret opstillet og angiver " & ge_km & " km. til din adresse<br>"
  response.Write "Hvis der er uoverensstemmelse mellem det opgivne og den faktisk afstand, betales efter den faktiske afstand ved leveringen.<br>"
 end if 
  
 response.write "<br><b>BEM?RK</b>, at reservationen slettes efter 5 dage, hvis ikke vi har modtaget kr. " & round(samlet_pris / 2) & ",- senest 5 dage efter din reservation.<br><br>"

 ' VIS BEKR?FTIGELSE
 %>
 <form name="book_bekraeftigelse" action="genstand_book_bekraeftigelse.asp" method="POST" > 
  <input type="submit" value="Bekr?ft reservationen" name="B1">
  <%
  response.write "<input type='hidden' name='ku_ID' value='" & ku_ID & "'>" 
    
  response.write "<input type='hidden' name='ge_ID' value='" & ge_ID & "'>"
  response.write "<input type='hidden' name='ge_book_start_dato' value='" & ge_book_start_dato & "'>"
  response.write "<input type='hidden' name='ge_book_slut_dato' value='" & ge_book_slut_dato & "'>"
  response.write "<input type='hidden' name='ge_opstilling' value='" & ge_opstilling & "'>"
  response.write "<input type='hidden' name='ge_forsikring' value='" & ge_forsikring & "'>"
  response.write "<input type='hidden' name='ge_km' value='" & ge_km & "'>"
  response.write "<input type='hidden' name='ge_trailer' value='" & ge_trailer & "'>"
  response.write "<input type='hidden' name='ge_trailer_forsikring' value='" & ge_trailer_forsikring & "'>"
  
  response.write "<input type='hidden' name='trailer_ledig' value='" & trailer_ledig & "'>"
  response.write "<input type='hidden' name='tr_ID' value='" & tr_ID & "'>"
  
  response.write "<input type='hidden' name='pris_leje' value='" & pris_leje & "'>"
  response.write "<input type='hidden' name='pris_forsikring' value='" & pris_forsikring & "'>"
  response.write "<input type='hidden' name='pris_opstilling' value='" & pris_opstilling & "'>"
  response.write "<input type='hidden' name='pris_depositum' value='" & pris_depositum & "'>"
  response.write "<input type='hidden' name='pris_km' value='" & pris_km & "'>"
  response.write "<input type='hidden' name='pris_trailer' value='" & pris_trailer & "'>"
  response.write "<input type='hidden' name='pris_trailer_forsikring' value='" & pris_trailer_forsikring & "'>"
  response.write "<input type='hidden' name='pris_trailer_depositum' value='" & pris_trailer_depositum & "'>"
 
  response.write "<input type='hidden' name='samlet_pris' value='" & samlet_pris & "'>"


  %>
 </form>

 <%
 conn.close
 set conn = nothing 
 %>
  
</body>
</html>