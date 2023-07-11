<html>

 <!-- #include file="_system.asp" -->

<head>
</head>

<body bgcolor="#FFFFFF">
 <%
 ku_ID = request.form("ku_ID")
 
 ge_ID = request.form("ge_ID")
 ge_book_start_dato = request.form("ge_book_start_dato")
 ge_book_slut_dato = request.form("ge_book_slut_dato")
 ge_opstilling = request.form("ge_opstilling")
 ge_forsikring = request.form("ge_forsikring")
 ge_km = request.form("ge_km")
 ge_trailer = request.form("ge_trailer")
 ge_trailer_forsikring = request.form("ge_trailer_forsikring")

 trailer_ledig = request.form("trailer_ledig")
 tr_ID = request.form("tr_ID")
 
 pris_leje = request.form("pris_leje") 
 pris_forsikring = request.form("pris_forsikring") 
 pris_opstilling = request.form("pris_opstilling") 
 pris_km = request.form("pris_km") 
 pris_trailer = request.form("pris_trailer") 
 pris_trailer_forsikring = request.form("pris_trailer_forsikring") 
 pris_depositum = request.form("pris_depositum") 
 pris_trailer_depositum = request.form("pris_trailer_depositum") 
 
 samlet_depositum = pris_depositum + pris_trailer_depositum
 
 samlet_pris = request.form("samlet_pris")

 ' Console.WriteLine("test")

 if NOT GenstandLedig(ge_ID, ge_book_start_dato, ge_book_slut_dato) OR (ge_trailer = "1" AND trailer_ledig AND NOT GenstandLedig(tr_ID, ge_book_start_dato, ge_book_slut_dato)) then
  response.write "<h2>Fejl i reservation</h2><b>Reservationen kan ikke gennemf�res</b><br><br>Det kan skyldes, at der er foretaget en anden reservation, inden du fik trykket p� bekr�ftigelse<br>eller der kan v�re sket en fejl i systemet.<br>" 
  response.write "<hr><a href=genstand_book_kalender.asp?ge_ID=" & ge_ID & "><img border=0 src='_images_knapper/kalender.gif'> G� tilbage til kalender</a>" 
  response.end 
 end if

 set conn = server.createObject("ADODB.connection")
 conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"

 ' HENT KUNDEDATA
 set rs = conn.execute("select * from kunder where ku_ID = " & ku_ID) 
  ku_navn_adresse = rs("ku_navn_adresse")
  ku_email = rs("ku_email")
 rs.close 
 set rs = nothing
 
 
 ' HENT GENSTANDDATA 
 set rs = conn.execute("select * from genstand where ge_ID = " & ge_ID) 
  ge_navn = rs("ge_navn")
  ge_billede = rs("ge_billede")
  ge_pris_forste_dag = rs("ge_pris_forste_dag")
  ge_pris_efterfolgende_dag = rs("ge_pris_efterfolgende_dag")
  ge_pris_opstilling = rs("ge_pris_opstilling")
  ge_pris_forsikring = rs("ge_pris_forsikring")
  ge_selvrisiko = rs("ge_selvrisiko")
  ge_storrelse = rs("ge_storrelse")  
  ge_depositum = rs("ge_depositum")
 rs.close 
 set rs = nothing  

  ' HENT TRAILERDATA HVIS VALGT OG LEDIG
 if  ge_trailer = "1" AND trailer_ledig then
  set rs = conn.execute("select * from genstand where ge_ID = " & tr_ID)
   tr_navn = rs("ge_navn")
   tr_billede = rs("ge_billede")
   tr_pris_forste_dag = rs("ge_pris_forste_dag")
   tr_pris_efterfolgende_dag = rs("ge_pris_efterfolgende_dag")
   tr_pris_opstilling = rs("ge_pris_opstilling")
   tr_pris_forsikring = rs("ge_pris_forsikring")
   tr_selvrisiko = rs("ge_selvrisiko")
   tr_storrelse = rs("ge_storrelse")
   tr_depositum = rs("ge_depositum")
  rs.close 
  set rs = nothing  
 else
  tr_ID = 0 
 end if 
  
 antal_dage = datediff("d", ge_book_start_dato, ge_book_slut_dato) + 1
  
 betalings_dato = dateadd("d", 5, DKdate2USdate(now))

 fakturanummer = HentNaesteFakturanummer
 
 ' RESERVER GENSTAND 
 sql =       "insert into reservationer (re_genstand_ID, re_kunde_ID, re_start_dato, re_slut_dato, re_status, re_samlet_pris, re_opstilling, re_forsikring, re_trailerID, re_faktura, "
 sql = sql & "re_betalings_dato, re_depositum) values ('" & ge_ID & "','" & ku_ID & "','" & DKdate2USdate(ge_book_start_dato) & "','" & DKdate2USdate(ge_book_slut_dato) & "','R','" & samlet_pris 
 ''' bypass error
 if samlet_depositum = "" then samlet_depositum = 0
 sql = sql & "','" & ge_opstilling & "','" & ge_forsikring & "','" & tr_ID & "','" & fakturanummer & "','" & DKdate2USdate(betalings_dato) & "','" & samlet_depositum & "')"
 conn.execute(sql)

 ' RESERVER TRAILER HVIS VALGT OG LEDIG 
 if  ge_trailer = "1" AND trailer_ledig then
  sql =       "insert into reservationer(re_genstand_ID, re_kunde_ID, re_start_dato, re_slut_dato, re_status, re_samlet_pris, re_opstilling, re_forsikring, re_trailerID, re_faktura, "
  sql = sql & "re_betalings_dato) values ('" & tr_ID & "','" & ku_ID & "','" & DKdate2USdate(ge_book_start_dato) & "','" & DKdate2USdate(ge_book_slut_dato) & "','R','" & samlet_pris 
  sql = sql & "','0','" & ge_trailer_forsikring & "','9999','" & fakturanummer & "','" & DKdate2USdate(betalings_dato)  & "')"
  conn.execute(sql)
end if

 faktura = HentFaktura
 
 faktura = replace(faktura, "%fakturanr%", fakturanummer)
 faktura = replace(faktura, "%kundenr%", ku_ID)
 fnr = fakturanummer
 navn_adresse = CrLf2BR(ku_navn_adresse)
 if ku_email <> "" then navn_adresse = navn_adresse & "<br>" & ku_email
 faktura = replace(faktura, "%navn_adresse%", navn_adresse)
 

 ' Teksten til faktura
 leje_tekst = "<table style='font-size: 10pt; font-family: Verdana'><tr><td colspan='2'>Leje af " & lcase(ge_navn) & "</td><td align='right'>kr. " & pris_leje & ",-</td></tr>"
 
 leje_tekst = leje_tekst & "<tr><td colspan='3' style='font-size: 8pt; font-family: Verdana'><b>"
 if ge_book_start_dato = ge_book_slut_dato then
   leje_tekst = leje_tekst & formatdatetime(ge_book_start_dato, 1) & " kl. 09.00 - " & formatdatetime(dateadd("d", 1, ge_book_start_dato), 1) & " kl. 09.00</td></tr>"
 else
 leje_tekst = leje_tekst &  formatdatetime(ge_book_start_dato, 1) & " kl. 09.00 - " & formatdatetime(dateadd("d", 1, ge_book_slut_dato), 1) & " kl. 09.00</td></tr>"
 end if
 leje_tekst = leje_tekst & "</b></td></tr><tr><td colspan='3'>&nbsp;</td></tr>"
  
 if ge_opstilling = "1" then 
  leje_tekst = leje_tekst & "<tr><td>Levering og afhentning</td><td></td><td align='right'>kr. " & pris_opstilling & ".-</td></tr>"
  leje_tekst = leje_tekst & "<tr><td>K�rsel, " & ge_km & " km. til leveringsadressen</td><td></td><td align='right'>kr. " & pris_km & ",-</td></tr>" 
 end if 

 if ge_forsikring = "1" then 
  leje_tekst = leje_tekst & "<tr><td>Forsikring af " & lcase(ge_navn) & "</td><td></td><td align='right'>kr. " & pris_forsikring & ",-</td></tr>"
  leje_tekst = leje_tekst & "<tr><td>Selvrisiko</td><td align='right'>kr. " & ge_selvrisiko & ",-</td><td></td></tr>"
 end if 
 
 if ge_opstilling <> "1" AND ge_trailer = "1" AND trailer_ledig then
  leje_tekst = leje_Tekst & "<tr><td>Trailer til transport</td><td></td><td align='right'>kr. " & pris_trailer & ",-</td></tr>"
  
  if ge_trailer_forsikring = "1" then 
   leje_tekst = leje_tekst & "<tr><td>Forsikring af trailer</td><td></td><td align='right'>kr. " & pris_trailer_forsikring & ",-</td></tr>"
   leje_tekst = leje_tekst & "<tr><td>Selvrisiko p� trailer</td><td align='right'>kr. " & tr_selvrisiko & ",-</td><td></td></tr>"
  end if 
 
 end if
  
 leje_tekst = leje_tekst & "</table>"
   
 faktura = replace(faktura, "%leje_tekst%", leje_tekst )

 faktura = replace(faktura, "%samlet_uden_moms%", formatcurrency(samlet_pris * 0.8, 2))
 faktura = replace(faktura, "%moms%", formatcurrency(samlet_pris * 0.2))
 faktura = replace(faktura, "%samlet_med_moms%", formatcurrency(samlet_pris))

 genkendnr = right("00000000000000" & fnr,14)

	cif1 = mid(genkendnr,1,1)
	cif2 = mid(genkendnr,2,1)
	cif3 = mid(genkendnr,3,1)
	cif4 = mid(genkendnr,4,1)
	cif5 = mid(genkendnr,5,1)
	cif6 = mid(genkendnr,6,1)
	cif7 = mid(genkendnr,7,1)
	cif8 = mid(genkendnr,8,1)
	cif9 = mid(genkendnr,9,1)
	cif10 = mid(genkendnr,10,1)
	cif11 = mid(genkendnr,11,1)
	cif12 = mid(genkendnr,12,1)
	cif13 = mid(genkendnr,13,1)
	cif14 = mid(genkendnr,14,1)
	gcif14 = clng(cif14) * 2
	gcif13 = clng(cif13) * 1
	gcif12 = clng(cif12) * 2
	gcif11 = clng(cif11) * 1
	gcif10 = clng(cif10) * 2
	gcif9 = clng(cif9) * 1
	gcif8 = clng(cif8) * 2
	gcif7 = clng(cif7) * 1
	gcif6 = clng(cif6) * 2
	gcif5 = clng(cif5) * 1
	gcif4 = clng(cif4) * 2
	gcif3 = clng(cif3) * 1
	gcif2 = clng(cif2) * 2
	gcif1 = clng(cif1) * 1

	
	IF clng(gcif14) >= 10 THEN
		gcif14 = clng(mid(gcif14,1,1)) + clng(mid(gcif14,2,1))
	END IF
	IF clng(gcif13) >= 10 THEN
		gcif13 = clng(mid(gcif13,1,1)) + clng(mid(gcif13,2,1))
	END IF
	IF clng(gcif12) >= 10 THEN
		gcif12 = clng(mid(gcif12,1,1)) + clng(mid(gcif12,2,1))
	END IF
	IF clng(gcif11) >= 10 THEN
		gcif11 = clng(mid(gcif11,1,1)) + clng(mid(gcif11,2,1))
	END IF
	IF clng(gcif10) >= 10 THEN
		gcif10 = clng(mid(gcif10,1,1)) + clng(mid(gcif10,2,1))
	END IF
	IF clng(gcif9) >= 10 THEN
		gcif9 = clng(mid(gcif9,1,1)) + clng(mid(gcif9,2,1))
	END IF
	IF clng(gcif8) >= 10 THEN
		gcif8 = clng(mid(gcif8,1,1)) + clng(mid(gcif8,2,1))
	END IF
	IF clng(gcif7) >= 10 THEN
		gcif7 = clng(mid(gcif7,1,1)) + clng(mid(gcif7,2,1))
	END IF
	IF clng(gcif6) >= 10 THEN
		gcif6 = clng(mid(gcif6,1,1)) + clng(mid(gcif6,2,1))
	END IF
	IF clng(gcif5) >= 10 THEN
		gcif5 = clng(mid(gcif5,1,1)) + clng(mid(gcif5,2,1))
	END IF
	IF clng(gcif4) >= 10 THEN
		gcif4 = clng(mid(gcif4,1,1)) + clng(mid(gcif4,2,1))
	END IF
	IF clng(gcif3) >= 10 THEN
		gcif3 = clng(mid(gcif3,1,1)) + clng(mid(gcif3,2,1))
	END IF
	IF clng(gcif2) >= 10 THEN
		gcif2 = clng(mid(gcif2,1,1)) + clng(mid(gcif2,2,1))
	END IF
	IF clng(gcif1) >= 10 THEN
		gcif1 = clng(mid(gcif1,1,1)) + clng(mid(gcif1,2,1))
	END IF
		
	sammen = clng(gcif1) + clng(gcif2) + clng(gcif3) + clng(gcif4) + clng(gcif5) + clng(gcif6) + clng(gcif7) + clng(gcif8) + clng(gcif9) + clng(gcif10) + clng(gcif11) + clng(gcif12) + clng(gcif13) + clng(gcif14)
		
	div = clng(sammen) / 10
	
	div = Replace(div,",",".")
	
	IF instr(div,".") THEN
		sdiv = split(div,".")
		div = sdiv(1)
		IF clng(div) = 0 THEN
			checkcif = 0
		ELSE
			checkcif = 10 - clng(div)
		END IF
	ELSE
		checkcif = 0
	END IF
	
	genkendnr = genkendnr & checkcif
		
    faktura = replace(faktura,"%Additional",genkendnr)

 halvdelen = samlet_pris / 2

 faktura = replace(faktura,"%50%",formatcurrency(halvdelen))
 faktura = replace(faktura,"%Rest%",formatcurrency(halvdelen))
 
 b = ""

 if (ge_ID < 1000 AND ge_trailer = "1" AND trailer_ledig) OR (ge_ID > 999) then
  totalpris = int(samlet_pris) + int(samlet_depositum)
 
   b = "Du skal endvidere indbetale et depositum p� kr. " & samlet_depositum & ",- for traileren senest ved afhentningen<br><br><br>"
   b = b & "<b>Afhentning ved selvbetjening</b><br><br>Hvis du �nsker at g�re brug af vores selvbetjening, skal den lejekontrakt du senere modtager returneres i underskrevet stand (gerne scannet ind og sendt pr. e-mail), tillige med kopi af k�rekort.<br><br>"
   b = b & "Betaling kan enten ske, ved brug af nedent�ende betalingsoplysninger (husk at l�gge depositum til 'Total med moms'), eller med Dankort via. Ewire ved at klikke p� dette link :<br><br>"
   b = b & "<br><br>"
   b = b & "<b>Afhentning i �bningstiden med betjening</b><br><br>Hvis du �nsker at benytte dig af vores betjening, underskrives lejekontrakten ved afhentningen, hvor ogs� restbel�bet og depositum skal betales (Kontant eller check)<br><br>"
   b = b & "Betaling kan enten ske, ved brug af nedent�ende betalingsoplysninger (husk at l�gge depositum til 'Total med moms'), eller med Dankort via. Ewire ved at klikke p� dette link :<br><br>"
   b = b & "<br>" 

 else
   b = "Betaling kan enten ske, ved brug af nedent�ende betalingsoplysninger, eller kan betale med <B>MobilePay</B>. Send betalingen til nr. 67 946, og angiv dit faktura- eller kundenummer.<br><br>" 
   b = b & "" 

  end if
 
 faktura = replace(faktura,"%betalingstekst%", b)
  
 test = SendHTMLEmail(ku_email, "Faktura fra Hoppeborg.nu", "", faktura)
 ' test = SendHTMLEmail("mail@hoppeborg.nu", "Faktura p� leje", "", faktura)

 response.write "<table><tr><td><img height='50' border='0' hspace='10' src='_images_genstand/" & ge_billede & "'></td>"
 response.write "<td><h2>Booking - Bekr�ftigelse af booking for " & ge_navn & "</h2></td></tr></table><br>"
 response.write "Der er afsendt en bekr�ftigelse p� lejen til e-mail adressen <b>" & ku_email & "</b><br><br>"
 response.write "Tak for din reservation."

 conn.close
 set conn = nothing 
 %>
    
  
</body>
</html>