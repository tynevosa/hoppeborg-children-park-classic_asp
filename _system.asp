<head>
<link href="default.css" rel="stylesheet">
<meta http-equiv="imagetoolbar" content="no">

<script language="javascript" src="sniffer.js"></script>
<script language="javascript1.2" src="custom.js"></script>
<script language="javascript1.2" src="style.js"></script>
</head>

<body>
<script language=JavaScript>


var message="Dette er ikke muligt";

function clickIE4(){
 if (event.button==2){
  alert(message);
  return false;
 }
}

function clickNS4(e){
 if (document.layers||document.getElementById&&!document.all){
  if (e.which==2||e.which==3){
   alert(message);
   return false;
  }
 }
}

if (document.layers){
 document.captureEvents(Event.MOUSEDOWN);
 document.onmousedown=clickNS4;
}
else if (document.all&&!document.getElementById){
 document.onmousedown=clickIE4;
}

document.oncontextmenu=new Function("alert(message);return false")


</script>
</body>

<% 
 
session.LCID = 1030

' ------------------------------------------------------------
'   STAMDATA
' ------------------------------------------------------------

 km_sats = 400

' ------------------------------------------------------------
'   PASSWORD
' ------------------------------------------------------------

 login_password = "gnhoppeborg"   

' ------------------------------------------------------------

 function PasswordGodkendt(PG_password)
  if PG_password = login_password then
   PasswordGodkendt = true
  else
   PasswordGodkendt = false
  end if
 end function   

' ------------------------------------------------------------
'   DATABASE hoppeborg
' ------------------------------------------------------------

  db_server = "server7.gullestrupnet.dk"
  db_bruger = "88707200_hoppeborg"
  db_password = "Sq3749pm!"
  db_database = "88707200_hoppeborg"  

 ' ------------------------------------------------------------

  function GenstandLedig(ID, startdato, slutdato)

   set a_conn = server.createObject("ADODB.connection")
   a_conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; PORT=3306; DATABASE=88707200_hoppeborg; UID=root; PWD=mysql;"    

   sql = "select * from reservationer where (re_start_dato <= '" & DKdate2USdate(slutdato) & "') AND (re_slut_dato >= '" & DKdate2USdate(startdato) & "' ) AND re_genstand_ID = '" & ID & "'"
 
   set rs = a_conn.execute(sql)
    if rs.eof then 
     GenstandLedig = true
    else
     GenstandLedig = false
    end if  
    rs.close 
   set rs = nothing  
   
   a_conn.close
   set a_conn = nothing 
   
  end function  
  

' ------------------------------------------------------------
'   DATABASE gullestrup.net ordre
' ------------------------------------------------------------

  ' or_server = "gn51.gullestrupnet.dk"
  ' or_bruger = "gnordre"
  ' or_password = "G0d8lz5&7"
  ' or_database = "gnordre"
  or_server = "localhost"
  or_bruger = "root"
  or_password = "mysql"
  or_database = "88707200_hoppeborg"

' ------------------------------------------------------------

  function HentNaesteFakturanummer

   set a_conn = server.createObject("ADODB.connection")
   a_conn.open "DRIVER={MySQL ODBC 8.0 ANSI Driver};SERVER=" & or_server & ";UID=" & or_bruger & ";PWD=" & or_password & ";DATABASE=" & or_database    

   set rs = a_conn.execute("select * from t_setup order by id desc limit 1")
   
    if rs.eof then
     HentNaesteFakturanummer = 0
    else
     HentNaesteFakturanummer = int(rs("value"))
    end if  

    a_id = rs("id")

    rs.close 
   set rs = nothing  
   
   a_conn.execute("update t_setup set value = '" & HentNaesteFakturanummer + 1 & "' where id = '" & a_id & "'")
   
   a_conn.close
   set a_conn = nothing 
  
  end function

' ------------------------------------------------------------
'   RECORD SET
' ------------------------------------------------------------
 
 function Text2Sql(T2SText)
  if T2SText <> "" then Text2Sql = replace(T2SText, "'", "''") 
 end function

' ------------------------------------------------------------

 function Sql2Text(S2TText)
  if S2TText <> "" then Sql2Text = replace(S2TText, """", "''") 
 end function  

' ------------------------------------------------------------

 function CrLf2BR(C2B_text)
  if C2B_text <> "" then CrLf2BR = replace(C2B_text, vbCrLf, "<br>")
 end function

' ------------------------------------------------------------
'   ?VRIGE PROCEDURER OG FUNCTIONER
' ------------------------------------------------------------

  function SendHTMLEmail(sm_modtager, sm_emne, sm_tekst, sm_HTML)    
   ' Set objMessage = CreateObject("CDO.Message")
   ' objMessage.From = "babybear@valloon.me"
   ' objMessage.To = sm_modtager
   ' objMessage.Subject = sm_emne
   ' objMessage.HtmlBody = sm_HTML

   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "web.valloon.me"
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "babybear@valloon.me"
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "babybear"
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = "false"
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
   ' 'objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = "true"
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

   Set objMessage = CreateObject("CDO.Message")
   objMessage.From = "samsimon623@outlook.com"
   objMessage.To = sm_modtager
   objMessage.Subject = sm_emne
   objMessage.HtmlBody = sm_HTML

   Response.Write("start")
   objMessage.Configuration.Load -1
   Response.Write("start")
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false
   ' objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusetls") = true
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpstarttls") = true
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "samsimon623@outlook.com"
   objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "kimujin123!@#"

   objMessage.Configuration.Fields.Update

   On Error Resume Next
   objMessage.Send
   If Err.Number <> 0 Then
    Response.Write "Error: " & Err
   End If

   Set objMessage = Nothing

	 Response.Write "Message sent succesfully!"
  end function  
 
 ' ------------------------------------------------------------

  function EmailGyldig(EG_email)
   if (EG_email = "") OR (instr(EG_email, "@") < 3) OR (instr(mid(EG_email, 5, 100), ".") = 0) OR (len(mid(EG_email, 5, 100)) < 3) OR (instr(right(EG_email, 4), ".") = 0) then
    EmailGyldig = false
   else
    EmailGyldig = true
   end if  
  end function
  
  ' ------------------------------------------------------------

  function HentFaktura
    set fso = CreateObject("Scripting.FileSystemObject")
    set f = fso.OpenTextFile(Server.MapPath("hoppeborg_faktura.txt"), 1, True)
    HentFaktura = f.ReadAll
    f.Close
    set f = fso.OpenTextFile(Server.MapPath("hoppeborg_lejebetingelser.txt"), 1, True)
    HentFaktura = HentFaktura & "<br><br><br>" & f.ReadAll
    f.Close    
  end function

' ------------------------------------------------------------

 function DKdate2USdate(DK2US_date)
  if NOT isdate(DK2US_date) then
   DKdate2USdate = ""
  else 
   DKdate2USdate = year(DK2US_date) & "-" & month(DK2US_date) & "-" & day(DK2US_date)
  end if 
 end function

%>