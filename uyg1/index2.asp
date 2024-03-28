<style type="text/css">

body
{
font-size: 15px;
}

table
{
background:#B36200;
}

.tdbas
{
background:#f6cd9f;
}

.tddis
{
background:#faecdc;
}

.tdic
{
background:#faf7c5;
}

</style>

<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("vt.mdb")
%>

<%Set rs= conn.Execute("SELECT * FROM tbl_ogrenci WHERE (vize*0.4+final*0.6)>=55")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td width="12%" height="50" class="tddis"><b>NUMARA</b></td>
		<td width="20%" class="tddis"><b>AD</b></td>
    	<td width="20%" class="tddis"><b>SOYAD</b></td>
		<td width="12%" class="tddis"><b>VÝZE</b></td>
		<td width="12%" class="tddis"><b>FÝNAL</b></td>
		<td width="12%" class="tddis"><b>ORTALAMA</b></td>
		<td width="12%" class="tddis"><b>DURUM</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("numara")%></td>
    	<td class="tdic"><%=rs("ad")%></td>
		<td class="tdic"><%=rs("soyad")%></td>
		<td class="tdic"><%=rs("vize")%></td>
		<td class="tdic">
			<%if (rs("final")>=50) then
				response.write("<font size='+3' face='Courier New'>"&rs("final")&"</font>")
			else
				response.write(rs("final"))
			end if	
			%>
		</td>
		<td class="tdic"><%=rs("vize")*0.4+rs("final")*0.6%></td>
		
		<%if (rs("vize")*0.4+rs("final")*0.6>=60 AND rs("final")>=50) then
			response.write("<td class='tdic'>Geçti")
		else
			response.write("<td bgcolor='#0000ff'>Kaldý")
		end if%>
		</td>
	</tr>
		<%rs.MoveNext
	Loop%>
</table>

<%
rs.close
Set rs=nothing
conn.close
set conn=nothing
%>