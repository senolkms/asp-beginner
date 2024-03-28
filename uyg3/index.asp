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

<%Set rs= conn.Execute("SELECT * FROM tbl_market ORDER BY id")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td width="14%" height="50" class="tddis"><b>NUMARA</b></td>
		<td width="30%" class="tddis"><b>ADI</b></td>
    	<td width="14%" class="tddis"><b>ALIÞ</b></td>
		<td width="14%" class="tddis"><b>SATIÞ</b></td>
		<td width="14%" class="tddis"><b>ADET</b></td>
		<td width="14%" class="tddis"><b>TUTAR</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("id")%></td>
    	<td class="tdic"><%=rs("ad")%></td>
		<td class="tdic"><%=rs("alis")%></td>
		<td class="tdic"><%=rs("satis")%></td>
		<td class="tdic"><%=rs("adet")%></td>
		<td class="tdic"><%=rs("satis")*rs("adet")%></td>
	</tr>
		<%rs.MoveNext
	Loop%>
</table><br><br>

<%Set rs= conn.Execute("SELECT * FROM tbl_market WHERE adet*satis>=3000")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td width="14%" height="50" class="tddis"><b>NUMARA</b></td>
		<td width="30%" class="tddis"><b>ADI</b></td>
    	<td width="14%" class="tddis"><b>ALIÞ</b></td>
		<td width="14%" class="tddis"><b>SATIÞ</b></td>
		<td width="14%" class="tddis"><b>ADET</b></td>
		<td width="14%" class="tddis"><b>TUTAR</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("id")%></td>
    	<td class="tdic"><%=rs("ad")%></td>
		<td class="tdic"><%=rs("alis")%></td>
		<td class="tdic"><%=rs("satis")%></td>
		<td class="tdic"><%=rs("adet")%></td>
		<td class="tdic"><%=rs("satis")*rs("adet")%></td>
	</tr>
		<%rs.MoveNext
	Loop%>
</table><br><br>

<%Set rs= conn.Execute("SELECT * FROM tbl_market WHERE (satis-alis)/alis>=0.5")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td width="14%" height="50" class="tddis"><b>NUMARA</b></td>
		<td width="30%" class="tddis"><b>ADI</b></td>
    	<td width="14%" class="tddis"><b>ALIÞ</b></td>
		<td width="14%" class="tddis"><b>SATIÞ</b></td>
		<td width="14%" class="tddis"><b>ADET</b></td>
		<td width="14%" class="tddis"><b>TUTAR</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("id")%></td>
    	<td class="tdic"><%=rs("ad")%></td>
		<td class="tdic"><%=rs("alis")%></td>
		<td class="tdic"><%=rs("satis")%></td>
		<td class="tdic"><%=rs("adet")%></td>
		<td class="tdic"><%=rs("satis")*rs("adet")%></td>
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