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

<%Set rs= conn.Execute("SELECT * FROM tbl_ogrenci")%>

<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td height="50" class="tddis"><b>NUMARA</b></td>
		<td class="tddis"><b>ADI</b></td>
    	<td class="tddis"><b>SOYADI</b></td>
		<td class="tddis"><b>VÝZE</b></td>
		<td class="tddis"><b>FÝNAL</b></td>
		<td class="tddis"><b>ORTALAMA</b></td>
		<td class="tddis"><b>DURUM</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("numara")%></td>
    	<td class="tdic"><%=rs("ad")%></td>
		<td class="tdic"><%=rs("soyad")%></td>
		<td class="tdic"><%=rs("vize")%></td>
		<td class="tdic"><%=rs("final")%></td>
		<td class="tdic"><%=rs("vize")*0.4+rs("final")*0.6%></td>
		<td class="tdic">
			<%if (rs("vize")*0.4+rs("final")*0.6)>=60 AND rs("final")>=50 then%>
				GEÇTÝ
			<%else%>
				KALDI
			<%end if%>
		</td>
	</tr>
	<%rs.MoveNext
	Loop%>
	<tr>
		<td colspan="7" class="tdic" align="center">
			<form method="post" action="goster.asp">
				Ortalamaya Göre Sorgu<input type="text" name="ortalama">
				<input type="Submit" value="Göster">
			</form>
		</td>
	</tr>
</table>

<%
rs.close
Set rs=nothing
conn.close
set conn=nothing
%>