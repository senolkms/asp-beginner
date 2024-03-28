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

<%Set rs= conn.Execute("SELECT * FROM tbl_musteri ORDER BY mus_id")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td height="50" class="tddis"><b>M ID</b></td>
		<td class="tddis"><b>M AD</b></td>
    	<td class="tddis"><b>U ID</b></td>
		<td class="tddis"><b>U AD</b></td>
		<td class="tddis"><b>U ALIÞ</b></td>
		<td class="tddis"><b>U SATIÞ</b></td>
		<td class="tddis"><b>SATIÞ ADET</b></td>
		<td class="tddis"><b>TOPLAM</b></td>
		<td class="tddis"><b>ÖDENEN</b></td>
		<td class="tddis"><b>KALAN</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("mus_id")%></td>
    	<td class="tdic"><%=rs("mus_ad")%></td>
		<td class="tdic"><%=rs("urun_id")%></td>
		<td class="tdic"><%=rs("urun_ad")%></td>
		<td class="tdic"><%=rs("urun_alis")%></td>
		<td class="tdic"><%=rs("urun_satis")%></td>
		<td class="tdic"><%=rs("satis_adet")%></td>
		<td class="tdic"><%=rs("urun_satis")*rs("satis_adet")%></td>
		<td class="tdic"><%=rs("odenen")%></td>
		<td class="tdic">
			<%if ((rs("urun_satis")*rs("satis_adet"))=rs("odenen")) then%>
				<b><font color="#00ff00">0</font></b>
			<%else%>
				<%=rs("urun_satis")*rs("satis_adet")-rs("odenen")%>
			<%end if%>
		</td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table><br><br>



<%Set rs= conn.Execute("SELECT * FROM tbl_musteri WHERE urun_satis*satis_adet>=600 ORDER BY mus_id")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
		<td class="tddis"><b>M AD (600'den fazla alýþ-veriþ yapan müþterilerin adý)</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td class="tdic"><%=rs("mus_ad")%></td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table><br><br>



<%Set rs= conn.Execute("SELECT * FROM tbl_musteri WHERE (urun_satis*satis_adet-odenen)>200 ORDER BY mus_id")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
		<td class="tddis"><b>M AD (borcu 200'den fazla olan müþterilerin adý)</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td class="tdic"><%=rs("mus_ad")%></td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table><br><br>



<%Set rs= conn.Execute("SELECT * FROM tbl_musteri WHERE (urun_satis*satis_adet-odenen)=0 ORDER BY mus_id")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
		<td class="tddis"><b>M AD (borcu olmayan müþterilerin adý)</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td class="tdic"><%=rs("mus_ad")%></td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table><br><br>



<%Set rs= conn.Execute("SELECT COUNT(*) AS sayac FROM tbl_musteri WHERE urun_ad='HDD'")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
		<td class="tddis"><b>ADET (HDD alan müþterilerin sayýsý)</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td class="tdic"><%=rs("sayac")%></td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table><br><br>



<%Set rs= conn.Execute("SELECT SUM(satis_adet) AS sayac FROM tbl_musteri WHERE urun_ad='HDD'")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
		<td class="tddis"><b>ADET (HDD alan müþterilerin satýþ adetleri toplamý)</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td class="tdic"><%=rs("sayac")%></td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table><br><br>



<%Set rs= conn.Execute("SELECT urun_id, urun_ad, urun_alis FROM tbl_musteri WHERE urun_alis>(SELECT AVG(urun_alis) FROM tbl_musteri)")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
		<td class="tddis"><b>U ID (1)</b></td>
		<td class="tddis"><b>U AD</b></td>
		<td class="tddis"><b>U ALIÞ</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td class="tdic"><%=rs("urun_id")%></td>
    	<td class="tdic"><%=rs("urun_ad")%></td>
    	<td class="tdic"><%=rs("urun_alis")%></td>
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