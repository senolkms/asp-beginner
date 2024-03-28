<%numara=request.form("numara")
ad=request.form("ad")
soyad=request.form("soyad")
vize=request.form("vize")
final=request.form("final")%>

<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("vt.mdb")
%>

<%Set rs= conn.Execute("INSERT INTO tbl_ogrenci (numara, ad, soyad, vize, final) VALUES(" & numara & ", '" & ad & "', '" & soyad & "', " & vize & ", " & final & ")")%>

<table width="200px" align="center" cellpadding="5" cellspacing="0" style="font-size:11px">
	<tr align="left">
		<td>NUMARA</td>
		<td><%=numara%></td>
	</tr>
	<tr>
		<td>AD</td>
		<td><%=ad%></td>
	</tr>
	<tr>
		<td>SOYAD</td>
		<td><%=soyad%></td>
	</tr>
	<tr>
		<td>VÝZE</td>
		<td><%=vize%></td>
	</tr>
	<tr>
		<td>FÝNAL</td>
		<td><%=final%></td>
	</tr>
	<tr>
		<td>ORTALAMA</td>
		<td><%=vize*0.4+final*0.6%></td>
	</tr>
	<tr>
		<td>DURUM</td>
		<td>
			<%if (vize*0.4+final*0.6)>=59.5 AND final>=50 then%>
				GEÇTÝ
			<%else%>
				KALDI
			<%end if%>
		</td>
	</tr>
	<tr bgcolor="#ffffff">
		<td colspan="2">KAYIT GERÇEKLEÞTÝ</td>
	</tr>
	<tr bgcolor="#ffffff">
		<td colspan="2"><a href="index.asp">GERÝ DÖN</A></td>
	</tr>
</table>