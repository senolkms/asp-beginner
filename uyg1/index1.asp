<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("vt.mdb")
%>

<%Set rs= conn.Execute("SELECT * FROM tbl_ogrenci")%>

<table cellpadding="0" cellspacing="1" bgcolor="#000000" width="600px" align="center">
	<tr bgcolor="#ffffff">
		<td>NUMARA</td>
		<td>AD</td>
		<td>SOYAD</td>
		<td>VÝZE</td>
		<td>FÝNAL</td>
		<td>ORTALAMA</td>
		<td>SONUÇ</td>
	</tr>
<%Do while Not rs.eof%>
	<tr bgcolor="#ffffff">
		<td><%=rs("numara")%></td>
	</tr>
	<%rs.MoveNext
Loop%>
</table>


<%rs.MoveFirst
Do while Not rs.eof%>
	<%=rs("numara")%>-<%=rs("ad")%>-<%=rs("soyad")%>-<%=rs("vize")%>-<%=rs("final")%>-<%=rs("vize")*0.4+rs("final")*0.6%>-
	<%if ((rs("vize")*0.4+rs("final")*0.6)>=60 AND rs("final")>=50) then
		response.write("Geçti")
	else
		response.write("Kaldý")
	end if%><br>
	<%rs.MoveNext
Loop%>

<%
rs.close
Set rs=nothing
conn.close
set conn=nothing
%>