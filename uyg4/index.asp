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

<%Set rs= conn.Execute("SELECT * FROM tbl_sayi ORDER BY sayi")%>
<table width="500px" align="center" cellpadding="1" cellspacing="1" style="font-size:11px">
	<tr align="center">
    	<td width="20%" height="50" class="tddis"><b>SAYI</b></td>
		<td width="20%" class="tddis"><b>KARE</b></td>
    	<td width="20%" class="tddis"><b>KÜP</b></td>
		<td width="20%" class="tddis"><b>FAKTÖRÝYEL</b></td>
		<td width="20%" class="tddis"><b>ASAL</b></td>
	</tr>

	<%Do while Not rs.eof%>
  	<tr align="center">
    	<td height="40" class="tdic"><%=rs("sayi")%></td>
    	<td class="tdic"><%=rs("sayi")*rs("sayi")%></td>
		<td class="tdic"><%=rs("sayi")*rs("sayi")*rs("sayi")%></td>
		<td class="tdic">
			<%sonuc=1
			for i1=2 to rs("sayi")
				sonuc=sonuc*i1
			next
			response.write(sonuc)
			%>
		</td>
		<td class="tdic">
			<%drm=false
			for i1=2 to rs("sayi")-1
				if rs("sayi") mod i1=0 then
					drm=true
					exit for
				end if
			next
			if drm=true then
				response.write("Asal Deðil")
			else
				response.write("Asal")
			end if%>
		</td>
	</tr>
		<%
	rs.MoveNext
	Loop%>
</table>

<%
rs.close
Set rs=nothing
conn.close
set conn=nothing
%>