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
<%numara=request.form("numara")%>

<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("vt.mdb")
%>

<%Set rs= conn.Execute("DELETE FROM tbl_ogrenci WHERE numara=" & numara)%>

<script type="text/javascript">
	alert("Kayýt Silinmiþtir!");
	window.location.href("index.asp");
</script>