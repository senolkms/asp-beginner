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
<!--Set rs= conn.Execute("INSERT INTO tbl_ogrenci (numara, ad, soyad, vize, final) VALUES(20, 'ali', 'veli', 49, 50)")-->

<script type="text/javascript">
	alert("Kayýt Gerçekleþti!");
	window.location.href("index.asp");
</script>