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

<%Set rs= conn.Execute("SELECT * FROM tbl_ogrenci ORDER BY numara")%>

<table width="500px" align="center" class="table" cellpadding="1" cellspacing="1" style="font-size:11px">
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
			<%if (rs("vize")*0.4+rs("final")*0.6)>=59.5 AND rs("final")>=50 then%>
				GEÇTÝ
			<%else%>
				KALDI
			<%end if%>
		</td>
	</tr>
	<%rs.MoveNext
	Loop%>
</table>

<br><br>

<script type="text/javascript">
function kayit_kontrol()
	{
	var returnVAL = true;
	
	if (document.getElementById('numara').value=='' || isNaN(document.getElementById('numara').value) || document.getElementById('numara').value<=0)
		{
		alert("NUMARA Alaný Boþ Olamaz ve Sayýsal Olmayan Ýfade Ýçeremez!");
		document.getElementById('numara').value="";
		document.getElementById('numara').focus();
		returnVAL =false;
		}
	else if (document.getElementById('ad').value=='')
		{
		alert("AD Alaný Boþ Olamaz!");
		document.getElementById('ad').focus();
		returnVAL =false;
		}
	else if (document.getElementById('soyad').value=='')
		{
		alert("SOYAD Alaný Boþ Olamaz!");
		document.getElementById('soyad').focus();
		returnVAL =false;
		}
	else if (document.getElementById('vize').value=='' || isNaN(document.getElementById('vize').value) || document.getElementById('vize').value<=0)
		{
		alert("VÝZE Alaný Boþ Olamaz ve Sayýsal Olmayan Ýfade Ýçeremez!");
		document.getElementById('vize').value="";
		document.getElementById('vize').focus();
		returnVAL =false;
		}
	else if (document.getElementById('final').value=='' || isNaN(document.getElementById('final').value) || document.getElementById('final').value<=0)
		{
		alert("FÝNAL Alaný Boþ Olamaz ve Sayýsal Olmayan Ýfade Ýçeremez!");
		document.getElementById('final').value="";
		document.getElementById('final').focus();
		returnVAL =false;
		}
	return returnVAL;
	}
</script>

<table width="200px" align="center" cellpadding="5" cellspacing="0" style="font-size:11px">
	<tr align="left">
		<td colspan="2">YENÝ KAYIT EKLEME ALANI</td>
	</tr>
	<form method="post" onsubmit="return kayit_kontrol()" action="kaydet.asp">
	<tr align="left">
		<td>NUMARA</td>
		<td><input type="text" name="numara" value="20"></td>
	</tr>
	<tr>
		<td>AD</td>
		<td><input type="text" name="ad" value="ali"></td>
	</tr>
	<tr>
		<td>SOYAD</td>
		<td><input type="text" name="soyad" value="veli"></td>
	</tr>
	<tr>
		<td>VÝZE</td>
		<td><input type="text" name="vize" value="49"></td>
	</tr>
	<tr>
		<td>FÝNAL</td>
		<td><input type="text" name="final" value="50"></td>
	</tr>
	<tr>
		<td><input type="Submit" value="Kaydet"></td>
		<td><input type="Reset" value="Temizle"></td>
	</tr>
	</form>
</table>

<br><br>

<table width="300px" align="center" cellpadding="5" cellspacing="0" style="font-size:11px">
	<tr align="left">
		<td colspan="2">KAYIT SÝLME ALANI</td>
	</tr>
	<form method="post" action="sil.asp">
	<tr align="left">
		<td>Silinecek Numara</td>
		<td>
		<%Set rs= conn.Execute("SELECT * FROM tbl_ogrenci ORDER BY numara")%>
		<select name="numara">
			<option>Silinecek Numarayý Seçiniz</option>
			<%Do while Not rs.eof%>
			<option value=<%=rs("numara")%>><%=rs("numara")%>
			<%rs.MoveNext
			Loop%>
		</select>
		</td>
	</tr>
	<tr>
		<td colspan="2" align="center"><input type="Submit" value="Sil"></td>
	</tr>
	</form>
</table>

<%
rs.close
Set rs=nothing
conn.close
set conn=nothing
%>