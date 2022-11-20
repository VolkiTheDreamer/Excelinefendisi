<%@ Page Title='Özel Olaylar' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Olaylar'>
</asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='6'></asp:Label></td></tr></table></div>
<h1>Özel Olaylar</h1>
	<p>Bundan önceki bölümlerde Excel'in kendi nesnelerinin eventlerinin otomatik 
	olarak kod pencresine 
	eklendiğini (ve bunların değiştirilmemesi gerektiğini) söylemiştik. Bu bölümde kendi nesnelerimize(classlarınıza) ait eventleri nasıl 
	yaratırız bunu göreceğiz. Bunun için Class kavramını biraz biliyor 
	olmanız gerekiyor. Bilmeyenler ön bilgiyi
	<a href="Ileriseviyekonular_ClassveClassModuller.aspx">buradan</a> 
	edinebilirler.&nbsp;</p>
	<h2>Kendi Classlarınıza ait olaylar</h2>
	<div class='konu'>
		<p>
		Bu bölümü yapmaya başladığımda nerden başlayacağıma bi türlü karar 
		veremedim, zira Class konusuna bayağı bir girmek gerekiyordu. Şuan ise 
		sayfaları sırayla hazırladığım için ve henüz Class konusuna girmediğim 
		için biraz anlamsız geldi. O yüzden bu konuyu üstteki paragrafta 
		belirttiğim linke bırakıyorum.&nbsp;</p>
		<h2>
		Diğer özel olaylar</h2>
		<p>Her ne kadar özel olay olmasa da farklı tanımlama ve erişim şekilleri 
		itibarıyle klasik olaylardan farklı oldukları için 
		<span class="keywordler">WithEvent</span> deyimi ile 
		tanımlanan olayları da özel olay gibi düşünebiliriz. Bundan önce <strong>
		Grafik 
		</strong>ve <strong>Application </strong>için bunları nasıl yazdığımızı görmüştük. Bu tür özel olayları 
		başka nesneler 
		için de kullanabiliyoruz. Tüm kullanılabilir nesnelerin listesini 
		aşağıdaki gibi bir değişken tanımlarken <strong>As</strong>'den sonra 
		boşluk tuşuna basınca görebilirsiniz.&nbsp;</p>
		<p><img src="../../images/vbaeventcustom1.jpg"></p>
		<p>Bunlardan ComboBox, ListBox, Image v.s gibi <strong>ActiveX</strong> 
		kontrollerine ait eventleri zaten Formlar konusunda ayrıca ele alıyor 
		olacağız. Bu sayfadaki konsepte göre tanımlamanın pek bir esprisi yok 
		bence. O yüzden bunları geçiyoruz.</p>
		<p>Biz burada sadece <strong>QueryTable </strong>nesnesinin eventlerine 
		bakacağız. 
		ActiveX eventlerine ise
		<a href="Formlar_Kontroller.aspx">buradan</a> ulaşabilirsinz.</p>
		<h3>QueryTable Event örneği</h3>
		<p>Aşağıdaki kod ile kullanıcıyı uzun bir refresh işlemi için 
		uyarıyoruz, refresh işlemi bitince de bir mesaj kutusu ile haber 
		veriyoruz.</p>
		<p>Örnek dosyayı
		<a href="../../Ornek_dosyalar/Makrolar/eventQTclass.xlsm">buradan</a> 
		indirebilirsiniz. Aşamalarımız şöyle:</p>
		<ul>
			<li>Öncelikle yeni bir dosya açın</li>
			<li>VBE'ye geçip bir class modül ekleyin ve propertiesten buna 
			myClass adını verin.</li>
			<li>Sonra aşağıdaki kodu bu modüle yapıştırın.</li>
			<li>Sonrasında dosyanıza bir veri bağlantısı ekleyin(Access, Excel, 
			Oracle, SQL Server v.s olabilir). Örnek olması adına benim sitemdeki
			<a href="http://www.excelinefendisi.com/Ornek_dosyalar/pivotdata.xlsx">
			http://www.excelinefendisi.com/Ornek_dosyalar/pivotdata.xlsx</a> 
			dosyasından herhangi bir sayfayı ekleyebilirsiniz.(Veri tabanı 
			bağlantılarına aşina değilseniz önce
			<a href="../Excel/DataMenusu_BaskaVeriKaynaklariilecalismak.aspx">
			buraya</a> bakın)</li>
		</ul>
		<pre class="brush:vb">
Private WithEvents myQT As QueryTable
Private Sub Class_Initialize()
    Set myQT = ActiveSheet.ListObjects(1).QueryTable
End Sub
Private Sub myQT_AfterRefresh(ByVal Success As Boolean)
If Success = True Then
    MsgBox "refresh işlemi bitti"
Else
    MsgBox "refresh sırasında bir hata oluştu"
End If
End Sub
Private Sub myQT_BeforeRefresh(Cancel As Boolean)
    cevap = MsgBox("uzun sürecek, iptal edeyim mi", vbYesNo)
    If cevap = vbYes Then
        Cancel = True
        MsgBox "iptal edildi"
    End If
End Sub
</pre>
		<ul>
			<li>Son olarak <strong>ThisWorkbook </strong>modülünün içine de 
			şunları yazın.</li>
		</ul>
		<pre class="brush:vb">
Dim myClassNesnesi As myClass
Private Sub Workbook_Open()
   Set myClassNesnesi = New myClass
End Sub</pre>
		<p>
		Bu kadar basit. Şimdi dosyanızı kaydedip kapatın ve tekrar açın, sonra 
		da data üzerinde bir yere gelip, sağ tıklayıp <strong>Refresh </strong>
		diyin. Aşağıdaki mesajı görmeniz lazım.</p>
		<p>
		<img src="../../images/vbaeventcustom2.gif"></p>
		<p>
		Kodun çalışma prensibi şöyle:</p>
		<p>
		Bir class modülümüz var. Dosya açılır açılmaz bu classtan bir nesne 
		yaratılıyor. Bu nesne yaratılınca class'ın Initialize eventi devreye 
		giriyor ve bu sefer de myQT nesnesine sayfadaki 1 indexli QueryTable 
		atanıyor. Sonrasındaki refresh işlemleri ise aşikar.</p></div>
</asp:Content>
