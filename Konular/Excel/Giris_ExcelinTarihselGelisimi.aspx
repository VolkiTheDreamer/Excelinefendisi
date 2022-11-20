<%@ Page Title='Excelin Tarihsel Gelişimi' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td>

<asp:Label ID='Label1' runat='server' Text='Excel'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Excelin Tarihsel Gelişimi</h1><p>
	Bu bölümde Excel'in 2007 versiyonundan itibaren geçirdiği değişime bakıyor olacağız.
	Artık 2003 ve önceki versiyonları kullanan çok kalmamıştır diye düşünüyorum, 
	o yüzden 2003 ve öncesine bakmayacağız.</p>
	<p>
	Bazı versiyonlarda radikal değişimler oluyorken bazıları minör 
	değişikliklerle idare ediyorlar.
	Bu arada şöyle de birşey var ki, ilgili Office sürümünün Business versiyonunda olan bir özellik Home 
	versiyonunuda olmayabiliyor. Mesela ben evde Office 2016 kullanıyorum, bende 
	2013le birlikte gelen Inquire menüsü çıkmıyor. Keza PowerPivot denen harika 
	araç da sadece Professional Plus versiyonunda bulunuyor.</p>
	<p>
	Son olarak, bi genel kültür bilgisi de vermiş olalım. MS Office artık 2013 
	versiyonundan itibaren Office 365 çatısı altında da sunuluyor. Office 365, 
	kullanıcılara bulutta(OneDrive) depolama, Office Apps kullanma, 
	Ücretsiz Skype hizmeti gibi imkanlar sunuyor. Ayrıca benim de şuan evde 
	kullandığım Home sürümü ile 1 TB bulut alanı sağlıyor. 365'in bir özelliği 
	de artık, bi kere parasnı verip ürünü almış olmuyorsunuz, aylık ve yıllık 
	abonelike kaydoluyorsunuz, her yeni office sürümünü de ücretsiz elde etme 
	hakkına sahip olmuş oluyorsunuz, ki bu bence harika bir yöntem olmuş. Office 365'i 
	biraz incelemenizi tavsiye ederim. Bu
	<a href="https://www.microsoftstore.com/store/msusa/en_US/cat/Compare-Office-suites/categoryID.68155000">
	linkten</a> de sürümler hakkında detaylı bilgi edinebilirsiniz.</p>
	<p>
	Şimdi her versiyonda gelen yeniliklere bakmadan önce versiyon numaralarına 
	bakalım, bunlar özellikle makro yazımında işinize yarayacak. <a href="/Konular/VBAMakro/DortTemelNesne_Application.aspx#Versiyon">VBA>Application</a> sayfasında
	göreceğimiz gibi bir kodla kullanıcının excel sürümünü kontrol 
	edebilirsiniz. </p>
	<table class="alterantelitable">
		<th style="text-align: center">Versiyon Adı/Yılı</th>
		<th style="text-align: center">Versiyon Numarası</th>
		<tr>
			<td style="text-align: center">Excel 2003</td>
			<td style="text-align: center">11</td>
		</tr>
		<tr>
			<td style="text-align: center">Excel 2007</td>
			<td style="text-align: center">12</td>
		</tr>
		<tr>
			<td style="text-align: center">Excel 2010</td>
			<td style="text-align: center; text-decoration: underline;" title="uğursuz sayı 13 atlanmış :)">14</td>
		</tr>
		<tr>
			<td style="text-align: center">Excel 2013</td>
			<td style="text-align: center">15</td>
		</tr>
		<tr>
			<td style="text-align: center">Excel 2016</td>
			<td style="text-align: center">16</td>
		</tr>
	</table>
	<h2 class='baslik'>2016 ile gelen yenilikler</h2>
	<div class='konu'>
		<p>Şahsi fikrim 2016 versiyonunda çok büyük değişimlerin olmadığıdır, 
		gerçi yeni formüller gayet faydalı, nerdeyse hepsi için ayrı ayrı VBA 
		ile UDF yazmıştım, bunlar çöpe gidecek ama yine de eski versiyon 
		kullanıcıları bu UDF'leri kullanabilir. Şimdi kısaca yeniliklere bir 
		bakalım.</p>
		<ul>
			<li>Power Query entegrasyonu(Data Menüsü altında&gt;Get&amp;Transform)</li>
			<li>Power Map entegrasyonu, 3D maps olarak Insert menüsü altında</li>
			<li>Dosya kaydederken Read-only modunu zorlama seçeneği</li>
			<li>Pivot Table ve Slicerlara klavye erişimi</li>
			<li>Yeni grafik türleri</li>
			<li>Tahminleme aracı ve formülleri(FORECAST)</li>
			<li>Yeni formüller(TEXTJOIN,IFS,SWITCH,MAXIFS,MINIFS</li>
			<li>Pivot tablolarda otomatik tarih gruplama</li>
			<li>Data kartları(Power Viewda)</li>
			<li>Smart lookup</li>
			<li>ve münferit birkaç değişiklik daha...</li>
		</ul>
		<p>Bu
		<a href="https://support.office.com/en-us/article/What-s-new-in-Excel-2016-for-Windows-5fdb9208-ff33-45b6-9e08-1f5cdb3a6c73">
		linkte</a> daha detaylı bilgi bulabilirsiniz.</p>
		</div>
		
<h2 class='baslik'>2013 ile gelen yenilikler</h2>
	<div class='konu'>
		<p>Yine şahsi fikrimi söyleyeceğim, 2016 versiyonundan daha çok yenilik 
		barındırıyor ama 2010'un yenilikleri daha fazla ve faydalıydı. Hemen 
		bakalım:</p>
		<ul>
			<li>Artık her dosya kendi ayrı penceresinde(İlk başta çok rahatsız edici ama 
			zamanla alışıyorsunuz, özellikle 2 monitor kullanıyorsanız çok 
			faydalı)</li>
			<li>FlashFill(Özellikle metin formülleri yazamayan kişiler için çok 
			faydalı olmuştur)</li>
			<li>Excel Data Model<ul>
				<li>Distinct count(Yıllardır bunu bekliyordum)</li>
				<li>Birden çok tabloya dayanan Pivot tablolar</li>
			</ul>
			</li>
			<li>Power Query</li>
			<li>Power Map</li>
			<li>Power View</li>
			<li>Eskiden ayrı bir add-in olarak gelen PowerPivot sadece Professional 
			Plus versiyonu içine alındı, 2010'da parayla 
			alabiliyordunuz, bu versiyonla para ödeseniz bile alamıyorsunuz</li>
			<li>Özet(Pivot) Tablo önerileri</li>
			<li>Özet(Pivot) Tablo için Timeline</li>
			<li>Table'lar için de Slicer özelliği</li>
			<li>Office Apps</li>
			<li>Inquire</li>
			<li>50 yeni fonksiyon(Bu
			<a href="https://support.office.com/en-us/article/New-functions-in-Excel-2013-075c82bd-15b9-4ad6-af31-55bb6b011cb9?CorrelationId=82306f3c-6b4c-4f2b-aa51-a6b7c3b41a58&amp;ui=en-US&amp;rs=en-US&amp;ad=US&amp;fromAR=1&amp;ocmsassetID=HA103980604">
			linkte</a> hepsi mevcut-DAYS, IFNA, ISFORMULA, ARABIC beğendiklerim)</li>
			<li>Quick Analysis</li>
			<li>Grafikler<ul>
				<li>Grafik önerileri</li>
				<li>Animasyonlu grafikler</li>
				<li>İyileştirmeler, </li>
				<li>Yeni grafikler, özellikle iki eksenli combo grafik, bu işi 
				zahmetlice yapmaktan kurtardı</li>
			</ul>
			</li>
			<li>Varsayılan kaydetme yeri Onedrive oldu(Microsoft, geleceği 
			bulutta gördü heralde)</li>
		</ul>
		<p>Bu
		<a href="https://support.office.com/en-us/article/What-s-new-in-Excel-2013-1cbc42cd-bfaf-43d7-9031-5688ef1392fd?CorrelationId=c81b62a0-6070-4a4d-bb48-937a7876c05e&amp;ui=en-US&amp;rs=en-US&amp;ad=US&amp;fromAR=1&amp;ocmsassetID=HA102809308">
		linkte</a> daha detaylı bilgi bulabilirsiniz.</p>
	</div>
	
	<h2 class='baslik'>2010 ile gelen yenilikler</h2>
	<div class='konu'>
		<p>Şahsi fikrim, 2007den sonraki en büyük gelişmeler bu versiyonda oldu.</p>
		<ul>
			<li>64 bit desteği geldi(Bu şu demek:4 GB'a kadar dosya boyutu)</li>
			<li>Pivot tablolarda geliştirme<ul>
				<li>
				Performance iyileştirmeleri</li>
				<li>
				2005'te yazdığım efsanevi Pivotta Boşlukları Otomatik Doldurma 
				makromu çöpe atan "Repat All Item Labels" özelliği</li>
				<li>
				Show Values As özelliği. Değerleri %sel, fark, rank v.s şeklinde 
				gösterme</li>
				<li>
				PivotChart iyileştirmeleri</li>
			</ul>
			</li>
			<li>Power Pivot Add'in (En büyük gelişmelerden biri)</li>
			<li>Daha fazla Conditional Formatting özelliği<ul>
				<li>
				Yeni ikon setleri</li>
				<li>
				Data barlar için yeni seçenekler</li>
			</ul>
			</li>
			<li>Resim editleme özellikleri</li>
			<li>Sparkline: Hücre içi grafikler(Harika)</li>
			<li>Yapıştırma öncesinde önizleme görebilme</li>
			<li><span>File Menüsü yeniden geldi</span></li>
			<li>Ribbonu özelliştirme imkanı. 2007'deyken sadece XML toollarını 
			bilenler yapabilirdi</li>
			<li>Yeni formüller(Kaydadeğer bir formül yok, genelde mevcut 
			formüllerin türevlerinin sunulması)</li>
			<li>Slicerlar(Sadece Özet Tablelarda)</li>
			<li>Solver add-in'ininde iyileştirme</li>
			<li>Filtrelere arama kutusu geldi</li>
			<li>Insert menüsünde Screenshots butonu</li>
			<li>Format As Table butonu(Çok faydalı)</li>
			<li>Ve daha birsürü ufak tefek iyileşme</li>
		</ul>
		<p>Bu
		<a href="https://support.office.com/en-us/article/What-s-New-in-Excel-2010-44316790-a115-4780-83db-d003e4a2b329">
		linkte</a> daha detaylı bilgi bulabilirsiniz.</p>
	</div>

<h2 class='baslik'>2007 ile gelen yenilikler</h2>
	<div class='konu'>
		<p>Bu versiyon, 97 versiyonundan beri hem nicelik hem nitelik açısından en çok değişikliği 
		içeren versiyondur, aradaki versiyonlarda çok da kritik değişiklikler 
		olmamıştır, ancak bu versiyon hem görünüm hem işlevesellik açısından 
		radikal değişimlere sahiptir.</p>
		<ul>
			<li>Ribbon yapısı sunulmuştur.</li>
			<li>XML tabanalı dosya yapısına geçiş, standart uzantı .xlsx oldu. 
			Dosya tipleri hakkında detay bilgi <a href="Giris_DosyaUzantilari.aspx">burda</a> 
		var.</li>
			<li>Satır ve sütun sayıları arttı. Satırlar 65 binden 1 milyon küsura ve kolonlar 256dan 
		16bin küsura.</li>
			<li>Özet tablolarda iyileştirme. Görünüm ve işlevesellik</li>
			<li>Remove duplicates özelliğie</li>
			<li>List ismi Table oldu.(Insert Table dediğimizde çıkan objeler)</li>
			<li>Grafiklerde iyileştirme</li>
			<li>Name Manegerda değişklik</li>
			<li>Formül &amp; Fonksyion<ul>
				<li>Yeni Fonksiyonlar:IFERROR, SUMIFS, AVERAGEIF, AVERAGEIFS, 
			COUNTIFS gibi kritik fonkisyonlar başta olmaz üzere.</li>
				<li>Boyutlanabilir formül çubuğu</li>
				<li>Formüllerde "auto complete" özelliği</li>
				<li>İçiçe en fazla 7 fonksiyon kısıtı 64e çıktı</li>
			</ul>
			</li>
			<li>Renge göre sıralama ve filtreleme</li>
			<li>Filtrede birden çok eleman seçebilme</li>
			<li>Daha gelişmiş conditional formatting</li>
			<li>Smart Art grafikler</li>
			<li>Pdf olarak kaydedebilme</li>
			<li>VBA ile oluştuduğunuz özel menüler en sağda Help'in yanına giderken 
		artık Add-ins altında bir menüye gitmeye başladı</li>
		</ul>
		<p>Bu
		<a href="https://technet.microsoft.com/en-us/library/cc179167(office.12).aspx">
		linkte</a> daha detaylı bilgi bulabilirsiniz.&nbsp;</p>
	</div>

</asp:Content>
