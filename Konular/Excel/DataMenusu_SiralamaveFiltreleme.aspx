<%@ Page Title='Sıralama ve Filtreleme' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='Excel'>
</asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Data Menüsü'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>Filtreleme ve Sıralama</h1>
<!--	<p>dvileri koymayı unutma</p>
	<p>Filter içine slicer da ekle</p>
	<p>
	<span style="color: rgb(0, 0, 0); font-family: SourceSansPro-Regular, &quot;Open Sans&quot;, sans-serif; font-size: 13px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: normal; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;">
	There is a big caveat to keep in mind: All of the cells that will be 
	involved in the sorting (or potentially involved in the sorting) must be 
	unlocked. This includes any column headings for the data that may be 
	sorted.(https://excelribbon.tips.net/T000137_Sorting_Data_on_Protected_Worksheets.html)</span></p>
<p>Excelin olmazsa olmazlarında Filtreleme ve Sıralama sanırım en basit seviyede Excel kullanıcılarının bile bildiği araçlardır. 
Biz de burada biraz az bilinen ve çeşitli püf noktaları üzerinde duracağız.</p>


	<h2 class="baslik">Sıralama</h2>
	<div class="konu">
	<p>Sıralamayla ilgili alt konular geçmeden önce belirtmek isterim ki 
	sıralama yaparken bulunduğun hücredeyken sıralama yapmanız yeterlidir, tüm 
	alanı seçmeye gerek yoktur.</p>
	<h3>Renge göre sıralama</h3>
	<p>Excel 2007 ile birlikte renge göre de sıralama özelliği gelmiştir.</p>
	<p>
	<img src="../../images/insertsortfilter2.jpg"></p>
	<p>
	Sıralamadan sona</p>
	<p>
	<img src="../../images/insertsortfilter3.jpg"></p>
        <p>
	        MIS açısından pratikteki kullanımı ise şöyle olabilir. Diyelim ki data listenize bir filtre uyguladınız ve bunları sarıya boyadınız. Sora başka bir filtre uyguladınız, bunları da mavi yaptınız. Sonra sarı ve mavileri başta görüp diğer boyasız kayılarla karşılaştırmak isteyebilirsiniz.</p>
	<h3>
	Sıralama sırası</h3>
	<p>
	Sıralama yapıldığında, Excel önce veri tiplerine göre sıralama, sonra da her 
	veri tipini kendi içinde boy sırası yapar. Veri tiplerinin artan şekilde 
	sıralaması şöyledir:</p>
	<ul>
		<li>Sayılar</li>
		<li>Metinler</li>
		<li>Mantıksal değerler(True,False)</li>
		<li>Hatalar</li>
		<li>Boş hücreler</li>
	</ul>
	<h3>Formüllü hücrelerde sıralama</h3>
	<p>Bazı durumlarda mevcut sayfadaki bir formülünüzde o sayfanını adı 
	bulunabilir. Aşağıdaki bir örnek bulunuyor. Bu liste şuan Bölge adına göre 
	dizili durumda. </p>
	<p><img src="../../images/insertsortfilter4.jpg"></p>
	<p>Şimdi bu listeyi Sapma oranına göre sıralayalım ve sorunu görelim.</p>
	<p><img src="../../images/insertsortfilter5.jpg"></p>
	<p>Bu problem, aynı sıralama işlemini ikinci kez yapınca düzelmekte birlikte 
	böyle bir maceraya grimeye gerek yoktur. Siz en iyisi mecut sayfa üzerine 
	bir hücreye başvuran bir formül yazacaksanız, formülde sayfa ismi olmamasına 
	özen gösterin. Bu durum özellikle formül yazarken başka bir sayfaya gidip 
	geri geldiğimizde yaşanır. Böyle durumlarda sayfa adını manuel 
	silebilrsiniz.&nbsp;</p>
	<h3>Sıralama seçenekleri</h3>
	<p>Sıralamaların %99u yukarıdan aşağıdır ancak Excel'in soldan sağa sıralama 
	özelliği de bulunuyor. Sort menüsünden Sort Options'a tıklayınca Orientation 
	bölümünden seçebiliyorsunuz. Burada, gördüğünüz gibi bir de küçük/büyük 
	harfe duyarlı sıralama yapma ayarı da bulunmakta.</p>
	<p><img src="../../images/insertsort5.jpg"></p>
	<p>Aşağıdaki örnekte Aylık Grçekleşen <strong>satırı</strong> Büyükten 
	küçüke sıralanmıştır.</p>
	<p><img src="../../images/insertsort6.jpg"></p>
	<h3>Sıralamada başlık uçmasın</h3>
	<p>Excel'in başlık satırınızı başlık olarak algılaması için bazı kriterler 
	bulunur. Bunlar;</p>
	<ul>
		<li>Başlık satırında boş hücre olmamalı</li>
		<li>Başlık satırı alttaki data grubundan fakrlı şeklide formatlanmış 
		olmalı. Bold, renkli v.s</li>
		<li>Data grubu tamamen metinse ve başlıklar da metinse başlık olarak 
		algılanmaz</li>
	</ul>
	<h3>Merged cell olan bir listede sıralama</h3>
	<p>Normalde merged cell olan bir lstede sıralama yapılamaz. Bunu yapmayaı 
	sağlayan bir makro
	<a href="http://excel.tips.net/T002581_Sorting_Data_Containing_Merged_Cells.html">
	burada</a> var ama ben böyle bir makroya hiç ihityaç duymadım. Onun yerine 
	geçici olarak merged cell'leri birbiriniden ayırır, sıralamayı yapar, sonra 
	gerekirse tekrar merge'lerim. Hatta birçok durumda genelde merged cell'in 
	diptoplam satırında olduğu bir listede sıralamayı yapmak için en alt satırla 
	diptoplam arasına bir satır açar ve bunu gizlerim veya satır yüksekliğini 
	çok düşük belirlerim. </p>
	<p><img src="../../images/insertsortfilter6.jpg"></p>
	<p>Satır açtıktan sonra</p>
	<p><img src="../../images/insertsortfilter7.jpg"></p>
	<p>&nbsp;</p>
	
	</div>
	
	<h2 class="baslik">Filtreleme</h2>
	<div class="konu">
	<p>Filtrede birden çok eleman seçebilme özelliği ve renge göre filtreleme 
	2007de ve arama kutusu 2010'da geldi. Hala renge göre filtrelemeyi&nbsp; 
	kullanmayan arkadaşlarım olduğunu bilmek hem üzücü hem şaşırtıcı. </p>
	<p><img src="../../images/insertsortfilter1.jpg"></p>
	<h3>Advanced Filter</h3>
	<p>ilterli alanda fomrül olursa subtotal.</p>
	<p>slicer with table</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	</div>
--></asp:Content>
