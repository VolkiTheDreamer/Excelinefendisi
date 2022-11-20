<%@ Page Title='Formlar Kontroller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Formlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>
	<h1>Form Kontrolleri</h1>
    		<h2 class="baslik">Giriş</h2>
		<div class="konu">
		<p> Nesne yönelimli programlamanın en somutlaştığı yer olarak 
		ben şahsen kontrolleri görüyorum. Bunlar, gerçek dünya nesnelerine çok benziyorlar. Excelle 
		çalışırken, bir hücre, bir sayfave ya workbook'un kendisi Excel ile o 
		kadar bütündür ki onları içselleştirmişizdir, bu yüzden onları bir nesne gibi görmek bazen biraz zor 
		olabilir. Ama eminim bu sayfada nesne yönelimli programlama konusunu 
		iyice anlamış olacaksınız.</p>
		<p> Zira birçok progralama dilinde ve onlarla geliştirme yaptığımız 
		IDE'lerde olduğu gibi bu nesnelerin özelliklerini Properties 
		penceresinden değiştirebileceğiz ve bu deneyim de bizi programlama 
		dünyasına biraz daha yakın hissettirecektir. Gerçi kontrollerin propertylerine 
		sadece properties penceresinden(DesginTime) değil kodların çalışması 
		sırasında da (Runtime) erişebileceğiz. Ve yine gerçi Nesne Yönelimli 
		olmak demek, sadece özelliklerin Properties'ten değiştirilebilmesi demek 
		değildir, bundan çok daha büyük bir kavramdır ancak yeni başlayanlar 
		için kolaylık sağladığını düşünebiliriz.</p>
		<p> Bu sayfada temel olarak baz alacağımız örnek dosyaya
		<a href="../../Ornek_dosyalar/Makrolar/useformlar.xlsm">buradan</a> 
		ulaşabilirsiniz.</p>
            </div>

		<h2 class="baslik">Kontrol Tipleri</h2>
		<div class="konu">
		<p >Excel'de 2 tür kontrol bulunmaktadır.</p>
		<ol>
			<li>
			<p ><strong>Form kontrolleri</strong>: Worksheetler üzerine konan ve sınırlı 
			fonksiyonaliteye sahip kontroller.</p>
			</li>
			<li>
			<p ><strong>ActiveX kontrolleri</strong>:Daha gelişmiş 
			fonksiyonaliteye sahip olan, hem Worksheetler hem de UserFormlar 
			üzerine konan kontrollerdir.</p>
			</li>
		</ol>
		<p> Niye 2 tür kontrol grubu var diye soracak olursanız, önceleri sadece Form kontrolleri vardı, 
		sonra ActiveX kontrolleri geldi diye cevaplanabilir.</p>
		<p> Aşağıda Developer menüsünden ikisinin de içeriğini görebilirsiniz. 
		Birbirine çok benzeyen bu kontrollerin temel bazı farkları bulunuyor. 
		Bunlara aşağıda değiniyorum.</p>
		<p> <img src="../../images/vbaformkontroldev.jpg"></p>
		<h3>Worksheet/Form Kontrolleri</h3>
		<p> Bunlar, Excel arayüzünde <strong>Developer </strong>menüsü altında bulunurlar. 
		Bunların VBA olmadan genel kullanımlarını burada ele almayacağız. 
		Bu detayları <a href="../Excel/DeveloperMenusu_Kontroller.aspx">şu sayfada</a> bulabilirsiniz. VBA'siz 
		de kullanılan bu kontroller oldukça faydalı 
		kontrollerdir ve özellikle dashboard tarzı çalışmaların yaratımında oldukça kullanışlıdırlar.</p>
		<p> Bunların VBA'li kullanımında ise 
		ana olay(event) için makro oluşturulur. Mesela sayfa üzerine bir Button(düğme) yerleştirip o düğmenin 
		Click eventinin tetiklenmesiyle(özetle ona tıklayarak) başka bir makroyu 
		çalıştırma amaçlı kullanabiliriz.</p>
		<p >Bunların VBA'li kullanımdaki tek avantajları Windows'ta 
		oluşturduğunuz bir dosyanın Apple Mac bir bilgisayardaki Excel'de de 
		çalışacak olmasıdır. Zira Mac işletim sistemi ActiveX kontrolleri 
		desteklemezken bunları destekler.</p>
		<p ><strong>NOT</strong>:Gariptir ki Excelin 5.0 versiyonundan beri kullanılamayan 
		TextBox kontrolü(ve ne olduğunu bilmediğim diğer 2 kontrol) pasif olarak 
		ilgili menüde hala görünmektedir.</p>
		<h5 >Makro atama</h5>
		<p>Bu kontrollere sağ tıklanıp <strong>Assign Macro&gt;New </strong>denince default event için kod 
		ekranı çıkar. Oraya da istediğiniz kodu yazarsınız.</p>
		<h5 >Metin değiştirme</h5>
		<p>Uygun olan kontroler için Sağ tıklanıp&nbsp;<strong>Edit Text</strong> denerek ilgili 
		kontrolün üzerinde görünen metin değiştirilebilir.</p>
		<h3>ActiveX Kontrolleri</h3>
		<p>ActiveX kontrolleri hem worksheetlerde hem de VBA UserForm'ları 
		üzerinde kullanılırlar. VBA fonksiyonalitesi olarak worksheet formlarına 
		göre çok daha üstündürler, ancak Excel fonksiyonalitesi olarak ise 
		worksheet form kontrolleri daha kullanışlıdır. O yüzden size tavsiyem 
		bunları <strong>sadece UserFormlar üzerinde kullanın</strong>, diğerlerini de Excelin bir 
		hücre grubuyla ilişkendirmek için VBA'siz şekilde kullanın.</p>
		<p>Bir düğmeyle bir makro çalıştırmak için de yine worksheet/form 
		kontrollerini kullanbilirsiniz demiştik. Başka neler yapabilirsiniz. Listbox/Combobox'tan seçilen 
		değere göre, seçim yapılır yapılmaz o seçime ait bir veritabanı 
		sorgulaması yapılabilir. Mesela ürün kodlarının olduğu bir Listbox'ta, 
		seçilen ürüne ait özellikler boş sayfaya yazdırılabilir, yeya ikinci bir 
		Listbox'ın 
		içeriği doldurulabilir, mesela alt kategorideki ürünlerle.</p>
		<p >Yukarıda belirttiğim gibi ActiveX kontrollerinin en büyük 
		dezavantajı Mac kullanan bir bilgisayara Windows'ta hazırlanmış bir dosya 
		göndermek olacaktır. Ancak amacımız, ilgili kontrollerin ana eventi dışında 
		bir eventi kullanmaksa o zaman başka çareniz yoktur, mecburen ActiveX 
		kontrolü kullanacaksınız. Mesela, CommandButtonun sadece click eventini kullanacaksanız 
		Worksheet Form kontrolü iş görür, keza 
		Listbox'ın change eventi yeterliyse yine Worksheet Form kontrolü iş 
		görür, ama MouseUp eventini kullanacaksanız ActiveX kullanmak 
		zorundasınız.</p>
		<p >Aşağıda, toolboxta default olarak bulunan tüm kontrollerin 
		listesini görebilirsiniz.</p>
		<p ><img src="../../images/vbauserform_2.jpg"></p>
		<p>
		Bunların önemli olanlarının detay özelliklerine 
		aşağıda yer vereceğim, diğerlerini sizin keşfetmeniz gerekiyor.</p>
		<p>Bilgisayarlarımızda, Excel ve diğer Microsoft programlarınca kullanılan 
		başka ActiveX kotrolleri de vardır. Bunları, ActiveX kontrollerinin 
		olduğu blokta, sağ 
		alttaki(aşağıda kırmızlı işaretli) buton ile görebilirsiniz ama 
		bunların çoğu worksheetlerde kullanılamaz.&nbsp; Zaten eklemeye 
		çalışsanız bile bi uyarı çıkacaktır. Hangilerinin kullanılabileceğine dair bir liste var mı, 
		açıkçası bilmiyorum. İlgisini çekenler kurcalayabilir.</p>
		<p><img src="../../images/vbauserformkontrol1.jpg"></p>
	<p> Bununla beraber bunların hepsi <strong>Userformlar </strong>üzerinde kullanılabilirler. Bunun 
	için herhangi bir kontrolün üzerine sağ tıklayıp Additional Controls'e tıklamak 
	yeterlidir.</p>
		<p> <img src="../../images/vbauserformkontrol2.jpg"></p>
		<h4 >Worksheet'te bir kontrole makro atama</h4>
		<p >Developer'dan Design Mod yapılıp sağ tıklanır. <strong>View Code
		</strong>denir. İlk başta default(temel) event gelir, istenen event 
		seçilerek kod yazılır.</p>
		<h4 >Worksheet'te metin değiştirme</h4>
		<p >Developer'dan Design Mod yapılıp sağ tıklanır. Properties'ten
		<strong>Caption </strong>veya <strong>Text </strong>özelliği değiştirilir. 
		Veya yine objeye sağ tıklanıp XXXObject&gt;Edit denilerek doğrudan 
		metin editlenir.</p>
		<h3 >Karşılaştırma</h3>
		<ul>
			<li>
			<p >Excel hücreleriyle etkileşim, Form kontrolleriyle 
			kolayca sağlanır, VBA'siz kullanılır.</p>
			</li>
			<li>
			<p >Temel event(button için Click, Listbox için Change) kullanıp Mac 
			bilgisayara gönderme ihtimalimiz varsa:Form kontrol</p>
			</li>
			<li>
			<p >Temel event dışındaki eventler için ActiveX kontrolleri 
			kullanılır</p>
			</li>
			<li>
			<p >VBA Userformlar üzerinde mecburen ActiveX kontrolleri 
			kullanılır</p>
			</li>
		</ul>
		<p >
		Sayfanızda 2 tür kontrol de var diyelim. Hangisinin ne tür olduğunu nasıl anlarsınız?
		Form kontrollerine sağ tıklayabilirken, ActiveX'lere sağ tıklanamaz, 
		bunlara sağ tıklamak için Design Mod'da olmalısınız. Diyelim ki o sırada 
		Design Moddasınız, 
		bu durumda nasıl anlaşılır? Sağ 
		tıklayınca formül çubuğunda EMBED(...) diye bi formül çıkıyorsa ActiveX'tir, çıkmıyorsa 
		Form 
		kontrolüdür. Aynı zamanda ActiveX'e sağ tıklayınca <strong>Properties</strong> ve 
		<strong>View 
		Code</strong> çıkarken 
		diğerinde bunun yerine&nbsp;<strong>Assign Macro </strong>çıkar.</p>
"<p>		<img alt=" "  src="../../images/vbauserform1.jpg"></p>
			<h3>		Kontrollerin sayfa davranışını yönetmek</h3>
			<p>		Gerek form kontrollerinin gerek ActiveX kontrollerinin sayfa 
			üzerindeki konumu, görünürlüğü, aktif/pasifliği gibi özelliklerini 
			yönetmek için Shape ve OleObject kavramlarını incelemek gerekiyor. 
			Bu bilgiler, kavramsal olarak buraya uygun olmayıp, onları
			<a href="Ileriseviyekonular_ShapesveOleObjects.aspx">şu sayfada</a> 
			inceleyeceğiz.</p>
		</div>
		<h2 class="baslik">Temel Kontroller</h2>
		<div class="konu">
		<h3>Command Button</h3>
		<p>Kontroller arasında en sık kullanılanı ve en aşina olunanı CommandButton'dur.</p>
		<p>CommandButon'un default event'i <strong>Click</strong> olmakla birlikte başka 
		eventleri de vardır. Her zamanki yaklaşımımla ben bununla ilgili diğer 
		eventleri şimdiye kadar kullanmadığım için burda da örneklerini 
		vermeyeceğim. Arzu eden ve ihtiyaç duyan araştırabilir.</p>
		<p>Click event'i ile bir başka makro çalıştırılabileceği gibi, ekrana 
		bir FileDialog penceresinin gelmesi de sağlanabilir. FileDialog detayına
		<a href="DortTemelNesne_Application.aspx#filefolder">buradan</a> ulaşabilrsiniz. 
		Diğer button kullanım amaçları şöyle sıralanabilir:</p>
		<ul>
			<li>MsgBox ile bilgi gösterme</li>
			<li>InputBox ile kullanıcıdan bilgi girmesi/alan seçmesi isteme</li>
			<li>Hücreden bir bilgi okuma</li>
			<li>Hücreye bir bilgi yazdırma</li>
			<li>Veritabanına bilgi yazdırma</li>
			<li>Veritabanından bilgi okuma</li>
			<li>Spin Butonun değerini artırıp/azaltma</li>
			<li>Çeşitli değerleri/nesnelerin içeriklerini resetleme</li>
			<li>v.s</li>
		</ul>
		<h3>TextBox ve Label</h3>
		<h4>Label</h4>
		<p>Label en basit kontroldür. Üzerine genelde ya bir açıklama ya da bir 
		işin sonucunda sonuç mesajı yazdırırız.</p>
		<h4>TextBox</h4>
		<p>Kullanıcıdan birşeyler girmesini beklediğimiz kutulardır. Girilen 
		değerin ne olduğunu <strong>Text </strong>ve <strong>Value </strong>özellikleri 
		ile elde ederiz. TextBox'larda bu iki özellik genelde aynı değeri verir. 
		Text ve Value farkını aşağıda daha detaylı göreceğiz.</p>
		<p><span class="keywordler">ControlSource</span>:Kutuya, bir hücreden değer ataması yapmak istiyorsak bu 
		özelliği kullanırız. UserFormlarda pek kullanılmaz.</p>
		<p><span class="keywordler">Multiline</span>:Kutumuz, birden çok satır içerecekse bu 
		özelliğe True atarız.</p>
			<p><span class="keywordler">EnterKeyBehaviour</span>: Buna True 
			atandığı zaman Enter tuşu ile bir alt satıra geçebilirsiniz. False 
			durumundayken alt satıra geçmek için Ctrl+Enter kombinasyonunu 
			kullanmanız gerekir. Tabi alt satıra geçmesi için Multiline 
			özelliğine True atanmış olmasını söylemeye gerek yok sanırım.</p>
		<h3 >OptionButton'ları, Check Box'lar ve Çerçeveler</h3>
		<h4>OptionButton ve CheckBoxlar</h4>
		<p >Option buttonları kullanıcıya <strong>birden çok seçenek 
		içinden <span style="text-decoration: underline">birini </span></strong>
		seçtirmek için kullanılır. Checkboxlar ise <strong>birden çok seçenek içinden
		<span style="text-decoration: underline">çoklu </span></strong>seçim 
		yapmaya imkan sağlar. İkisinde de seçenek sayısının az olması tercih 
		sebebedir, çok seçenek olacağı zaman ListBox veya ComboBox kullanılması 
		önerilir.</p>
		<h4 >Çerçeveler</h4>
		<p >Çerçeveler, genelde Option butonları ve CheckBox'ları 
		gruplamak için kullanılmakla birlikte, ortak özelliği olan bütün 
		kontrolleri gruplamakta kullanılabilir. Bunlar .Net'taki GroupBox'larla 
		aynı işlevi görürler.</p>
		<p ><img src="../../images/vbaformsframe.jpg"></p>
		<p >Gruplamanın amacı sadece estetik ve anlamsal bir bütünlük 
		katmak değil, aynı zamanda çerçeve içindeki tüm kontrolleri tek seferde 
		enable/disable veya visible/invisible yapmak için de oldukça 
		kullanışlıdır.</p>
		<h4>Frame alternatifi</h4>
		<p >Bir grup OptionButton/CheckBox yaratmanın alternatifi de 
		bu kontrollerin <strong>GroupName</strong> özelliğine ortak bi değer atamaktır. Bu şekilde 
		kullanıldığında biri seçiliyken öbürleri seçimsiz olurlar. Başkaları 
		önerse de ben bu şekilde bir gruplamayı tercih etmiyorum. Zira yukarda 
		belirttiğim gibi gruplamanın amacı kontrolleri sadece aynı çatı altında 
		toplamak değil, tek seferde visible/enable özelliklerini de kontrol 
		etmektir.</p>
		<p >GroupName'i önerenler tarafından öne sürülen avantajlarını 
		ve benim yorumlarımı şöyle sayabiliriz.</p>
		<div ms.cmpgrp="content body">
			<div id="oaContent">
				<div id="api-doc-contents">
					<ul>
						<li>
						<p>Fazladan bir kontrol koymayarak kodun performansını 
						artırırsınız(Ben bunun ihmal edilebileceğini düşünüyorum)</p>
						</li>
						<li>
						<p>Frame içindeki tüm kontrollerin frame içine 
						sığdırılması zorunludur, bu da sıkışık bir görüntüye 
						neden olabilir. 
						GroupName kullanımında ise kontroller formun istediğiniz yerinde 
						olabilir(Neden olsun ki, bi seçeneği formun sağ üstüne 
						diğerini sol üstüne koyacak değilsiniz ki!)</p>
						</li>
						<li>
						<p>Çerçeveli bir görüntü istemiyorsanız kullanışlıdır. 
						Framede ise transparanlığı bozmuş olursunuz.(Genelde 
						çerçeve sınırı olur, yani Frame tercih edilmelidir)</p>
						</li>
					</ul>
				</div>
			</div>
		</div>
		<h4 >Başlangıç ayarları ve seçimler</h4>
		<p>Bir CheckBox düşünün, ilk başta seçili değil. Bu checkbox 
		seçildiğinde konuyla ilgili diğer tüm kontrolleri içeren bir çerçeveyi 
		görünür hale getiriyor. İlk başta bu çerçevenin Properties'ten Visible 
		özelliğine False atarız ki bunlar ilk başta görünmesin. Şimdi, Formumuz 
		açıldığında Checkbox'ı seçtiğinizde onla ilgili diğer tüm kontrollerin 
		de visible olmasını, seçimi tekrar kaldırdığınızda ilgili çerçevenin de 
		tekrar gizlenmesini istiyoruz. Formumuz ve kodumuz aşağıdaki gibidir.</p>
		<p >
		<img src="../../images/vbausrform1.jpg"></p>
		<pre class="brush:vb">
Private Sub CheckBox3_Click()
    If CheckBox3.Value = True Then
        frAktifPasif.Visible = True
    Else
        frAktifPasif.Visible = False
    End If
End Sub	&nbsp;</pre>
		<p>
		Yani diyoruz ki, CheckBox3(Falan filan..... yazan) seçiliyorsa 
		frAktifPasif çerçevesini(ve dolayıysla içindeki tüm kontroller) gizle, 
		seçili değilse göster.</p>
		<p>
		Tabi bunu yapmanın daha basit bir yolu var. Yazım şekli şu şekildedir.</p>
		<p>
		<strong><em>Kontrol.BooleanÖzellik=Not Kontrol.BooleanÖzellik</em></strong></p>
		<p>
		Yani diyoruz ki, kontrolün ilgili özelliğine zıttını ata. Boolean tipteki 
		tüm zıt değer atamalarında bu işlem yapılabilir.</p>
		<pre class="brush:vb">
Private Sub CheckBox3_Click()
     frAktifPasif.Visible = Not frAktifPasif.Visible
End Sub	&nbsp;</pre>
		<p>
		Bu şekilde yukarıda bahsettiğimiz gibi tek seferde tüm frame içindeki 
		kontrolleri yönettik. Frame yerine GroupName özelliğini kullansaydık, bunları tek 
		tek yapmak gerekecekti.</p>
		<h3 >Spin Button ve ScrollBar</h3>
		<p >Kullanıcının bir işlem yaparkan değerleri tek tek(5er 5er, 
		10ar 10ar v.s) artırma/azaltma gibi denemeler yapması sözkonusuya bu 
		kontrolleri kullanırız. Bunlar, kullanıcıdan Textboxa veya bir hücreye her seferinde bir 
		fazla/eksik değer girmesini beklemenin daha pratik bir yöntemini bize sunar.</p>
		<p>Genelde tek başına kullanımları yoktur. VBA tarafında bi Textbox 
		içindeki değeri veya bir değişkenin tuttuğu değeri belli miktarda 
		değiştirmek için kullanılırlar. Excel sayfasında ise, bir hücre içeriğini 
		değiştirmek için kullanılabileceği gibi Çeşitli Örnekler bölümünde 
		göreceğimiz gibi Excel 
		Filtre değerleri arasında dolaşmak için de kullanılabilirler. Biz şu an 
		VBA tarafına odaklanalım.</p>
		<p >Bu iki kontrolün de <span class="keywordler">Orientation<strong> </strong>
		</span>özelliğine Vertical/Horizontal 
		değerlerini atayarak yatay mı dikey mi duracağını belirleyebilirsiniz.</p>
		<p ><strong>Min, Max, SmallChange</strong> ikisinde de ortak olup açıklamaları 
		şöyledir:</p>
		<p ><span class="keywordler">Min</span>: Kontrolün alacağı en küçük değerdir, negatif 
		olabilir.</p>
		<p ><span class="keywordler">Max</span>: Kontrolün alacağı en büyük değerdir, negatif 
		olabilir.</p>
		<p ><span class="keywordler">SmallChange</span>:Oklara tıklandığında olacak değişim miktarını 
		gösterir.</p>
		<p >Scrollda ise fazladan <span class="keywordler">LargeChange
		</span>var. Bunda Scrollbarın 
		ortasına tıklandığında kaçar kaçar değişeceğini belirtiriz. Normal 
		değişim miktarı 1 ise bunu 10 yapabilirsiniz mesela.</p>
		<p >Aşağıdaki örneğe bakalım,</p>
		<p ><img src="../../images/vbausrform2.gif"></p>
		<p >Üstteki scroll ve ortadaki spin için şu kodları 
		yazabiliriz. </p>
		<pre class="brush:vb">
Private Sub ScrollBar1_Change()
   lblSıra.Caption = ScrollBar1.Value * 2
End Sub

Private Sub SpinButton1_Change()
   txtNo.Value = SpinButton1.Value
End Sub</pre>
		<p>Ben bunların min/max özelliğini Properties'ten ayarladım. Tabi 
		istenirse Runtime sırasında da bunlar değiştirilebilir. Mesela bir ComboBox'tan 
		veya Textbox'tan değişimin kaçar kaçar yapılacağını kullanıcıya da 
		bırakabiliriz.</p>
		<pre class="brush:vb">Private Sub TextBox1_Change()
  If TextBox1.Value &gt; 100 Then
    MsgBox "1-100 arası değer girilmelidir"
    Exit Sub
  End If
  SpinButton1.SmallChange = TextBox1.Value
End Sub</pre>
		<p>Şimdi de en alttaki Spine bakalım. Onun için önce modülün başında bi 
		global değişken(Dictionary olacak) tanımlayıp, form yüklenir yüklenmez de içine 5 değer atıyorum.</p>
		<pre class="brush:vb">
Public dict As Object
Private Sub UserForm_Initialize()

	Set dict = CreateObject("Scripting.Dictionary")
	dict.Add 1, "Volkan"
	dict.Add 2, "Ayşe"
	dict.Add 3, "Elif"
	dict.Add 4, "Murat"
	dict.Add 5, "Hakan"

End Sub</pre>
		<p>Bu sefer değişimi 1er 1er yaptırıp(smallchange özelliği=1) 1-5 arasındaki kişileri 
		öğreniyorum.</p>
		<pre class="brush:vb">
Private Sub SpinButton2_Change()
	Label1.Caption = dict(SpinButton2.Value)
End Sub</pre>
		<h3 >TabStrip ve MultiPage</h3>
		<h4 >MultiPage</h4>
		<p>MultiPage'ler, bir veya daha çok Page nesnesini birarada tutan 
		yapılardır. Framelerin bir üst modeli olarak düşünebilirsiniz. Bir alana 
		sadece 1 frame koyabilirken aynı alana birkaç Page'i olan bir MultiPage 
		koyabilirsiniz. Tek farkı yerden tasararuf değil aynı zamanda daha üst 
		seviyede bir gruplama imkanı da verir. Örneğin oluşturduğunuz form, 
		departmanınızdaki raporlara ulaşmayı sağlayan bir arayüz ise, 
		kullandığınız Multipage'in sayfalarından biri Kredi raporlarını diğeri 
		Mevduat raporlarını v.s gruplamış olabilir. Çeşitli Örnekler bölümünde 
		bununla ilgili bir <a href="#kokpit">çalışmamız</a> olacak.</p>
		<p>İlk başta bir Multipage içinde iki sayfa bulunur. Yeni sayfalar 
		eklemek için en üste sağ tıklayıp "Add Pages" diyin. Her sayfanın 
		içindeki kontroller, diğer sayfalardan tamamen bağımsızdır.</p>
		<p >Sayfalar 0 nolu indexten başlarlar. Bunlara index numarasıyla ulaşabileceğiniz gibi sayfa 
		ismi veya obje ismiyle de ulaşabilirsiniz.</p>
		<pre class="brush:vb">MultiPage1.Pages(0).Caption 'index
MultiPage1.Pages("Krediler").Caption 'sayfa ismi
MultiPage1.Page4.Caption 'obje ismi</pre>
<br>
			<pre class="brush:vb">Private Sub MultiPage1_Change()

MsgBox "sayfa indeksi:" &amp; MultiPage1.Value 'seçili sayfanını indexini
MsgBox "SelectedItem.Name yani obje adı:" &amp; MultiPage1.SelectedItem.Name 'seçili sayfanın obje adını
MsgBox "SelectedItem.Caption:" &amp; MultiPage1.SelectedItem.Caption 'seçili sayfanın adını

End Sub</pre>
		<h4>TabStrip</h4>
		<p>TabStrip kontrolü görünüm olarak MultiPage'e çok benzemekle birlikte, 
		bunun içine koyduğumuz kontroller tüm sayfalarda aynen görünür, yani 
		MultiPage'de olduğu gibi farklı sayfalarda farklı kontroller 
		bulunmayabilir. <strong>Ancak burda kritik olan, kontrollerin içeriğinin 
		farklı olmasını sağlıyor olmamızdır</strong>. Bunu da Tab 
		değiştikçe(bunu bir eventle yönetiriz) içeriğin değişmesini sağlayacak 
		bir kodla sağlarız. MultiPage'de ise Event olmasına gerek yok, zaten her sayfa 
		birbirinden bağımsız içeriğe sahiptir. </p>
		<p>Tablara erişim şekli MultiPage'de Page'lere erişim ile aynıdır.
		<a href="http://bigdon-in-vbaland.blogspot.com.tr/2014/01/tabstrip-vs-multipage.html">
		Bu sayfada</a> iki kontrol arasındaki farkları daha detaylıca görebilirsiniz.</p>
		<p >Şimdi kendi <strong>TabStrip </strong>örneğimize geçebiliriz.</p>
		<p >Bu örnekte Listbox da var, bunun detayını daha aşağıda göreceğiz, 
		şimdilik ona takılmayın. Sadece 
		listeyi doldurduğumuzu bilin o kadar.</p>
		<pre class="brush:vb">Private Sub UserForm_Initialize()
'form yüklenir yüklenmez ilk sekme açılır ve 2.sayfadan(0+2) data yüklenir
  TabStrip1.Value = 0
  Call ListeDoldur(0)
End Sub</pre>
<br>
		<pre class="brush:vb">Private Sub TabStrip1_Change()
'her sayfa değişiminde ilgili sayfaya ait ürünler doldurulur
  Dim ws As Worksheet
  Dim alan As Range
  Dim i As Integer

  i = TabStrip1.Value 'tab'larda index 0'dan başlar
  Call ListeDoldur(i)
End Sub</pre>
<br>
		<pre class="brush:vb">Sub ListeDoldur(k As Integer)
  Set ws = ActiveWorkbook.Worksheets(k + 2) 'worksheetlerde index 1den başlar
  Label1.Caption = ws.Name
  Set alan = ws.Range("A1").CurrentRegion 'ilgili sayfad

  ListBox1.Clear 'önce listeyi boşaltalım
  For Each urun In alan
     ListBox1.AddItem (urun)
  Next urun
End Sub</pre>
</div>
		<h2 class="baslik">Liste Kontrolleri</h2>
		<div class="konu">
		<p ><strong>Combobox</strong> ve <strong>Listboxlara</strong> 
		önceden belirlenmiş değerler atanabileceği gibi form üzerindeki düğmeler 
		aracılığıyla, bunların içeriği zengileştirilebilir veya içlerindeki 
		elemanlar silinebilir.</p>
		<h3 >Karşılaştırma</h3>
		<ul>
			<li>
			<p >Comboboxlar kullanıcıya tek değer gösterirken 
			Listboxlar tüm değerleri tek seferde gösterebilir(hepsi sığmazsa 
			scrollbar çıkar). Eğer amacınız tüm değerleri tek seferde göstermek 
			değilse yerden tasarruf amacıyla Combobox tercih edebilirsiniz. </p>
			</li>
			<li>
			<p >Diğer önemli fark ise Comboboxtan sadece 1 eleman 
			seçebilirken Listboxtan ise çoklu eleman seçimi yapabilirsiniz.</p>
			</li>
			<li>
			<p >Bir diğer fark ise, Comboboxların, listede olmayan bir 
			değeri girmeye izin vermesidir. Listboxta bu mümkün değildir.</p>
			</li>
		</ul>
		<p>
		Aşağıdaki görselde, bu iki kontrolü görebilirsiniz. Combobox açılmış 
		durumdadır.</p>
		<p ><img src="../../images/vbalistbbox1.jpg"></p>
		<p >Şimdi de bu kontrollerin çeşitli üyelerine(özellik, metod ve olay) bakalım. 
		Öncelikle şunu söyleyeyim. Bazı özelliklere hem kod yazarken hem 
		properties penceresinden, bazısına ise sadece kod yazarken erişilebilir. 
		"Properteis'ten 
		erişebiliyorken neden kod ile uğraşayım ki?" diye düşünebilirsiniz. 
		Bunun bir cevabı "Eğer 
		form üzerinde birbiriyle aynı türde çok fazla kontrolünüz varsa, mesela 
		10 tane combobox gibi, herbirine tek tek değer atama yerine, döngüsel 
		şekilde tek seferde kod ile yapabilrisiniz." olabileceği gibi, diğer bir 
		cevap ise "runtime sırasında değer atama gerekliliğidir". Mesela bir 
		butona tıkladığınızda başka bir kontrolün <span class="keywordler">
		Enabled </span>özelliğine False değeri atamak gibi.</p>
		<h3 >Listeleri doldurma</h3>
		<h4 >Yöntem1</h4>
		<p >Liste doldurma yöntemlerinden en bilineni ve basit olanı, 
		Formun <strong>Initialize</strong> eventi içine dizi olarak eklemektir.</p>
		<pre class="brush:vb">Private Sub UserForm_Initialize()
  cbYetkiSeviye.List = Array("Müdürler", "Yöneticiler", "Yetkililer")
  Me.cbYetkiSeviye.ListIndex = 0 'ilk eleman seçilir. -1 ile seçili hiç bir eleman olmaz, son eleman için me.cbYetkiSeviye.ListCount - 1
End Sub</pre>
		<p >Diğer yöntemler arasında Exceldeki bir sayfadan okuma, text 
		doyasından okuma veya Access gibi bir veritabanından okuma olabilir. Bu 
		işlemleri yine Initialize içinde yapabileceğiniz gibi bir Button'a 
		tıklayarak da yapabilirsiniz, tabi pratikte genelde listeler form açıldığında, 
		yani Initialize sırasında, doldurulur. 
		Bunların hepsinde de ilgili kontrolün
		<span class="keywordler">AddItem</span> metodu kullanılır.</p>
		<h4 >Yöntem2&nbsp;</h4>
		<p >İkinci yöntem olarak bir text dosyadan okuma yapalım:</p>
		<pre class="brush:vb"> dosya = "C:\....\Ornek_dosyalar\Makrolar\userformlist.txt"
Open dosya For Input As 1

Do Until EOF(1)
  Line Input #1, Content
  Me.ListBox2.AddItem Content
Loop

Close #1</pre>
		<h4 >Yöntem3</h4>
		<p >Excel sayfadan okuma için aklınıza döngüler gelmiş 
		olabilir, ne var ki buna hiç gerek yok. İlgili alanı Properties'ten <span class="keywordler">RowSource</span> 
		özelliğine referans verebilirsiniz. Ör. Sheet1!A17:A19(Sayfa adını ve ! 
		işaretini belirterek) veya runtime 
		sırasında <strong>lbŞehirler.RowSource=Range("A17:A19").Address</strong> 
		diyebilirsiniz.</p>
		<p >Değerlerin bulunduğu alan sabit değil de değişkense bunun 
		için aşağıdaki gibi bir kod kullanabilirsiniz.</p>
		<pre class="brush:vb">Private Sub RefEdit1_Change()
  Me.ListBox2.RowSource ="Sheet1!A1:A"&amp; Sheet1.Cells(Rows.Count, "A").End(xlUp).Row
End Sub</pre>
		<p >Bu işlemi bir <strong>refEdit</strong> elemanına da yaptırabiliriz.</p>
		<pre class="brush:vb">Private Sub RefEdit1_Change()
  Me.ListBox2.RowSource = Me.RefEdit1.Value
End Sub</pre>
			<p><strong>NOT</strong>: RefEdit kontrolünü form modal açılmışken 
			kullanmalısınız, modeless açılmış formlarda sıkıntı yaşanmaktadır.</p>
		<h4 >Yöntem4</h4>
		<p >Access'ten okuma yapmak için ya DAO ya ADO tekniklerini 
		biliyor olmak gerekiyor. Bunlar için
		<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">bu 
		sayfaya</a> bakınız.</p>
		<p >Bu arada belirtmek isterim ki yeni eleman eklemelerini sona 
		yapacaksınızdır ama olur da başa veya arada bir yere yapmak isterseniz,&nbsp; 
		<strong>AddItem</strong>'ın ikinci parametresini kullanabilirsiniz. 
		<strong>Ve unutmayın ki liste 
		kontrollerindeki indexler 1'den değil 0'dan başlar.</strong></p>
		<pre class="brush:vb">Me.ListBox2.AddItem Content,0 'ilk sıraya ekledi.</pre>
		<h4>AddItem detaylar</h4>
		<p>AddItem metodu ekleme işini, çok kolonlu bir listenin ilk kolonuna 
		yapar. Daha ileri kolonlara ekleme yapmak için List veya Column 
		propertylerini kullanabilrsiniz. Yine aynı propertiler kullanılarak aynı 
		anda birden fazla satır da ekleyebilirsiniz. Bu da şu anlama gelir: 
		Excel sayfasındaki bir grup hücreyi tek seferde ilgili liste kontrolüne 
		ekleyebilirsiniz.</p>
		<p>Bir diziyi olduğu gibi eklemek için List propertysi kulalnılırken, 
		trasnpose halini eklemek için Column propertysi kullanılır. ,Yani 
		dizi(i,j)'yi olduğu gibi eklemek için listbox.List(i,j)&nbsp; 
		kullanılabilir</p>
		<p>Aşağıda bununla ilgili bir örnek bulabilirsiniz, ki bu yöntem 
		elemanların bulunduğu alanı <strong>Rowsource </strong>olarak belirtmenin bir başka 
		yoludur. Ancak AddItem ile ilgili önemli bir detay da, liste kontrolümüz 
		bir datayla ilişkiliyse(Rowsource ile) Additem'ın çalışmayacağıdır. O 
		yüzden sonrasında dinamik şekilde yeni elemanlar eklemek istiyorsak 
		Rowsource ile değil aşağıdaki gibi ilerlemeliyz.</p>
		<pre class="brush:vb">Private Sub CommandButton10_Click()
   Me.lstÇiftKolon.List = Range("çiftkolon").Value
End Sub</pre>
		<h3 >Listeleri boşaltma</h3>
		<p>Listeleri <span class="keywordler">Clear</span> metodu ile boşaltırız. 
		Genelde dolu olan bir listeyi tekrardan doldurmadan önce boşaltmak iyi 
		bir fikirdir. Özellikle formu henüz kapatmamışken, işleri baştan almak 
		istediğinizde formun doldurma işlemi varsa, bu mükerrer doldurmaya neden 
		olacağı için kodumuza her zaman (boş bile olsa) önce listeyi boşaltarak 
		başlamak iyi bir pratiktir.&nbsp; <strong>AddItem</strong> gibi bu metod da, eğer 
		ki listemiz bir veri kümesine bağlıysa çalışmaz. Böyle bir durumda 
		öncelikle <strong>RowSource</strong> özelliğinin temizlenmesi gerekir.</p>
		<pre class="brush:vb">Sub listeboşalt()
   lbŞehirler.Rowsource=""
   lbŞehirler.Clear
End Sub</pre>
		<p>Eğer ki sadece belirli bir elemanı çıkarmak istiyorsak
		<span class="keywordler">RemoveItem</span> metodunu kullanırız. 
		Parametre olarak kaçıncı elemanın çıkartılacağı verilir.(Index'in 0'dan 
		başladığını unutmayın).</p>
		<h3 >Liste öğelerine erişim</h3>
		<p >Daha önceki kontrollerde <strong>Text </strong>ve <strong>
		Value </strong>özelliklerinden 
		bahsetmiştik. Bunlar diğer kontrollerde neredeyse her zaman eşittirler, 
		ancak liste kontrollerinde farklı olma durumları oldukça rastlanan 
		durumlardır.</p>
		<p >Liste kontrollerinde <span class="keywordler">Text</span>, 
		sizin gördüğünüz değeri verirken, <span class="keywordler">Value</span> 
		altta yatan değeri verir. Örneğin listeyi bir veritabanındaki 2 kolonlu 
		bir tablodan(veya excelde 2 kolonlu bi alandan) doldurmuşsunuz diyelim. 
		İlk 
		kolon şehir ismi ikinci kolon şehir kodudur. <strong>ColumnCount</strong> 
		özelliğine 1 derseniz, sadece bir kolon gösterilecektir. İlk kolon 
		listbox içinde gösterilecektir, ancak listboxtan şehir seçimi yapıldığnda 
		<strong>Value </strong>değerine şehir adı değil de kodu atansın istiyorsak <strong>
		BoundColumn</strong> özelliğine 2 atarız. <strong>TextColumn</strong> 
		özelliğine ise 1 atarız. Çok kolonlu listelerde TextColumn'un genelde 1 
		yapıldığı görülür ancak pratikte bunun farklı olduğu durumlar olabilir. 
		Örneğin aşağıdaki örnekte listeye ülkeler yüklenir, Value olarak id 
		tutulur, Text olarak da başkentler tutulabilir.</p>
		<p ><img src="../../images/vbalistbbox2.jpg"></p>
		<p >Bunun için yapılması gerekenler:</p>
		<p ><strong>ColumnCount</strong>:1 (Buradaki 1, <strong>kaç</strong> 
		kolon gösterilecek anlamımnda)<strong><br>BoundColumn</strong>:2 (Buradaki 2, <strong>kaçıncı</strong> 
		kolon Value<strong> </strong>olacak<strong><br>TextColumn</strong>:3 (<strong>kaçıncı</strong> 
		kolon text değerini tutacak, yani gösterilecek)</p>
		<p >Liste kontrollerinde eğer büyük veritabanlarıyla çalışıyorsak 
		peformans açısından <strong>Value </strong>özelliği ile işlerimizi halletmeliyiz.</p>
		<p ><strong>NOT</strong>: TextColumn'a&nbsp;-1 atandıysa(default budur), 
		Text özelliği seçilen değerin görünen değerini, 0 verilmişse elemanın indexini, 0'dan büyükler 
		için kaç verilmişse o kolondaki değeri verir. Yani ilk kolon için 
		TextValue=1, ikinci kolon için TextValue=2 v.s</p>
		<h4 >Eleman erişimi</h4>
		<p >Peki hangi elemana erişeceğimzi nasıl belirliyoruz?
		<span class="keywordler">List</span> propertysi ve elemanın index 
		numarası ile. </p>
		<pre class="brush:vb" style="width: 499px">listbox1.List(0) 'ilk eleman
</pre>
		<p >Normalde List property'si iki eleman alır: satır ve sütun. 
		İkinci eleman belirtilmezse ilk kolon baz alınır. Yani yukardaki kod ile listbox1.List(0,0) 
		özdeştir. (Not:List propertysinin parametreleri 0'dan başlar, 1'den 
		değil)</p>
		<p >Peki indexi bilmiyorsak, yani dinamik bir şekilde ele almamız 
		gerekiyorsa, onun da yolu var. Aşağıdaki örneğe bakalım:</p>
		<p ><span class="keywordler">ListIndex ve List</span> 
		özellikleri<strong>: </strong>O an seçili elemana erişim için bu iki 
		özellikle kombine bir şekilde kullanılır. <strong>lbYıl.List(lbYıl.ListIndex)</strong>.</p>
		<p >ListIndex bize o an seçili elemanın indexini verirken, bu 
		indexi List propertry'sine parametre gönderince seçili elemanın 
		görünen değerini bize verir. ListIndex 0'dan başlar. (Yukarıda 
		bahsettiğimiz gibi TextColumn özelliğine 0 atayarak da indeksi elde 
		edebiliyoruz)</p>
		<p>Aşağıda 3 ayrı değer 
		erişim yöntemi bulunuyor. Farkları inceleyerek anlamaya çalışın. Örnek 
		olarak Japonya seçilyse;</p>
		<pre class="brush:vb">Private Sub CommandButton4_Click()
   MsgBox "Value:" &amp; lstBağımlı.Value '200
   MsgBox "Text:" &amp; lstBağımlı.Text 'Tokyo
   MsgBox "List&amp;listindex:" &amp; lstBağımlı.List(lstBağımlı.ListIndex) 'Japonya
End Sub</pre>
		<h4 >Çok kolona erişim</h4>
		<p >Çok kolona erişmeyi yine <span class="keywordler">List</span> 
		özelliği ile yapıyoruz. Bu yöntemi sadece erişim için değil, veri ekleme için de 
		kullanabilirsiniz.</p>
		<pre class="brush:vb">lstSozluk.AddItem "iyi"
lstSozluk.List(0,1)="good" 'ikinci kolona
lstSozluk.List(0,2)="gut" 'üçüncü kolona</pre>
			<p>Çok kolonlu bir listeye yeni eleman eklemek de şöyle olur</p>
			<pre class="brush:vb">Private Sub CommandButton1_Click()
   Me.lst1.AddItem "kötü"
   Me.lst1.List(lst1.ListCount - 1, 1) = "bad"
   Me.lst1.List(lst1.ListCount - 1, 2) = "schlecht"
End Sub</pre>
		<h4 >Listbox'ta çoklu seçim:MultiSelect özelliği</h4>
		<p >MultiSelect özelliğinin alabileceği 3 değer vardır.</p>
		<ul>
			<li>
			<p ><strong>fmMultiSelectSingle</strong> (numerik değeri 0):Tekli 
			seçim. Her elemana tıklayışta sadece o seçilir.</p>
			</li>
			<li>
			<p ><strong>fmMultiSelectMulti</strong> (numerik değeri 1):Her 
			tıklamada, tıklanan eleman seçili kalır, tekrar aynı elemana 
			tıklanırsa seçim kalkar.</p>
			</li>
			<li>
			<p ><strong>fmMultiSelectExtended</strong> (numerik değeri 2):İki 
			eleman arasındaki tüm elemanları tek seferde seçmek için SHIFT 
			tuşuna basılır. CTRL tuşu ile ise fmMultiSelectMulti modu taklit 
			edilebilir.</p>
			</li>
		</ul>
		<p>
		Çoklu seçimde hangi elemanların seçili olduğunu <span class="keywordler">
		Selected</span> özelliği ile test edebiliriz. Parametre olarak elemanın 
		indexini alır: <strong>Listbox1.Selected(n)</strong></p>
		<p>
		Mesela aşağıdaki kod ile sadece seçili elemanları bir Collection'a 
		atıyoruz.</p>
		<pre class="brush:vb">
Private Sub CommandButton5_Click()
Dim coll As New Collection
For i = 0 To ListBox3.ListCount - 1
    If ListBox3.Selected(i) Then
        coll.Add ListBox3.List(i)
    End If
Next i
MsgBox "collectionda " & coll.Count & " adet eleman var"
End Sub</pre>
			<h5>
			Listedeki elemanları bir collection'a atama</h5>
			<p>
			Yukarıdaki işlemi bir de fonksiyon haline getirirsek bundan sonra ne zaman 
			bir listboxtan seçili elemanları almamız gerekse bu fonksiyonu 
			kullanabiliriz.</p>
			<pre class="brush:vb">
Function ListBoxtakiSeçiliElemanlarıSeç(lst As MSForms.ListBox) As Collection
Dim col As New Collection

If lst.List(lst.ListIndex) = -1 Then GoTo atla

For i = 0 To lst.ListCount - 1
    If lst.Selected(i) = True Then col.Add lst.List(i)
Next i

atla:
Set ListBoxtakiSeçiliElemanlarıSeç = col
End Function

'Kullanımı
Sub testListBox()
Dim col As Collection 'new yok, fonksiyonla dolduracağız

Set col = ListBoxtakiSeçiliElemanlarıSeç(UserForm1.LitBox1)
End Sub</pre>
		<h4>
		Listede belirli bir elemanı seçmek(işaretlemek)</h4>
		<p>
		Şimdiye kadar elemana erişim ile hep onun değerini elde etmeyi 
		kastettik. Ancak bazen ilgili elemanı seçmek de isteyebiliriz. Bu işlem 
		genelde, listedeki ilk elemanı seçmek için yapılır, ancak tabiki 
		herhangi bir eleman seçiminde de kullanılabilir.</p>
		<p>
		Bunun için iki yöntem var:</p>
		<pre class="brush:vb">
Private Sub CommandButton5_Click()
    ListBox3.Selected(0)=True 'Çoklu seçim modunda işe yaramaz
    'veya
    ListBox3.ListIndex=0
End Sub</pre>
		<h4>Kolon gizleme</h4>
		<p>3 kolonlu bir veri kümemiz olsun. Diyelim ki üçünü değil de baştaki ile 
		sondakini almak istiyorsunuz. Böyle bir durumda üçünü de RowSource'a alırız, 
		ancak ortadakini gizleriz. 
		Gizlemek için <span class="keywordler">ColumnWidths</span> özelliğine 0 atarız. 
		Ancak ColumWidths özelliği kullanılırken malesef tek bir kolona değer 
		ataması yapılamıyor, üç kolon için de değer girmek 
		lazım.</p>
		<pre class="brush:vb">listbox1.ColumnWidths="50;0;50"</pre>
			<h4>ListBox'ta dinamik filtreleme</h4>
			<p>Filtreleme amacı gören bir textbox'a yazacağınız metinlerle bir 
			listbox'taki elemanlarda dinamik filtreleme yapabilirsiniz. Bunun 
			için yol haritası şöyledir:</p>
			<ul>
				<li>Global bir Collection oluşturun</li>
				<li>Bu collection'ı ve ilgili listbox'ı aynı elemanlarla formun 
				başlangıcında doldurun</li>
				<li>Textbox'ın Change eventine de ilgili filtreleme kodunu yazın</li>
			</ul>
			<p>Kodlar aşağıdaki gibi olabilir:</p>
			<pre class="brush:vb">
'Global değişken
Dim ülkelerCol As New Collection

'Form başlangıcı
Private Sub UserForm_Initialize()

	For Each ülke In Range("ülkeler")
	   ülkelerCol.Add ülke.Value
	   Me.lstDinamik.AddItem ülke.Value
	Next ülke

End Sub

'TextBox change eventi
Private Sub txtFiltre_Change()
Dim filtreliÜlkeler As New Collection
    Me.lstDinamik.Clear 'önce boşaltıyoruz ki mükerrerlik olmasın

    For Each ü In ülkelerCol
        If InStr(1, ü, txtFiltre.Text, vbTextCompare) > 0 Then filtreliÜlkeler.Add ü
    Next ü

    For Each ü In filtreliÜlkeler
        Me.lstDinamik.AddItem ü
    Next ü
End Sub
</pre>
		</div>
		<h2 class="baslik">Diğer detaylar</h2>
		<div class="konu">
		<h3> Value, Text, Name, Caption</h3>
		<p> Yukarıda bahsettiğimiz konulara biraz daha detaylı bakalım.</p>
		<p><span class="keywordler">Text</span>: Ekranda gördüğümüz metni verir.</p>
		<p><span class="keywordler">Value</span>: Arkaplanda tutulan değeri 
		verir.</p>
		<p>
		Bu iki özellik genelde aynı değeri verir. Şu istisnalar hariç:</p>
		<ul>
			<li>Sözkonusu kontrol bir listbox veya combobox ise</li>
			<li>Gösterilen değer bound column'dan farklı ise</li>
		</ul>
		<h4>Value detaylar</h4>
		<ul>
			<li>Multiselect moddaki listboxta Value kullanılamaz</li>
			<li>Multicolumn listboxta BoundColumn varsa value değeri seçili 
			satırdaki bu kolondaki değeri verir</li>
			<li>Multipage'de sayfa indexini verir</li>
			<li>Checkbox OptionButton ve ToggleButtonda ilgili kontrolün seçili 
			olup olmadığını verir. Seçiliyse True, aksi halde False</li>
			<li>Spin ve ScrolBarda o anki değeri verir</li>
			<li>TextBox'ta Text ile aynı değeri verir.</li>
		</ul>
		<p> <span class="keywordler">Caption</span>: Label'da yazan metni, 
		Form'da ise form başlığını verir. Gariptir ki, Label'da Text veya Value 
		özelliği yerine Caption konmuş.</p>
		<p> <span class="keywordler">Name</span>: Nesnenin adını verir. Kod sırasında 
		bu nesneye bu isimle 
		başvuru yapılabilir. Ör: Yılları gösteren comboboxa "cbYıllar" diye 
		çağırdığımız gibi. Bu özelliği <strong>If control.Name="cbYıllar"</strong> 
		şeklinde döngüsel bir kod içinde ilgili nesnenin belirli bir nesne olup 
		olmadığını kontrol etmek için de kullanabiliriz.</p>
		<h3>
		List özellikleri</h3>
		<h4>
		ListCount</h4>
		<p>
		Readonly olan bu özellik, ilgili liste kontrolündeki satır sayısını 
		verir. ListRows'daki Rows ifadesi biraz kafa karışıklığı yaratabilir ama 
		satır sayısını ListRows değil ListCount vermektedir. Bu özelliğe sadece 
		kod ortamında ulaşılabilir.</p>
		<h4>
		ListRows</h4>
		<p>
		Sadece comboboxlarda bulunan bu özellik, comboboxta gösterilecek eleman 
		sayısını verir. Default değeri 8'dir. Belirtilen değerden daha fazla 
		satır varsa kenarda scrollbar çıkar. Aşağıdaki kod ile dinamik bir 
		şekilde gösterilecek eleman sayısını kontrol edebilrsiniz. Eğer 
		comboboxtaki eleman sayısı 5ten büyükse 5le sınırlayalım, 5ten küçükse 
		kaç satırsa o kadar görünsün.</p>
		<pre class="brush:vb">
Private Sub UserForm_Initialize()

With ComboBox1
	If .ListCount > 5 Then
	  .ListRows = 5
	Else
	  .ListRows = .ListCount
	End If
End With

End Sub	</pre>
		<h3> Me</h3>
		<p> Üzerinde çalıştığınız formun kendisine <span class="keywordler">Me</span> ifadesi ile 
		başvurabilirsiniz. Bu ifade, sadece forma başvuru için faydalı değil aynı zamanda 
		form üzerindeki kontrollere intellisense yardımıyla hızlıca ulaşma 
		imkanı verdiği için de faydalıdır.</p>
		<h3 >Kontrolleri tek tek dolaşma</h3>
		<p>Bazı durumlarda formdaki tüm kontrollerde dolaşıp, onların 
		tipine(TypeNeme), adına(Name) veya başka bir özelliğine bakarak işlem 
		yapmak isteriz. Bunu <strong>Controls</strong> collection'ına For Each uygulayarak 
		yaparız.</p>
		<p>Aşağıdaki örnekte Label olan tüm kontrollerin adını yazdırıyoruz.</p>
		<pre class="brush:vb">For Each ctrl In Me.Controls
    If TypeName(ctrl) = "Label" Then
        Debug.Print ctrl.Name
    End If
Next ctrl</pre>
		<p>
		<span style="font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif; font-size: 1em">
		Dolaşmak istediğimiz kontroller belli bir çerçeve(Frame) içindeyse;</span></p>
		<pre class="brush:vb">For Each ctrl In Me.Frame1.Controls
    If TypeName(ctrl) = "Label" Then
        Debug.Print ctrl.Name
    End If
Next ctrl</pre>
		<p>Tüm framelerde dolaşmak için</p>
		<pre class="brush:vb">
For Each cf In Me.Controls
    If TypeName(cf) = "Frame" Then
	For Each ctrl In cf.Controls
	    If TypeName(ctrl) = "Label" Then
	        Debug.Print ctrl.Name
	    End If
	Next ctrl
    End If
Next cf</pre>
		<h3> Event detayları</h3>
		<h4> Mouse eventleri</h4>
		<p> <strong>MouseDown</strong>: Mouse tuşu basıldığında meydana gelir</p>
		<p> <strong>MouseUp</strong>: Mouse tuşu bıraklıdığında meydana gelir</p>
		<p> <strong>Click</strong>: Mouse ile tıklanabilir bir kontrole 
		tıklandığında meydana gelir</p>
		<p> Önce MouseDown olur, sonra MouseUp, en sonra Click. İlk ikisi hem 
		sol hem sağ tuş ile tetiklenebilirken Click sadece sol tuş ile 
		tetiklenir. Mesela bir kontrolün ucuna tıklayıp yeniden 
		boyutlandıracaksanız, tıkladığınız anda MouseDown gerçekleşir, yeniden 
		boyutlandırma bittiğinde ve mousetan elinizi çektiğinizde Up 
		gerçekleşir.</p>
		<p> <strong>MouseMove</strong>:Üzerinden geçerken gerçekleşir. Bunu çok kullanma durumum olmadı 
		açıkçası. İlgili 
		kontrolün üzerine gelindiğinde bir mesaj vermek istiyorsanız bunu 
		<strong>ControlTip</strong> özelliği ile de verebilrisiniz.</p>
		<p> Buton parametresiyle sol/sağ hangisine basıldığı tespit edilebilir. 
		Mouse tuşlarının nasıl öğrenileceğini aşağıda klavye tuşlarının olduğu 
		bölümde görebilirsiniz.</p>
		<p> <strong>X ve Y</strong> parametreliryle hangi noktalara basıldığı 
		tespit edilebilir, yine bunlar da çok kullandığım özellikler değiller.</p>
		<p> <strong>Shift</strong> parametresiyle Shift, Ctrl, Alt tuşlarıdan birine basılıp 
		basılmadığı kontrol edilebilir.</p>
		<ul>
			<li>1:SHIFT</li>
			<li>2:CTRL</li>
			<li>3:SHIFT+CTRL</li>
			<li>4:ALT</li>
			<li>5:ALT+SHIFT</li>
			<li>6:ALT+CTRL</li>
			<li>7:üçüne birden</li>
		</ul>
		<p> Aşağıda çeşitli örnekler bulunmakta.</p>
		<pre class="brush:vb">
Private Sub CommandButton9_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	MsgBox "mousedown-" & Button & "-" & Shift & "-" & X
End Sub
--------------------------
Private Sub CommandButton9_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
	MsgBox "mouseup-" & Button & "-" & Shift & "-" & X
End Sub		
-------------------------
Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = KeyCodeConstants.vbKeyRButton Then
    ListBox2.AddItem ListBox1.List(ListBox1.ListIndex)
End If
End Sub
--------------------------
Private Sub cbYıllar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cbYıllar.DropDown
End Sub		</pre>
		<h4> Klavye eventleri</h4>
		<p> 3 adet klavye eventi vardır. Bunlar <strong>KeyDown</strong>, 
		<strong>KeyPress</strong> ve <strong>KeyUp</strong> olup 
		bu sırayla meydana gelirler. KeyDown ve Keyup parametre olarak Keycode 
		alırken, KeyPress KeyAscii alır. </p>
		<p> Hangi tuş veya tuş kombinasyonlarına(Ctrl+Enter gibi) basıldığını 
		öğrenmek için kullanılırıllar.</p>
		<p> Mesela bazen yer tasarrufu yapmak amacıyla Textboxa yazılan metinle 
		ilgili bir iş yapmak için form üzerine button koymak yerine yazmayı 
		bitirdikten sonra Enter'a(veya Ctrl+Enter) basılması durumunda ilgili işlemin yapılmasını 
		sağlayabilirsiniz.</p>
		<pre class="brush:vb">
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Me.cbYetkiSeviye.AddItem Me.TextBox1.Text
        Me.Label1.Caption = "Yetki seviyelerine " & Me.TextBox1.Text & " eklendi"
    End If
End Sub				&nbsp;</pre>
		<p> KeyCodelari aşağıdaki linklerden bulabileceğiniz gibi, VBA'de
		<strong>KeyCodeConstants</strong> yazıp "."'ya basınca intellisense aracılığı ile 
		constant değerlerini de yazabilrsiniz.</p>
		<p> <img src="../../images/vbauserformnkeycode.jpg"></p>
		<p> <strong>Not:Enter</strong> için <strong>vbKeyReturn</strong> diye bakmak lazım, 
		vbKeyEnter diye bişey bulunmuyor.</p>
		<p> <a href="http://www.asciitable.com/">http://www.asciitable.com/</a></p>
		<p> 
		<a href="https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/keydown-keyup-events">
		https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/keydown-keyup-events</a></p>
		<p> 
		<a href="mailto:https://stackoverflow.com/questions/1367700/whats-the-difference-between-keydown-and-keypress-in-net">
		Bu linkte</a> ise .Net dilindeki farklar detaylıca anlatılıyor ama 
		bu açıklamaların prensipte VBA için de geçerli olduğunu söyleyebilirim.</p>
			<h4 id="multievent"> 
			Birçok kontrol için tek event tanımlama</h4>
			<p> 
			Formumuzda diyelim ki 10 adet textbox var, ve hepsi için de ortak 
			bir Event tanımlamak istiyorum. Mesela içine girince içindeki yazı 
			silinsin istiyorum. Bunun için tek tek herbirine event tanımlamak 
			zahmetli olacaktır. İşte böyle durumlar için custom eventlerden 
			yararlanıyoruz. Örnek dosyayı şuradan
			<a href="../../Ornek_dosyalar/Makrolar/userform_multitextbox.xlsm">
			indirebilirsiniz</a>.</p>
			<p> 
			&nbsp;Adımlarımız şöyle:</p>
			<p> 
			Öncelikle bir Class Modül yaratırız. Tepesine aşağıdaki kodu 
			yazarız. Biz burada TextBox için yazıyoruz ama farklı kontroller 
			için de aynısı uygulanabilir.</p>
			<pre class="brush:vb">
Public WithEvents txtGroup As MSForms.TextBox</pre>
			
			<p> Sonra tepeden nesne kutusunda txtGroup seçilir, yandan da 
			mousedown eventi seçilir(Custom TextBoxlarda Enter eventi 
			bulunmuyor, ama mousedown da aynı görevi görecektir. Tabi ilgili 
			kutulara mouse ile tıklanamsı kaydıyla, tab tuşuyla ilerlenerek 
			gelinirse tetiklenmez). İçine de aşağıdaki kod yazılır.</p>
			<pre class="brush:vb">
Private Sub txtGroup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    With txtGroup
       .Text = ""
       .ForeColor = vbBlack 'Form açıldığında gri renkli bişeyler yazıyor olsun         
    End With
End Sub</pre>
			<p>
			Son olarak Form modülüne gelip tepeye Class1 tipli bir dizi 
			tanımlıyoruz, eleman sayısnı bilmediğimiz için boyutsuz 
			tanımlıyoruz. Initialize eventi içinde TextBoxlarda dolaşarak 
			boyutumuzu sürekli artırıyoruz.</p>
			<pre class="brush:vb">
'Global değişken bölgesine
Dim controller() As New Class1

Private Sub UserForm_Initialize()
	Dim adet As Integer
	Dim ctrl As Control
	For Each ctrl In UserForm1.Controls
	    If TypeName(ctrl) = "TextBox" Then
	        adet = adet + 1
	        ReDim Preserve controller(1 To adet)
	        Set controller(adet).txtGroup = ctrl
	    End If
	Next ctrl			
End Sub</pre>
			<p>
			Kontrol sayısı çok ise boyutsuz dizi tanımlamak yerine Collection 
			tanımlamak daha doğru bir çözüm olacaktır. Bununla ilgili bir örnek 
			şu
			<a href="https://stackoverflow.com/questions/1083603/vba-using-withevents-on-userforms">
			linkte</a> bulunmaktadır.</p>
		<h4> 
		Diğer Eventler</h4>
		<p> 
		ListBox'ın Change eventi, Formun Initialize ve Terminate eventleri 
		adları üzerinde olan eventler, bunları kurcalayarak kendinizin 
		görmesinde fayda var. Mesela ListBox'ta bir ana ürün seçildiğinde onun 
		yanındaki listboxa alt ürünlerin gelmesini ilk listbox'ın change 
		eventiyle yaptırabilirsiniz. Form'un Terminate eventi ise Workbookların 
		Close eventine benzer, form kapanırken devreye girerler ve kapanış 
		işlemlerinizi yapmanzı sağlar.</p>
		<p> 
		Diğer birçok eventi şimdiye kadar hiç kullanmadım.</p>
		<p> <strong>NOT</strong>:Listbox'ta seçilen bir eleman Excel 
		sayfasındaki bir hücreyi değiştiriyorsa bu değişklik Worksheet'in Change 
		eventini tetiklemez.</p>
		<h3 >Hizalama ve Ölçü işlemleri</h3>
		<h4>Userformlarda</h4>
		<p>VBA editöründeyken araç çubuğuna sağ tıklayın ve UserForm çubuğunu 
		aktive edin.</p>
		<p><img src="../../images/vbalistbox3.jpg"></p>
		<p>Bizim 
		ilgileneceğimiz, kırmızı halka içindekilerdir.</p>
		<p><img src="../../images/vbalistbox4.jpg"></p>
		<p>Onların da içerikleri aşağdaki gibidir. Soldakiyle çeşitli yönlerde 
		hizalama yaparız. Ortadakiyle kontrollerin arasındaki uzaklığı eşit hale 
		getiririz. En sağdakiyle ise kontrollerin ölçülerini eşit hale getiririz. 
		Bunlarla oynayarak ne işe yaradıklarını daha kolay görebilirsiniz.</p>
		<p><img src="../../images/vbalistbox5.jpg"></p>
		<h4>Worksheet'te</h4>
		<p >
		İlgili kontrolün konumu hücreler üzerinde rasgele durmasın da, uçları 
		hücrelerin köşelerine gelsin istiyorsanız, ilgili kontrol 
		seçiliyken Format menüsünden Arrange Grubundaki Align butonuna tıklayın 
		ve açılır kutudan <strong>Snap to Grid</strong>(kılavuzlara dayandır) diyin, arkasından 
		ilgili kontrolün uçlarını köşelere doğru çekin, otomatikman yerleşir 
		(Siz bu son adımı da yapmadan köşelere otomatikman yerleşmez)</p>
		<h3 >
		WorkSheet'te ActiveX listbox'a hücre bağlama</h3>
		<p >
		Önce developer menüsünden desgin moda geçilir. Sonra <strong>
		ListFillRange</strong> özelliğine istenen hücre grubu seçilerek 
		aktarılır. Aşağıdaki örnekteki gibi.</p>
		<p style="margin: 0in">
		<img src="../../images/vbauserform1.jpg" ></p>
		<h3>Diğer özellikler</h3>
		<ul>
			<li><span class="keywordler">WordWrap</span>: Text veya 
			Caption özelliğine birden fazla satırda yazma özelliği verir.</li>
			<li><span class="keywordler">ControlTipText</span>: ilgili kontrolün 
			üzerine gelince onun hakkında kısa bilgi veren, veya birtakım 
			talimatlar içeren bir balon çıkar.</li>
			<li><span class="keywordler">Enabled</span>: İlgili elemanla etkileşime 
			geçilip geçilemeyeceğini belirtir. Genelde bir başka kontrolle diğer 
			kontrollerin enabled özelliği kontrol edilir.</li>
			<li><span class="keywordler">Visible</span>:Enabled'ın kullanım 
			mantığına benzer. Bu, etkileşimden ziyade ilgili kontrolü gösterir 
			veya gizler.</li>
			<li>
			<span class="keywordler">TabIndex</span>: Kontroller arasında Tab 
			tuşu ile gezinebilirsiniz. Hangi sırada gezineceğinizi bu özelliğe 
			atayacağınız değerle yönetirsiniz.</li>
			<li>
			<span class="keywordler">ControlSource</span>: Bir kontrolde 
			seçtiğiniz/girdiğiniz değerin Excelde bir hücreye de yansımasını 
			istiyorsanız bu özelliğe o hücreyi atarsınız. Ör:Listboxtan 
			seçtiğiniz bir şube adı A1 hücresinde de çıksın isterseniz 
			ControlSource özelliğine A1 atayın. Genelde properties'ten 
			designtime sırasında kullanılır.</li>
		</ul>
		<p>
		Aşağıda <a href="http://www.globaliconnect.com">http://www.globaliconnect.com</a> 
		sitesinden aldığım bir kontrol-özellik matirisi var. Bu matristen, hangi 
		kontrolün hangi özellikleri mevcut, onları tek bakışta görebilirsiniz.</p>
		<p><img src="../../images/vbausrform3.jpg"></p>
		<h3> Çeşitli püf noktaları</h3>
		<p> UserForm kontrollerini kullanırken bazı püf noktalarını bilmek 
		oldukça faydalı olabilemktedir. Bunlardan birkaçını aşağıda vermeye 
		çalıştım.</p>
		<ul>
			<li><strong>Toggle işlemi</strong>: Bir kontrole tıklandığında 
			onunla ilgili bir boolean işlem yapılacaksa( başka bir kontrolün 
			enabled değerini, kendisinin durumuna veya zıttına ayarlamak gibi) 
			bunu If bloğu içinde yapmak yerine ters/aynı boolena değer atanarak 
			tek satırda yapabilirsiniz.<pre class="brush:vb">
If Checkbox1.Value= True Then 
   Frame1.Enabled=True
Else
   Frame1.Enabled=False  
End If		
'yerine
Frame1.Enabled=Checkbox1.Value
'ters işlem yapılacaksa başına Not ifadesi konur
Frame1.Enabled= Not Checkbox1.Value</pre>
			
			</li>
			<li>Değer girilmesi gereken yerler için kontrolünüz olsun. Ör: 
			Mail gönderim işlemi yapan bir Formunuz varsa, Subject(Konu) alanı 
			mutlaka dolu olmalı. <br>
			
			<pre class="brush:vb">
If konu.Text ="" Then 
	MsgBox "Lütfen konu alanını boş bırakmayın"
	Exit Sub
End If		</pre>
			
			</li>
			<li>Aşağıdaki linklerde hem genel olarak önemli noktalara temas var 
			hem de çeşitli püf noktaları da bulunuyor. Bunları da ayrıca 
			incelemenizi tavsiye ederim.</li>
		</ul>
	<p style="margin-left: 40px"> 
	<a href="https://support.microsoft.com/en-us/help/829070/how-to-use-visual-basic-for-applications-vba-to-change-userforms-in-ex">
	Microsoft userform dökümantasyonu</a></p>
			<p style="margin-left: 40px"> 
	&nbsp;<a href="http://what-when-how.com/excel-vba/userform-techniques-and-tricks-in-excel-vba/">http://what-when-how.com/excel-vba/userform-techniques-and-tricks-in-excel-vba/</a></p>
		<p style="margin-left: 40px"> 
		<a href="https://gregmaxey.com/word_tip_pages/userforms_advanced_tips.html">
		https://gregmaxey.com/word_tip_pages/userforms_advanced_tips.html</a></p>
		<ul>
			<li><strong>Cheklist</strong>: Formunuz bittikten sonra genel bir 
			kontrol listesine göre eksikleri kontrol etmek güzel bir 
			alışkanlıktır.<ul>
				<li>Hizalamalar tamam mı?</li>
				<li>Aynı kümedeki benzer özellikli kontrollerin ölçüleri eşit 
				mi?</li>
				<li>Tab indexler doğru sırada mı?</li>
				<li>Esc tuşuna basılarak formdan çıkılabiliyor mu?</li>
				<li>Form başlığı belirlendi mi?</li>
				<li>Formunuz bir add-in'de kullanılacaksa Add-in'den açılışı 
				test ettiniz mi?</li>
			</ul>
			</li>
		</ul>
		</div>
		<h2 class="baslik">Çeşitli Örnekler</h2>
		<div class="konu">
		<h4 class="baslik">Data Formları</h4>
		<div class="konu">
		<p>Bu başlık altında bir örnek olmayacak. Birçok yerde bu konu 
		anlatılırken, verilen örneklerde Data Formlarını çok gördüğüm için ben 
		de başlık olarak koydum ama konuyu bir örnekle anlatmak için değil, size bunun için 
		başka bir alternatif önermek için.</p>
		<p >Ben bu iş için Access kullanmanızı öneriyorum.&nbsp;Access'in 
		güzelliği sözkonusu datayı gerçek bir veritabanı uygulamasında saklıyor 
		olmasıdır. Bu anlamda Excel'i çok da veritabanı uygulaması gibi 
		kullanmanızı önermiyorum. Bunun için belki bir süre sonra bu siteye 
		temel düzeyde Access anlatan sayfalar da koyabilirim.</p>
		</div>
		<h4 id="kokpit" class="baslik">Kokpit uygulaması</h4>
				<div class="konu">
		<p>Bu uygulamayı aynen burdaki gibi çalıştırabilmeniz için 
		<a href="../../Ornek_dosyalar/Makrolar/raporlar.rar">bu eki</a> 
		indirmenizi tavisye ederim. Ek indikten son içindekileri C:\ sürücüsü 
		altında "raporlar" diye bir klasör oluşturup buraya kopyalayın. Bu ek ile uğraşmak yerine kodlarda gerekli 
		değişiklikleri yaparak da kendi istediğiniz adreslerdeki dosyaların 
		açılmasını sağlayabilirsiniz.</p>
		<p>Kokpit dosyasının kendisine ise
		<a href="../../Ornek_dosyalar/Makrolar/Kokpit.xlsm">bu ekten</a> 
		ulaşabilirsiniz.</p>
		<p>Bu örnek ile departmanınızda/bölümünüzde sık kullanılan dosyalara 
		belli kategoriler aracılığıyla ulaşılmasını sağlayabilecek, kimin ne 
		zaman hangi dosyaya ulaştığının da log kaydını tutmuş olabileceksiniz.
		<a href="Dosyaislemleri_Dosyaokumaveyazma.aspx#logger">Logger</a> 
		örneğini inceleyerek bu log kaydının nasıl tutulduğunu detaylıca 
		öğrenebilirsiniz.</p>
		<p>Ana ekran görüntüsü aşağıdaki gibi olan formumuzda 4 ana sekme 
		bulunuyor. sekmelerden bazılarında istenilen döneme ait raporun 
		açılmasını sağlayana combobxlar bulunuyor. Ayrıca tüm geçmiş raporların 
		da görüntülenmesini sağlamak için her sekmenin sağında mavi yazılarla 
		yazılmış, üzerine gelindiğinde büyük + işaretine dönen linkle 
		bulunmakta. Örnek olduğu için tüm düğmeler çalışmamaktadır, sadece belli 
		butonlara kod ataması yapılmıştır.</p>
		<p><img src="../../images/vbauserformkokpit1.jpg"></p>
		<p>Şimdi kodların üzerinden geçelim:</p>
		<p>Öncelikle, dosya açılır açılmaz çalışacak koda bakalım. Dosya 
		açıldığında, başkalarında açık kalması bazen probleme neden olabildiği 
		için, ilgili kişinin pc'sinde dosyanın gece 00:00da kapanmasını 
		sağlıyoruz. Sonra Kokpiti kimin ne zaman açtığını kaydedecek log 
		prosedürünü çağırıyoruz. son olarak da anaformumuzu gösteriyoruz.</p>
		<pre class="brush:vb">
Private Sub Workbook_Open()
    Application.OnTime TimeValue("23:59:59"), procedure:="kapat", schedule:=True
    Call logkaydı
    Anaform.Show vbModeless
End Sub		</pre>
		<p>
		ana form açılır açılmaz çalışacak kodu ise Initialize eventi içne 
		yazıyoruz. </p>
		<ul>
			<li>küçültme büyütme işlemlerinde kullanmak üzere boy ve üst nokta 
			ölçülerini alıyoruz. Tabi bunlar en tepede global olarak tanımlanan 
			değişkenler olmalı.</li>
			<li>Excel dosyanın kendisi gizli değilse gizliyoruz, ikinci kez açma 
			kapama durumlarında hata almamk için önce gizli olup olmadığını 
			kontrol eidyoruz.</li>
			<li>2 tane log butonununu sadece sizde(bu örnekte benim pc adım 
			yazılı, siz kendi pc adınızı yazarsınız) açılmasını sağlıyorsunuz. 
			Bu log butonlarında log dosyalarının(txt formatlıdır) içeriğinin 
			aktarıldığı Excel dosyalar açılmaktadır. (Bu örnekte txtden excele 
			alma detayı anlatılmamıştır)</li>
			<li>sonra da comboboxların ilk değer atamalarını, çeşşitli 
			yöntemlerle, yapıyoruz.</li>
		</ul>
		<pre class="brush:vb">
Private Sub UserForm_Initialize()
dHeight = Me.Height
dTop = Me.Top

If Windows("Kokpit.xlsm").Visible Then
    Windows("Kokpit.xlsm").Visible = False
End If

'log butonnları benden başkasına görünmesin
If Environ("username") <> "Volki" Then
    Me.cmdDetayLog.Visible = False
    Me.cmdLogAna.Visible = False
End If

'AddItem ile eleman ekleme
Me.cbYıl.AddItem (Yıl)
Me.cbYıl.AddItem (Yıl - 1)
Me.cbYıl.Text = Yıl 'veya Value

'List ve Array ile eleman ekleme
Me.cbGün.List = Array(1, 2, 3)
Me.cbGün.Value = 1 'veya Text

'düngüsel olarak 12 ayı doldurma
For i = 1 To 12
    'Me.cbAy.AddItem i 'bölgesel ayarlarda tarih formatının durumuna göre burası veya aşağısı
    Me.cbAy.AddItem IIf(i < 10, "0" & i, i)
Next i
Me.cbAy.Value = "01"
End Sub
</pre>
		<p>
		Rapor açan düğmelerdeki kodlardan birine örnek aşağıdaki gibidir. Burada 
		önce detay rapor loguna baz teşkil edecek işlemler yapılıyor, sonra, 
		açılacak dosyanın oluşuş oluşmadığı kontrol edildikten sonra rapor 
		açılmaya çalışılıyor. dosya henüz oluşmadıysa bir uyarı veirliyor. Dosya 
		oluşmasnı kontrol eden örneğin detayını buradan incelyebnilirsiniz. 
		(NOT: Benim, kurumumda yaptığım gibi tam otomatik işleyen bir sistemde, 
		ilgili raporlar uygun zamanı bekleyip kendileri çalışır, kendileri uygun 
		yere kaydolur ve ilgili kullanıcılara maille 'raporçıktı' 
		bilgilendirmesi yapılır. O yüzden bu tür bir okntrolün ypaılması 
		anlamsız olabilir, ama fazla kontrol göz çıkarmaz desturuyla hareket 
		edelim ve kontrolümüz yapalım)</p>
		<pre class="brush:vb">
Private Sub CommandButton25_Click()
On Error GoTo hata
rapor = "İşbirimi_Hacimsel_Gelişim"
frekans = "Aylık"
Call detayraporlogu(rapor, frekans)

dosya = aylıkyol & Me.cbYıl.Value & "\İşbirimi Hacimsel Gelişim Raporu.xlsx"
If dosyavarmı(dosya) Then
    Workbooks.Open Filename:=dosya, ReadOnly:=True
Else
    MsgBox "Dosya henüz oluşmamış, Volkanla görüşün"
End If
Exit Sub

hata:
MsgBox "Bi sorun oluştu, Volkanla görüşün"
End Sub</pre>
		<p>
		Tüm rapor arşivini gösteren kodumuz aşağıdaki gibidir</p>
		<pre class="brush:vb">
Private Sub Label10_Click()
On Error GoTo hata
Shell "explorer.exe" &amp; " " &amp; günlükyol &amp; "Günsonu Bakiyeler", vbMaximizedFocus
Exit Sub

hata:
MsgBox "Bi sorun oluştu, Volkanla görüşün"
End Sub</pre>
		<p>
		Aşağıdaki kodlar ise sırayla, bir access dosyası, bir internet linki ve 
		bir word dosyasını açan düğmelerin kodları bulunmakta</p>
		<pre class="brush:vb">
Private Sub CommandButton29_Click()
On Error GoTo hata

On Error Resume Next
Set ac = GetObject(, "Access.Application")
If ac Is Nothing Then
    Set ac = GetObject(, "Access.Application")
    ac.opencurrentdatabase "C:\raporlar\hedefler.accdb"
    ac.UserControl = True
    Set ac = Nothing
End If

Exit Sub
hata:
MsgBox "Bi sorun oluştu, Volkanla görüşün"

End Sub
--------------------------------------------------
Private Sub CommandButton30_Click()
    Shell ("Explorer http://www.excelinefendisi.com/Excelent/KullanimKilavuzu.pdf")
End Sub
--------------------------------------------------
Private Sub CommandButton31_Click()
On Error GoTo hata

Set wordapp = CreateObject("Word.Application")
Set wordDoc = wordapp.documents.Open("C:\raporlar\satış tanımları.docx")
wordapp.Visible = True
    
Exit Sub
hata:
MsgBox "Bi sorun oluştu, Volkanla görüşün"

End Sub</pre>
		<p>
		Formu büyütüp/küçültme işlemi aşağıdaki kodla yapılır.</p>
		<pre class="brush:vb">
Private Sub ToggleButton1_Click()
If Me.ToggleButton1.Value = True Then
    Me.Height = dHeight * 0.1
    Me.Top = 0
    Me.ToggleButton1.Caption = "Büyüt"
Else
    Me.Height = dHeight
    Me.Top = 150
    Me.ToggleButton1.Caption = "Küçült"
End If
End Sub</pre>
		<p>
		Son olarak
		form kapanırken, dosyayı da kapatıyoruz, kapanırken kaydolmasın 
		istiyoruz ve dosya açılırken schedule ettiğimiz kapat makrosunu devreden 
		çıakrıyoruz.</p>
		<pre class="brush:vb">
Private Sub UserForm_Terminate()
    Application.OnTime TimeValue("23:59:59"), procedure:="kapat", schedule:=False
    Windows("Kokpit.xlsm").Close savechanges:=False
End Sub
</pre>
</div>
		<h4 class="baslik">Spin buttonlu filtre değiştirme formu</h4>
		<div class="konu">
		<p>Yakında...</p>
		</div>
		<h4 class="baslik">Otomatik mail gönderme </h4>
		<div class="konu">
		<p> 
		Otomatik mail gönderme işlemi Outlook nesne modelini bilmeyi 
		gerektirdiği için onunla ilgili örneği
		<a href="DigerUygulamalarlailetisim_OutlookProgramlama.aspx#toplumail">buraya</a> 
		koydum.</p>
		</div>
		<h4 class="baslik">SQL Çalıştırma formu</h4>
		<div class="konu">
		<p> 
		<img src="../../images/vbauserformsql.jpg"></p>
		<p> 
		Bu form ile Toad, AQT, SQL Developer gibi araçlardan çektiğiniz büyük 
		dataları Excel'e yapıştırma zahmetinden kurtulmuş olursunuz, zira 
		bununla, istediğiniz sonuç doğrudan Excelin içine yerleşir.</p>
		<p> 
		Bunun için 
		Veritabanlarıyla ilgili 
		<a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">bölümde</a> anlatılan konuları bilmeniz gerekiyor. 
		Bu örneği normalde oraya koymam gerekirdi, ancak userformlarla neler 
		yapılabileceğine ait güzel bir örnek olduğu için buraya koydum.</p>
		<p> 
		İlk yapmamız gereken, formu açan bir kod yazmaktır. Aşağıdaki bu mini 
		kodu ya bir add-indeki düğmeye ya da QAT üzerine yerleştireceğimiz bir 
		düğmeye atarız. Siz şimdilik personal.xlsb dosyasında bir modüle koyarak 
		da ilerleyebilirsiniz.</p>
		<pre class="brush:vb">Sub adosql()
   frmSQL.Show
End Sub</pre>
		<p>Sonrasında ise formumuz açılır ve Çalıştır butonundaki kodumuz 
		aşağıdaki gibidir. Aşağıda commentlerde belirtildiği gibi, eğer 
		bağlandığımız veritabanı Oracle veya DB2 gibi şifre kullanımı zorunlu 
		olan bir database ise şifre değişkenini kullanmanız gerekir, ve 
		connection stringinizi de buna göre ayarlamanız gerekir, bunlara ait 
		bilgiler <a href="DigerUygulamalarlailetisim_VeritabaniProgramlama.aspx">
		Veritabanı programlama sayfasında</a> bulunuyor. Ancak biz şuan Access 
		gibi şifre zorunluluğu olmayan bir veritabanına bağlandığımız için 
		şimdilik bu değişkeni commentle pasif hale getirdik.</p>
		<pre class="brush:vb">Private Sub CommandButton2_Click()
'önce tools&gt;references'tan microsoft ado 6.1 seçilmeli

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim strDB As Stream
Dim strSQL As String
Dim constr As String
'Static şifre As String 'her çalıştırma sırasında sormasın diye, eğer şifreniz yoksa uncommentli kalsın, şifreyle ulaştığınız bir database sözkonusuysa comment işaretini kaldırın

On Error GoTo hata

Me.Hide 'formu gizliyoruz
strSQL = frmSQL.TextBox1.Text
If strSQL = "" Then Exit Sub

'şifreli bir veritabanı sözkonusuya aşağıdaki commentleri kaldırın
'If şifre = "" Then
' şifre = InputBox("Şifrenizi giriniz")
'End If

cevap = MsgBox("yeni dosya mı olacak", vbYesNoCancel)
constr = "Provider = Microsoft.ACE.OLEDB.12.0; data source=C:\Users\Volki\Documents\My Web Sites\mysite\Ornek_dosyalar\Makrolar\vbausrformsql.accdb"
con.Open ConnectionString:=constr

Application.ScreenUpdating = False

rs.Open Source:=strSQL, ActiveConnection:=con, CursorType:=adOpenKeyset, LockType:=adLockOptimistic
rs.MoveFirst

If cevap = vbYes Then
Workbooks.Add
End If

'önce başlıklar
For i = 0 To rs.Fields.Count - 1
ActiveCell.Offset(0, i).Value = rs.Fields(i).Name
Next i
'şimdi datayı yapıştıralım
ActiveCell.Offset(1, 0).Select
ActiveCell.CopyFromRecordset rs

'burdan sonrasında isterseniz özel tablo formatları da uygulayabilirsiniz

rs.Close
con.Close
Set rs = Nothing
Set con = Nothing

Unload frmSQL 'formu bellekten siliyoruz
Application.ScreenUpdating = True
Exit Sub

hata:
MsgBox Err.Description
Application.ScreenUpdating = True

End Sub
</pre>
		<p>Bu arada kodu elle yazmak yerine hazır kaydedilmiş bir sql 
		dosyasından da getirebilirsiniz, bunun için formdaki ilgili düğmedeki 
		koda atanan kod ise aşağıdaki gibidir.</p>
		<pre class="brush:vb">Private Sub CommandButton1_Click()
Dim fd As FileDialog
Dim fso As New FileSystemObject
Dim ts As TextStream

Set fd = Application.FileDialog(msoFileDialogFilePicker)

If fd.Show = 0 Then
Exit Sub
End If

Set ts = fso.OpenTextFile(fd.SelectedItems(1))
içerik = ts.ReadAll
ts.Close

Set ts = Nothing
Set fso = Nothing

Me.TextBox1.Text = içerik
End Sub</pre>
</div>
		<h4 class="baslik" id="bolme">Dosya Bölme formu</h4>
		<div class="konu">		
		<p> 
		Bu form, çalıştığım kurumda şuana kadar en çok rağbet gören <strong>
		Dosya Bölme</strong> makromu 
		içeren formdur. Aslında favori olma konusunda buna eşlik eden bir de 
		otomatik mail gönderme formu var, ki buna da yukarıda yer verdim. İşte 
		bu meşhur toplu mail gönderme işleminde parametrik ek 
		kullanımı da olacaksa&nbsp;bu makro ile bu ekleri parçalama işlemi yapılmaktadır. Mail gönderme formuna 
		ise
		<a href="DigerUygulamalarlailetisim_OutlookProgramlama.aspx#toplumail">
		buradan</a> ulaşabilirsiniz.</p>
		<p> 
		Bölme işleminde temel olarak
		<a href="DizilerveDizimsiYapilar_Dictionaryler.aspx">Dictionary</a> 
		kullanma yoluna gittim. Bunun ilk halinde dictionary kullanmıyordum ve 
		büyük dosyaları bölme işlemi uzun sürüyordu. Sonradan kodu elden geçirip 
		bu hale getirdim. (Excelent menüsünden indirebileceğiniz VSTO 
		add-in'imde ise ilk yöntemde kullandığım metodolojiyi benimsemiştim. 
		Ancak burdaki kodlar doğrudan VBA değil, VB.Net kodları olduğu ve kod 
		dönüştürme işlemi de zahmetli olduğu için buna henüz vakit ayıramadım. 
		İlk fırsatta bu dönüştürme işlemini de yapacağım.)</p>
		<p> 
		Evet, şimdi kodları incelemeye başlayabiliriz.(Formun ve kodların olduğu 
		dosya sayfanın başındaki user_formlardır).</p>
		<p> 
		&nbsp;Diyelim ki, elimizde şağıdaki gibi bir liste var. Her bir bölge 
		için ayrı dosya oluşturmak istiyoruz.</p>
		<p> 
		<img src="../../images/vbauserformbolme2.jpg"></p>
		<p> 
		Hedef olarak görmek istediğimiz şey şöyle:</p>
		<p> 
		<img src="../../images/vbauserformbolme3.jpg"></p>
		<p> 
		Bölme formumuzu açmak için, ya bir Add-in'deki düğmey ya da QAT 
		üzerindeki bir düğmeye aşağıdaki kodu atarız. Siz şimdilik personal.xlsb 
		üzerinden veya örnek dosya üzeriine gelip, doğrudan forma gelip F5 
		tuşuna basarak da formu aktive edebilirsiniz.</p>
		<pre class="brush:vb">Sub BölmeAç()
   frmBöl.Show
End Sub</pre>
		<p>Aşağıdaki gibi formumuz açılır.</p>
		<p> 
		<img src="../../images/vbauserformbolme1.jpg"></p>
		<p> 
		Bu kontrollere verdiğim isimleri tek tek burda yazmama gerek yok, kod 
		içindenkendiniz de bakabilrisinirsiniz.</p>
		<p> 
		Öncelikle form içindeki kodlara bakalım, sonrasında ana kodun bulunduğu 
		modül koduna bakacağız.</p>
		<p> 
		ilk olarak Initialize event koduna bakıyoruz. Burada comboların içeriği 
		dolduruluyor ve bir tanesi gizleniyor.</p>
		<pre class="brush:vb">
Private Sub UserForm_Initialize()
Me.cbDosyatip.List = Array("Excel", "PDF")
Me.cbDosyatip.Value = "Excel"

Me.cbPrint.List = Array("Landscape", "Portrait")
Me.cbPrint.Value = "Landscape"

Me.cbPrint.Visible = False
End Sub	</pre>
		<p> 
		Format korunsun checkbox'ına tıklandığında, tick konmuşsa Dosyatip 
		comboboxında seçilen değere göre bir mesaj çıkmakta, bu mesaj her 
		checkbox tıklanışında çıkmasın diye
		<a href="Temeller_DegiskenlerveVeriTipleri.aspx#static">static</a> 
		değişkenle kontrol edilmektedir, ayrıca yine tick konması durumunda 
		Validation checkbox'ı da aktif hale getirilmekte, tick kaldırılınca 
		tekrar pasif olmaktadır.</p>
		<pre class="brush:vb">
Private Sub chkFormat_AfterUpdate()
Static i As Integer 'bu chechkbox her değiştiğinde sürekli bu msgbox çıkmasın diye, bi kere uyarması yeterli
If Me.chkFormat.Value = True Then
    If i = 0 Then
        If Me.cbDosyatip.Value = "Excel" Then 'formatlı olsa bile pdf hızlı çalışır
            MsgBox "Dosya tipi Excel seçildiğinde, format korunursa işlem daha uzun sürecektir." & vbCrLf & _
            "Süre önemliise ya dosya tipini PDF seçin ya da işlemi formatsız yapın"
        End If
        i = i + 1
    End If
    Me.chkValidation.Enabled = True
Else
    Me.chkValidation.Enabled = False
End If
End Sub</pre>
		<p>
		Print checkbox'ı seçildiğinde ise print layoutunun gösterildiği combobox 
		gösterilmekte, seçim kaldırıldığında tekrar gizlenmektedir.</p>
		<pre class="brush:vb">
Private Sub chkPrint_AfterUpdate()
    Me.cbPrint.Visible = Me.chkPrint.Value
End Sub
</pre>
	
		<p>
		Aşağıdaki kod ise Çalıştır düğmesindeki kod olup, ana kod için ön 
		hazırlık yapmakta ve en son çeşitli parametrelerle ana kodu çağırmakta. 
		Burda iki kontrol bulunuyor. Formun sol üst köşesindeki iki işlemin 
		yapılmış ve bu chekboxların da işaretlenmiş olması lazım, aksi halde bir 
		mesaj gösterilmekte ve kodun çalışması durmaktadır.</p>
		<pre class="brush:vb">
Private Sub CommandButton1_Click()
On Error GoTo hata
Dim printayar As String

'kontroller
If Me.chkKontrolilkkolon.Value = False Then
    MsgBox "bölmeye baz teşkil edecek kolon ilk kolonda olmalı." & vbCrLf & _
    "Eğer durum gerçekten böyleyse 'Kontrol' çerçevesi içindeki ilgili checkboxı işaretleyin"
    Exit Sub
End If

If Me.chkKontrolSıralı.Value = False Then
    MsgBox "Datanız sıralı olmalı. Eğer durum gerçekten böyleyse 'Kontrol' çerçevesi içindeki ilgili checkboxı işaretleyin"
    Exit Sub
End If

'böl klasörü yoksa yaratalım
If filefolderexists("C:\böl") = False Then MkDir ("c:\böl")

'A kolonunda / işaretei kontrolü. Zira dosya isimlerinde / işareti olamaz.
On Error Resume Next 'bulamazsa devam etsin diye
Columns("A:A").Select
Selection.Replace what:="/", replacement:="-", lookat:=xlPart, _
    searchorder:=xlByRows, MatchCase:=False, searchformat:=False, ReplaceFormat:=False
'başlık satırından sonraki satırda hiç boş hücre olmamalı, space yapalım
Rows(Me.txtBaşlık.Value + 1).Replace what:="", replacement:=" ", lookat:=xlWhole, _
    searchorder:=xlByRows, MatchCase:=False, searchformat:=False, ReplaceFormat:=False
    
'hata kontrolünü tekrar getirelim
On Error GoTo hata

If Me.chkPrint.Value = True Then
    printayar = Me.cbPrint.Value
End If

Cells(CInt(Me.txtBaşlık.Text), 1).Select
Call filtrekontrol 
Application.Wait (Now + TimeValue("00:00:02"))
Call bölmekodu(Me.lblKlasör.Caption & "\", printayar, Me.cbDosyatip.Value, Me.chkFormat.Value, CInt(Me.txtBaşlık.Text), Me.chkValidation.Value)
Unload Me
Exit Sub

hata:
MsgBox "bir hata oluştu, volkanla görüşün" & vbCrLf & Err.Description
End Sub
</pre>

<p>Esas bölmeyi yapan kod ise şöyledir. Kod içinde yer yer açıklamalar var, ancak ilk etapta
F8 ile giderseniz anlaması daha kolay olacaktır.
</p>

<pre class="brush:vb">
Sub bölmekodu(klsr As String, pr As String, dosyatip As String, dformat As Boolean, bs As Integer, validateformat As Boolean)

Dim dict As New Scripting.Dictionary
Dim stbar As String, progress_char As String
Dim başlık As Variant, alan As Variant
Dim anaDosyam As Workbook, yeniDosyam As Workbook
Dim kolon As Integer

On Error GoTo hata 'kontroller v.s buton clikte yapılıyor

stbar = Application.StatusBar
Application.StatusBar = "işlem yapılıyor, bekleyiniz..."
Application.ScreenUpdating = False
Application.DisplayAlerts = False

devam:
Set anaDosyam = ActiveWorkbook
progress_char = Chr(8) & " "

isim = CreateObject("Scripting.FileSystemObject").GetBaseName(anaDosyam.Name)
başlık = Range(Range("a1"), Cells(bs, 1).End(xlToRight)) 'kaynaktan okuma to variant
kolon = Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight)).Cells.Count

Cells(bs + 1, 1).Select

Do 'dict ile uniq idleri alalım
    If Not dict.Exists(ActiveCell.Value) Then dict.Add ActiveCell.Value, ActiveCell.Row
    ActiveCell.Offset(1, 0).Select
Loop Until ActiveCell.Value = ""

Range("a1").Select 'do loop içinde en aşağı inmiştik, tekrar başa çıkalım
Set yeniDosyam = Workbooks.Add
Range(Range("a1"), Cells(bs, kolon)).Value = başlık 'hedefe yazdırma from variant

'print ayarı
If pr <> "" Then
    With ActiveSheet.PageSetup
        If pr = "Landscape" Then
            .Orientation = xlLandscape
        Else
            .Orientation = xlPortrait
        End If
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
End If

If dformat = False Then
    If dosyatip = "Excel" Then
        For Each d In dict.Keys
            anaDosyam.Activate
            ActiveSheet.Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight).End(xlDown)).AutoFilter Field:=1, Criteria1:=d
            alan = ilkvisiblesonrasıalan(Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight).End(xlDown))) 'kaynaktan okuma to variant
            yeniDosyam.Activate
            Range(Cells(bs + 1, 1), Cells(UBound(alan) + bs, kolon)).Value = alan 'hedefe yazdırma from variant
            yeniDosyam.SaveAs Filename:=klsr & Trim(d) & "-" & isim & ".xlsx", FileFormat:=xlWorkbookDefault
            Range(Cells(bs + 1, 1), Cells(UBound(alan) + bs, kolon)).Clear
            i = i + 1
            DoEvents
            Application.StatusBar = "Tamamlanma Oranı: " & WorksheetFunction.Rept(progress_char, Int(i * 100 / dict.Count)) & " %" & Int(i * 100 / dict.Count)
        Next d
        ActiveWorkbook.Close savechanges:=False ' son dosyayı kaydetmeden kapat, çünkü clear yapıldı, kaydetmeyelim
    Else 'PDF
        For Each d In dict.Keys
            anaDosyam.Activate
            ActiveSheet.Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight).End(xlDown)).AutoFilter Field:=1, Criteria1:=d
            alan = ilkvisiblesonrasıalan(Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight).End(xlDown))) 'kaynaktan okuma to variant
            yeniDosyam.Activate
            Range(Cells(bs + 1, 1), Cells(UBound(alan) + bs, kolon)).Value = alan 'hedefe yazdırma from variant
            yeniDosyam.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                    klsr & Trim(d) & "-" & isim _
                    , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
                    :=False, OpenAfterPublish:=False
            Range(Cells(bs + 1, 1), Cells(UBound(alan) + bs, kolon)).Clear
            i = i + 1
            DoEvents
            Application.StatusBar = "Tamamlanma Oranı: " & WorksheetFunction.Rept(progress_char, Int(i * 100 / dict.Count)) & " %" & Int(i * 100 / dict.Count)
        Next d

        ActiveWorkbook.Close savechanges:=False ' son dosyayı kaydetmeden kapat, zaten pdf yapıyoruz
    End If
Else 'format korunacaksa
    If dosyatip = "Excel" Then
        For Each d In dict.Keys
            yeniDosyam.ActiveSheet.Range("A1").CurrentRegion.Clear
            anaDosyam.Activate
            ActiveSheet.Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight).End(xlDown)).AutoFilter Field:=1, Criteria1:=d
            anaDosyam.ActiveSheet.Range(Cells(1, 1), Cells(bs, 1).End(xlDown).Offset(0, kolon - 1)).Copy
            yeniDosyam.Activate
            If validateformat = True Then
                Range("a1").PasteSpecial Paste:=xlPasteValidation
            End If
            Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                xlNone, SkipBlanks:=False, Transpose:=False
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            
            yeniDosyam.SaveAs Filename:=klsr & Trim(d) & "-" & isim & ".xlsx", FileFormat:=xlWorkbookDefault
            i = i + 1
            DoEvents
            Application.StatusBar = "Tamamlanma Oranı: " & WorksheetFunction.Rept(progress_char, Int(i * 100 / dict.Count)) & " %" & Int(i * 100 / dict.Count)
        Next d
        ActiveWorkbook.Close savechanges:=False ' son dosyayı kaydedip kapat, çünkü clear yapıldı, kaydetmeyelim
    Else 'PDF
        For Each d In dict.Keys
            yeniDosyam.ActiveSheet.Range("A1").CurrentRegion.Clear
            anaDosyam.Activate
            ActiveSheet.Range(Cells(bs, 1), Cells(bs, 1).End(xlToRight).End(xlDown)).AutoFilter Field:=1, Criteria1:=d
            anaDosyam.ActiveSheet.Range(Cells(1, 1), Cells(bs, 1).End(xlDown).Offset(0, kolon - 1)).Copy
            yeniDosyam.Activate
            Range("a1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
                xlNone, SkipBlanks:=False, Transpose:=False
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            yeniDosyam.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                    klsr & Trim(d) & "-" & isim _
                    , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
                    :=False, OpenAfterPublish:=False
            i = i + 1
            DoEvents
            Application.StatusBar = "Tamamlanma Oranı: " & WorksheetFunction.Rept(progress_char, Int(i * 100 / dict.Count)) & " %" & Int(i * 100 / dict.Count)
        Next d
        ActiveWorkbook.Close savechanges:=False ' son dosyayı kaydetmeden kapat,zaten pdf yapıyoruz
    End If
End If

 

Application.StatusBar = stbar
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Call filtrekontrol
Shell "explorer.exe" & " " & klsr, vbMaximizedFocus

Exit Sub
hata:
Application.StatusBar = stbar
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox Err.Description & vbCrLf & "Volkanla görüşün"

End Sub

Function ilkvisiblesonrasıalan(alan As Range) As Range
    Dim ilk As Range
    Dim son As Range
    Dim N As Integer, R As Integer
    
    N = alan.Columns.Count
    R = alan.SpecialCells(xlCellTypeVisible).Cells.Count / N - 1

    Set ilk = alan.Offset(1, 0).Resize(alan.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible).Cells(1, 1) 'bu kısım _
    ilk görünen hücreyi verir
    Set son = ilk.Offset(0, N - 1)
    Set ilktoright = Range(ilk, son)
    If R > 2 Then
        Set ilkvisiblesonrasıalan = Range(ilktoright, ilktoright.End(xlDown))
    Else
        Set ilkvisiblesonrasıalan = ilktoright
    End If
End Function

Sub filtrekontrol()

If ActiveSheet.AutoFilterMode = True Then
    If ActiveSheet.FilterMode = False Then
        'nothing
    Else
        ActiveSheet.ShowAllData
    End If
Else
    Selection.AutoFilter
End If

End Sub

Function filefolderexists(dosyaTamAdres As String) As Boolean
    If Not Dir(dosyaTamAdres, vbDirectory) = vbNullString Then filefolderexists = True
End Function

</pre>
		
		<p>
		Aşağıdaki kod ile hedef klasör değiştirilebilmektedir. Default değer
		<a href="file:///C:/böl">C:\böl</a> klasörüdür.</p>
		<pre class="brush:vb">
Private Sub CommandButton2_Click()
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFolderPicker)
With fd
    .Title = "Klasör seçin"
    If .Show = True Then
        lblKlasör.Caption = .SelectedItems(1)
    End If
End With
End Sub</pre>
		
		<p> 
		Bölme işlemi bittikten sonra kontrol ediyoruz ve gerçekten doğru olarak 
		bölündüğünü görüyoruz.</p>
		<p> 
		<img src="../../images/vbauserformbolme4.jpg"></p>
		<p> 
		Burda ise, formatın ve <a href="../Excel/DataMenusu_VeriDogrulama.aspx">
		validation</a> içeriklerinin korunduğu bir örneği görüyorsunuz.&nbsp;</p>
		<p> 
		<img src="../../images/vbauserformbolme6.jpg"></p>
		</div>
		<h4 class="baslik">Worksheet Formdaki değişime göre bir makronun 
		çalışması</h4>
		<div class="konu">		
		<p> 
		Bu örnekteki form çeşidi her ne kadar
		<a href="../Excel/DeveloperMenusu_Kontroller.aspx">worksheet formların</a> 
		konusu olsa da, işin büyük kısmı makro ile yapıldığı için bunu da buraya 
		aldım. Bunun için de biraz veritabanı uygulamarıyla iletişim bilmek 
		gerekiyor, ancak ben bunu veritabanı konusu yerine bu sefer buraya 
		almayı tercih ettim. Örnek dosyaları 
		<a href="../../Ornek_dosyalar/Makrolar/worksheet%20kontrol%20-%20vba.rar">burdan</a> indirebilirsiniz. Access 
		dosyayı uygun bir klasöre koyup aşağıdaki constr değişkenindeki konumunu da 
		değiştirmeniz gerekmektedir.</p>
		<p> 
		<img src="../../images/vbauserformsql1.jpg"></p>
		<p> 
		Listbox,'a sağ tıklayıp Control sekmesine geldim ve Input Range ile Cell 
		link özelliklerini aşağıdaki gibi değiştirdim.</p>
		<p> 
		<img src="../../images/vbauserformsql2.jpg"></p>
		<p> 
		A1-A5 arasını tamamen beyaz yaparsanız hiç görünmezler, hatta listboxı 
		tamamen A1-A5'i kapatacak şekilde üzerine de taşıyabilirsiniz./p>
		<p> 
		Her hücre içinde comment olarak eklenmiş SQL bulunmakta. Listboxtan bir 
		ürün seçildiğinde A5'e bu seçimin indeksi gelmekte, buna göre de ilgili 
		SQL çalıştırılmaktadır.</p>
<p>
<img src="../../images/vbauserformsqlcomment.jpg">
</p>
		<pre class="brush:vb">
Sub ListBox1_Change()
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim adet As Integer
    Dim constr As String
    Dim strsql As String
    'Static şifre As String
    
    On Error GoTo hata
    
    strsql = Cells([A5].Value, 1).Comment.Text
    If strsql = "" Then Exit Sub
    'şifresi olan bir databse ise aşağısı uncommentsiz
    'If şifre = "" Then
    '    şifre = InputBox("şifreyi girin")
    'End If
    
    constr = "Provider = Microsoft.ACE.OLEDB.12.0; data source=C:\falanfilanklasör\vbausrformsql.accdb"
    con.Open ConnectionString:=constr
    
    rs.Open Source:=strsql, ActiveConnection:=con, CursorType:=adOpenKeyset, LockType:=adLockOptimistic
    rs.MoveFirst
    
    [a7].Select
    Selection.CurrentRegion.ClearContents 'bir önceki run sonucunu temizleyelim
    'önce başlıkları getirelim
    For i = 0 To rs.Fields.Count - 1
        ActiveCell.Offset(0, i).Value = rs.Fields(i).Name
    Next i
    ActiveCell.Offset(1, 0).Select
    'şimdi datayı alalım
    ActiveCell.CopyFromRecordset rs
    
    rs.Close
    con.Close
    Set rs = Nothing
    Set con = Nothing
    Exit Sub
hata:
    MsgBox Err.Description
    Set rs = Nothing
    Set con = Nothing
End Sub</pre>
</div>
</div>
</asp:Content>
