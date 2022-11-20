<%@ Page Title='DebuggingveHataYonetimi Breakpointler' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>
<asp:Content ID="Content1" Runat="Server" ContentPlaceHolderID="SayfaIcerik"><div id='gizliforkonu'><table>
<tr><td>
<asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td>
<asp:Label ID='Label2' runat='server' Text='Debugging ve Hata Yönetimi'></asp:Label></td><td>
<asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>

<h1>Breakpointler ve İzleme Pencereleri</h1>
<p>Kodunuzu test ederken belirli noktalarda geçici olarak duraklatan araca
<strong>Breakpoint(BP)</strong> diyoruz. 
Kullanım amacı, kodu durdurduğunuz anda, o an itibarıyle değişkenlerin durumunu 
görmek olabileceği gibi, o ana kadar kodun ne tür etkiler yarattığını görmek 
de olabilir. Breakpoint konmuş durumdaki moda <strong>Break modu</strong> denir.</p>
	<p><strong>İzleme pencereleri</strong> ise, ister Break modunda olun ister olmayın 
	kodunuzun o an tam olarak ne yaptığını gösterebileceği gibi tek satırlık bir 
	kodun doğrudan burada çalışıtırılmasına da imkan verir.</p>
	<p>Şimdi bunlara detaylıca bakalım.</p>
	<h2 class="baslik">Breakpointler</h2>
	<div class="konu">
	<p><strong>Debug </strong>menüsünden <strong>Toggle Breakpoint </strong>diyerek 
	veya <strong>F9 </strong>tuşuna basarak BP koyabiliriz, veya var olan BP'i kaldırabiliriz. Genelde 
	menüyü kullanmak&nbsp;yerine ya F9'a tıklarız, ya da iki aşağıda bulunan 
	resimdeki bordo renkli dairenin bulunduğu yere tıklarız, kaldırmak için 
	bordo daireye tekrar tıklarız.</p>
	<p>Kodumuzun farklı yerlerine birden fazla olacak şekilde BP koyabiliriz. 
	Hepsini kaldırmak için menüde <strong>Clear All Breakpoints </strong>diyebilir veya 
	<strong>Ctrl+Shift+F9 </strong>tuşlarına basabiliriz.</p>
		<p>

<img src="/images/vbabreakpoint1.jpg"></p>
		<p>Örnek bir BP konmuş satırın 
		görüntüsü aşaıdaki gibidir. &nbsp;</p>
		<img src="/images/vbabreakpoint2.jpg"><p>
	Break moddayken izleme amacıyla, değişkenlerin üzerine gelip onların o 
	anki değerlerini görebilirsiniz, veya izleme(debugging) pencerlerini kullanabilirsiniz.</p>
	<p>BP'a denk geldikten sonra izleme dışında bir eylemde bulunmak isterseniz 
	3 seçeneğiz var:</p>
	<ul>
		<li>Reset tuşu ile durdurmak</li>
		<li>F5 ile tam gaz devam etmek</li>
		<li>F8 ile adım adım ilerlemek(F8 ve türevleri)</li>
	</ul>
	<p><strong>NOT</strong>:Bu BP'lar sadece ilgili oturum 
	için geçerlidirler, yani dosya ile birlikte kaydolmazlar.
	</p>
		<p><span class="dikkat">dikkat:</span>BP'a denk gelip incelemenizi 
	yaptığınızda kodu durdurmayı tercih ederseniz, dikkat etmeniz gereken bir 
	nokta var. Eğer program girişinde bazı Application propertylerine(DisplayAlerts 
		gibi) False 
	atadıysanız bunları tekrar True yapmadan durdurmamalısınız. Aksi halde üzücü 
		sonuçlarla karşılaşabilirsiniz.</p>
	<h3>Adım adım ilerlemek</h3>
		<h4>F8 ile adım adım ilerleme</h4>
		<p>Normalde bir makroyu çalıştırmak için Standart menü çubuğundan Play butonuna 
		veya F5 tuşuna basarız, böylece kod baştan sona tek seferde çalışmış olur.</p>
		<p>Ancak kodumuzun, hangi aşamalardan geçtiğini ve özellikle bir yerde 
		hata alıyorsa tam olarak nerede ve niçin hata aldığını görmek için F8 
		tuşu ile kodu çalıştırırız. Böylece kodunuz satır satır çalıştırılır ve 
		o ana kadar kodun neler yaptığını, değişkenlerin hangi değerlere sahip 
		olduğunu fareyi değişkenin üzerine gelip bekleterek görebilirsiniz.</p>
		<h4>Shift+F8 ile prosedürleri atlayarak adım adım ilerleme</h4>
		<p>Normal F8'le tek farkı, eğer kodumuzda bir satırda bir başka prosedürü 
		çağırıyorsak, bunun içine de girip satır satır ilermelek yerine o 
		prosedürü tamamen çalıştırıp tekrar ana prosedüre dönmemizi sağlar. Alt 
		prosedürün hatasız çalıştığından eminsek bunu kullanırız, böylece onun 
		içinde satır satır ilerleyerek boşuna vakit kaybetmemiş oluruz.</p>
		<h4>Shift+Ctrl+F8 ile adım adım modundan çıkma</h4>
		<p>Uzun bir süre bu kombinasyonun amacını tam anlayamamıştım, hatta bu 
		sayfayı ilk yazdığımda bu kombinasyonu gereksiz bulduğumu bile 
		belirtmiştim. Ama bir gün tam da böyle birşeye ihtiyacım oldu. Şöyle ki, 
		bir prosedür içindeyken F8 ile ilerlerken başka bir prosedüre 
		dallandığınızda bazen orada uzun bir döngüye girmiş olabiliyorsunuz ve 
		bir anca önce o prosedürden çıkıp ana prosedüre dönmek ve oradan F8 ile 
		devam etmek istiyorsunuz. İşte böyle bir durumda bu kombinasyon ile o 
		alt preosdürü hızlıca tamamlayıp ilk prosedürede kaldığınız yere 
		gelirsiniz.</p>
		<p>Şöyle düşünenleriniz olabilir tabi; bunu ana prosedürden alta 
		dallandığım yerin bir satır altına breakpoint koyarak ve F5'e basarak da 
		yapabilirdim. Evet yapabilirsiniz ama bunu en başta yapmanız gerekirdi. 
		Bir de bazen kodunuz o kadar karmaşık olabilir ki, ordan oraya 
		dallanarak gelmişsinizdir, tam olarak nereye breakpoint koyacağınızı 
		bile bilemeyebilisiz. O yüzden en güzeli, bu kombinasyonu 
		çalıştırmaktır.</p>
		<h4>Ctrl+F8 ile cursorın bulunduğu yere kadar hızlıca ilerleme</h4>
		<p>Bence bu kombinasyona da gerek yok, onun yerine ilglili yere BP koyun ve F5 yapın, 
		aynı görevi görür. Üstelik BP+F5'in şöyle bir avantajı da var, aynı kodu tekrar çalıştırmanız gerekebilir ve cursor farklı yerdeyse 
		tekrar oraya gelmeye çalışacaksınız, bu da ekstra zaman kaybı demektir. BP'i bi kere koyun ve 
		kontrolleriniz bitene kadar orada kalsın.</p>
		<h4>Ctrl+F9 ile çalıştırılacak kod kısmını belirlemek</h4>
		<p>Bu, break moddayken sarı oku mousela taşımaya benzer. Diyelim 5 
		sayfalık bir kodunuz var. Biliyorsunuzdur ki, kodun üst kısımlarında sorun 
		yok, son paragrafta bi yerde sorun çıkarıyor. Kodun üst kısımlarında F8 
		ile ilerlediniz, sonraki birkaç sayfalık kodu çalıştırmaya gerek yok diyorsunuz. 
		Şimdi normalde bu kombinasyon olmasaydı 
		fareyle sarı oku 5 sayfa aşağı indirmek gerekirdi, ki bunu deneyin, çok 
		sinir bozucu birşeydir. </p>
		<p>Halbuki, bu özellik 
		sayesinde, sarı okun gelmesini istediğim yere bi yere tıklarım ve 
		<strong>Ctrl+F9</strong>'a<strong> </strong>(veya debug menüsünde Set Next statement) 
		tıklarım, böylece aradaki 
		kodlar çalışmadan doğrudan bu satıra gelmiş olurum.</p>
		<p><img src="/images/vbadebugnext1.jpg" class="zoomla" width="60%" height="60%"></p>
		<p>Ctrl+F9'dan sonra doğrudan istediğimiz yere geldik.</p>
		<p><img src="/images/vbadebugnext2.jpg" class="zoomla" width="60%" height="60%"></p>
		<p>Tabi burda unutulmaması gereken birşey var; üst kısımlarda F8 ile 
		ilerlerken, aşağıya Ctrl+F9 ile geleceğiniz noktada bir değişken varsa, 
		üst tarafta buna değer atandığı yerleri de F8 ile geçtikten sonra 
		buraya sıçramamız gerekir yoksa istediğiniz sonucu elde 
		edemeyebilirsiniz.</p>
		<h3>Debug.Print ve Debug.Assert</h3>
		<h4>Debug.Print</h4>
		<p>Debug.Print ile biraz aşağıda değineceğimiz <strong>Immedaite Window'</strong>a doğrudan 
		birşeyler yazdırabiliyoruz. Ben genelde şöyle kullanıyorum: 
		Debug.Print ifadesinden hemen sonraki satıra BP koyarım, Immediate Windowa ne yazdığına bakarak testimin sonucunu görürüm, 
		veya testlerimde arka arkaya birşeylerin değerini test edeceksem MsgBox 
		alternatifi olarak kullanırım.</p>
		<p>Bunu illa Break modda da kullanmak zorunda değilsiniz, normal bir F5 
		çalıştırmasında da kullanılabilir.</p>
		<p><img src="/images/vbadebugprint1.jpg"></p>
		<p>&nbsp;</p>
		<p>
		<a href="../Fasulye/NeNeredeNasil_NullNothingEmptyveIlkdegeratama.aspx#DebugOrnekliKod">Burada</a> 
		içinde çok fazla Debug.Print barndıran bir örnek görebilrisiniz.</p>
		<h4>Debug.Assert</h4>
		<p>Bir nevi koşullu BP olarak düşünebileceğimiz Assert metodu nerede ve 
		ne zaman kullanılacağı biraz kafa karıştırcı bir metoddur. Kodunuzda bir hata olup olmadığını size önceden söyleyebilecek 
		güce sahiptir. </p>
		<p>Bunun, bir sonraki bölümde göreceğimiz, hata yakalama 
		ifadeleriyle karıştırılmaması gerekiyor. <strong>Hata yakalama blokları; 
		kullanıcı, kodu çalıştırırken bir hata ile karşılaştığında ne yapılması 
		gerektiğini söyler; Assert ise kod tasarlanırken kodu yazana nerde hata olduğunu 
		söyler</strong>. Yani kodunuzu çalıştırdığınızda bir yerlerde hata çıkıyor ama 
		bi türlü nerde olduğunu bulamıyorsunuzdur. İşte böyle durumlarda Assert 
		özelliğini baştan tanımlamanız durumunda hatayı kolaylıkla tespit 
		edebilirsiniz.</p>
		<p>Kullanım şekli şöyledir: <span class="keywordler">Debug.Assert koşul 
		gerçeklenmiyorsa</span></p>
		<p>Bu şekilde kullandığımızda VBA'e şunu demiş oluyoruz:<strong> Koşul 
		gerçeklenmediğinde Break moda gir.</strong></p>
		<p>Mesela aşağıdaki örnekte For döngüsü içinde bir yerde yeni sayfa 
		yaratılıyor. Ama biz 5'ten fazla sayfa olsun istemiyoruz. Kodumuz bize 
		bir şekilde 5ten fazla sayfa üretirse o anda Break moda girmiş olur ve 
		biz de duruma müdahale ederiz.</p>
		<pre class="brush:vb">
Debug.Assert Sheets.Count &lt;=5

For i=1 to x
   'çeşitli kodlar
    Sheets.Add
   'çeşitli kodlar
Next i</pre>
		<p>Bir diğer pratik kullanımı da bir döngüde hızlıca belli bi yere kadar 
		ilerlemek ve hatayı son bölümde incelemek olabilir. Mesela aşağıdaki kodu çalıştırdığınız aktif sayfa 
		1000 satırlık bir veri içeriyor olsun. Son satırda bi hata alıyorsunuz, F8 
		ile tek tek gitmek isteseniz 999 kere F8 yapmanız lazım, bunun yerine 
		<strong>Debug.Assert ActiveCell.Row < 1000</strong> diyip F5 yaparak 
		hızlıca 999. satıra geliriz.</p>
		<pre class="brush:vb">
Sub pagebreak()
    [a2].Select
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    Do Until IsEmpty(ActiveCell.Offset(1, 0))
        Debug.Assert ActiveCell.Row < 1000
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value <> ActiveCell.Offset(-1, 0).Value Then
            Set ActiveSheet.HPageBreaks(2).Location = ActiveCell
        End If
    Loop
End Sub		</pre>
		<h5>Debug.Assert False</h5>
        <p>Workbook_Open eventi gibi dosya henüz açık olmadığı için manuel kontrolü ele alamadığınız veya Worksheet_Change eventi gibi manuel debugging başlatamadığınız veya bir şekilde kodun bir yerine geldiğinde durmasını ve beklemeye geçmesini istediğiniz durumlarda kodun başlangıcında kontrolü ele almak isteyebilirsiniz. Bunu breakpointlerle de yapabilirsiniz ancak breakpointler dosya kapandığı zaman kaybolurlar. Bu breakpointleri tekrar tekrar oluşturmak istmeiyorsanız işte bu tür durumlarda Debug.Assert&#39;ün bu özel halini kullanabilirsiniz. Bu şekilde ilgili bu satıra gelindiği anda durur ve sizi bekler, bundan sonra F8 ile ilerleyebilirsiniz. </p>
        <p>Debug.Assert kullanımı konusunda daha geniş bilgiyi <a href="http://excelmacromastery.com/bulletproof-vba-code/">burada</a> 
		    bulabilirsiniz.</p>
		
	</div>
	
			<h2 class="baslik">İzleme Pencereleri</h2>
			<div class="konu">
			<p>3 adet izleme pencresi bulunuyor.</p>
				<ul>
					<li>Immediate Window</li>
					<li>Locals Window</li>
					<li>Watch Window</li>
				</ul>
				<p>Bunlardan en çok kullanılan dolayısıyla ana izleme penceresi 
				ünvanı verebileceğimiz pencere Immediate Window'dur. Sırayla 
				inceleyelim.</p>
				<h3>Immediate Window</h3>
				<p>Ctrl+G tuş kombinasyonu veya View meünsünden Immediate Window 
				butonu ile aktive edebilirsiniz.</p>
				<h4>Bununla neler yapabilriz?</h4>
				<ul>
					<li>F8 ile adım adım ilerlerken bu pencereyi açıp o an bir 
					değerin ne olduğunu görebiliriz</li>
					<li>Bir kod çalıştırabiliriz</li>
					<li>Application seviyesinde bazı özellikleri 
					sorgulayabiliriz.</li>
				</ul>
				<p>Birkaç örnekle bakalım.</p>
				<p>Mesela aşağıdaki kodda döngüde birkaç kez ilerledikten sonra 
				<strong>i</strong>'nin hangi değere ait olduğunu sorgulamak için Immediate 
				Window'a <strong>"?i"</strong> yazıp Enter'a bastım. Gördüğünüz 
				gibi bir 
				değişkenin değerinin ne olduğunu <strong>sorgulamak(bilgi 
				öğrenmek)</strong> için başına <strong>?</strong> işaret koyup Entera 
				basarız.</p>
				<p><img src="/images/vbadebugiw1.jpg"></p>
				<p>Bu sefer Application seviyesinde bir kod yazalım. Dikkat edin 
				F8 ile giderken yapmıyorum bunu, hatta bir makro içinde bile 
				yapmıyorum. Yine bilgi öğrenmek istediğim için başına ? koyuyorum. 
				Öğrendiğim bilgi de XLSTART klasörünün yeridir.</p>
				<p><img src="/images/vbadebugiw2.jpg"></p>
				<p>Şimdi ise sorgulama yapmak yerine bir kod çalıştıralım ve bu kod yine 
				Application 
				seviyesinde olsun.</p>
				<p><img src="/images/vbadebugiw3.jpg"></p>
				<p>Gördüğünüz gibi ilk olarak Enableevents özelliğine False 
				değerini atadım, hemen arkasından da ? koyarak bu özelliğin durumunu 
				öğrendim. Bu arada bazen geçici olarak bu ve bununu gibi 
				özelliklere bir değer atamanız gerekecek, bunun için bi prosedür 
				yaratıp içine bu kodu yazmaktansa bu pencereyi 
				kullanabilirsiniz.</p>
				<p>Son olarak da bi kodun içinde olalım olmayalım farketmez, daha normal(değişken 
				içermemesi kaydıyla) 
				bir kod çalıştıralım.</p>
				<p><img src="/images/vbadebugiw4.jpg"></p>
				<p>Bunu yazp Enter'a basınca bi mesaj kutusu çıkacaktır.</p>
				<h3>Locals Window</h3>
				<p><strong>Local </strong>penceresi, kodunuz üzerinde F8 ile tek 
				tek ilerlerken o anda değişkenlerin değerini ayrı ayrı 
				görebileceğiniz bir penceredir. Bu bizi aslında birsürü <strong>
				Debug.Print</strong> yazmaktan veya değişkenlerin üzerine gelip 
				bekleyerek pop-up kutucuk ile değişkenlerin değerini tek tek 
				görmee zahmetinden kurataran bir araçtır.</p>
				<p>Bu pencrede local değişkenlerin değerini ve UserForm 
				seviyesindeki tanımlanmış değişkenlerin değerini görebiliyoruz. 
				Modül seviyesindeki değişkenler için <strong>Watch</strong> 
				penceresine bakarız.</p>
				<p>Aşağıdaki pencerede ad değişkenine volkan değeri, i 
				değişkenine de 1 değeri atanmış durumda. Kelime değişkenine ise 
				henüz bir atama olmadığı için Empty görünmektedir. Ayrıca 
				farkettiyseniz değişkenlerin tiplerini de görmekteyiz. İlk iki 
				değişken deklare edilmediği için Variant tipindedir, ama 
				aldıkları değer itibarıyle String ve Integer değer 
				tutmaktadırlar.</p>
				<p><img src="/images/vbadebugginglocal.jpg"></p>
				<h3>Watch Window</h3>
				<p>Bu pencere ile Modül seviyesindeki değişkenlerin durumunu 
				görebiliyor, çeşitli koşullar girerek bu koşulların 
				gerçekleşmesi veya girdiğimiz değişkenlerin değeri değiştiğinde 
				kodun Break mod'a girmesini sağlayabiliyoruz.</p>
				<p>3 çeşit izleme çeşidi var.</p>
				<ul>
					<li><strong>Watch Expression</strong>:Local Window gibi 
					çalışır. Modül seviyesindeki değişiklikler de izlenir.</li>
					<li><strong>Break When Value is True</strong>:Verdiğiniz 
					koşul sağlandığında Break moda girer(Bir nevi <strong>
					Debug.Assert </strong>görevi görür)</li>
					<li><strong>Break When Value Changes</strong>:Verdiğiniz 
					değişkenin değeri değişince break moda girer</li>
				</ul>
				Aşağıdaki örnekte i=5 olunca break mod'a girsin demiş olduk. 
				Aynı anda hem Local hem Watch windowu görüyoruz.<p>
				<img src="/images/vbadebuggingwatch.jpg"></p>
				<h3>Call Stack</h3>
				<p>Bu pencereyi hiç kullanmadım açıkçası. Görevi, o anda aktif 
				olarak çağrı yapılmış tüm prosedürleri göstermek. İhtiyacınız 
				olmaz düşüncesiyle detaya girmiyorum.</p>
			</div>
</asp:Content>
