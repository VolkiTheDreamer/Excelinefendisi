<%@ Page Title='Formlar Temeller' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VBAMakro'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Formlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td></tr></table></div>
	<h1>Temeller</h1>
	<p>Bu bölümde küçük bir UserForm yaratıp, üzerine bikaç temel kontrol 
	koyacağız. Form eventlerini ve Form özelliklerini inceleyeceğiz.</p>
		<h2 class="baslik">Giriş</h2>
		<div class="konu">
		<h3>Formlar</h3>
		<p>Kullanıcımızla iletişim kurmak istediğimizde bazen mesaj 
		kutuları veya inputboxlar yetersiz kalabilmekte. İşte o zaman 
		UserFormları kullanma zamanı gelmiştir. Tabi, bunların tek amacı iletişim 
		kurmak değil. Daha genel amaçlı işler için de kullanılabilir. Örnek 
		olarak;</p>
		<ul>
			<li>
			<p>Veri Kayıt Arayüzleri</p>
			</li>
			<li>
			<p>Grafikler için ek fonksiyonalite</p>
			</li>
			<li>
			<p>Senaryo toolları</p>
			</li>
			<li>
			<p>Dosya Erişim arayüzü(Kokpit arayüzü)</p>
			</li>
			<li>
			<p>Dashboardlar(Kontrol panelleri)</p>
			</li>
			<li>
			<p>v.s</p>
			</li>
		</ul>
		<p>Bu arada her ne kadar öyle olmadığını bilsek de 
		worksheetler de form olarak düşünülebilir, zira bu 'form' üzerine çeşitli 
		form kontrolleri konulabilmektedir. O halde, formları, üzerine kontrol konan ve bu kontrollerle 
		çeşitli etkileşimlerde bulunduğumuz arayüzler olarak düşünebiliriz.</p>
		<h3>Kontroller</h3>
		<p>Programlama camiasına aşina değilseniz, kontrol kelimesini, 
		birşeyleri kontrol etmeye yarayan nesneler olarak düşünebilirsiniz. Ancak 
		bunlar, formlar üzerindeki görsel nesnelerden başka birşey değildir. 
		Button(Tıklanır düğme), Listbox(liste kutusu), Combobox(açılır liste 
		kutusu) vb.</p>
		<p>Bunları, bi makro çalıştırmak, bir hücre grubundaki 
		değerleri liste kutusuna aktarmak, liste kutusundaki değerleri bi hücre 
		grubuna yazdırmak, bir hücreyi seçmek, bir hücrenin değerini 1'er 1'er 
		artırıp azaltmak gibi&nbsp; amaçlarla kullanırız. </p>
		<p>Kontrollerin detaylarına bir sonraki sayfada değineceğiz. 
		Aşağıda ise bazısını küçük bir örnekte kullanacağız.</p>
	<h3> Form Eventleri(Olayları)</h3>
		<p> Buraya kadar sırayla okuyarak geldiyseniz, olaylar hakkında bilgi 
		sahibi olmuşsunuz demektir. Eğer bilginiz yoksa,
		<a href="Olaylar_OlaylaraGenelBakis.aspx">şuradan</a> temel bilgileri 
		aldıktan sonra tekrar buraya gelmenizi tavsiye ederim.</p>
		<p> Formların kendisi dahil olmak üzere tüm kontrollerin kendine has 
		eventleri vardır. Düğmeye tıklanması için Click eventi, bir listenin 
		güncellenmeden öncesi ve sonrasını gösteren BeforeUpdate, AfterUpdate 
		eventleri gibi. Bunların detaylarnı yine sonraki sayfada ele alacağız.</p>
		<p> <span class="dikkat">Dikkat</span>:Eventler konusunda gördüğümüz <strong>
		Application.EnableEvents=False </strong>atama işleminin Userform ve Kontrol 
		eventleri üzerinde bir etkisi yoktur.</p>
			</div>
			
			<h2 class="baslik">Basit bir UserForm oluşturma</h2>
			<div class="konu">
		<p>Şimdiye kadar konuları sırayla takip 
		ettiyseniz çoğunlukla Insert Module diyip kodları standart modüller 
		içine yazdığımızı görmüşsünüzdür. Eventleri ele aldığımız sayfalarda ise sheet ve workbook 
		nesneleri içine de kod yazmıştık. Şimdi ise bir başka nesne olan 
		UserFormların içine kod yazmaya geldi sıra.</p>
				<h3>Formu oluşturma aşamaları</h3>
				<ul>
					<li>Kendinize yeni bir dosya yaratın. </li>
					<li>VBE'e geçin ve Project 
		peceresinde Modüllere sağ tıklayıp Insert diyin, sonra da 
		<strong>Userform </strong>seçin.</li>
					<li>Karşımıza içi boş bir form gelecektir ve Control ToolBox otomatikman 
		açılacaktır, açılmazsa menüden kendiniz açın. Sonra da toolboxtan aşağıdaki nesneleri formun üzerine sürükleyin ve 
		bırakın: Bi tane commandbutton, bi label, bi textbox, bir de combobox.</li>
				</ul>
		<img alt="UserForm1"  src="../../images/vbauserformtemel1.jpg" >
		<p>
		Hemen bu noktada isimlendirme standardından bahsedeyim. Properties 
		penceresinden, Form nesnenisinin <strong>Name </strong>
		özelliğini <strong>frm</strong>Deneme olmak üzere diğer 
		nesneleri de sırayala <strong>cmd</strong>Run, <strong>lbl</strong>Mesaj, 
		<strong>txt</strong>Mesaj, <strong>cb</strong>Yıl olarak 
		değiştirin. (Formun özelliklerine erişebilmek için formda boş bir yere 
		tıklamanız yeterlidir.) Böylece kod yazarken nesnelere daha kolay referansta 
		bulunabilirsiniz, özellikle formunuzun üzerinde birçok kontrol 
		olacaksa. Buradaki <strong>standart </strong>şudur: Kontrolün tipinin 2-3 karakterlik bir 
		kısaltması, sonra da anlamlı bir isim.</p>
		<p>
		Ayrıca Form'un, CommandButton'un ve Label'ın <strong>Caption </strong>özellikleriyle 
		TextBox ve ComboBox'ın <strong>Text </strong>özelliklerine de anlamlı bir ifade verelim.</p>
		<img alt="" src="../../images/vbauserformtemel2.jpg" >
		<p>
		Farkettiyseniz, birçok form ve web uygulamasında olduğu gibi 
		<strong>grileştirilmiş</strong> metinle kullanıcıya "ipucu/talimat/açıklama içeren mesaj verme" tekniğini 
		kullandım.</p>
		<h3>
		Formu Çalıştırma</h3>
		<p>
		Formunuzun o an itibarıyle Canlı'da nasıl göründüğünü görmek için 
		<strong>Form seçiliyken </strong>F5 vey yeşil Play tuşuna basabilirsiniz. Bu yöntem 
		developer olan kişinin Form çalıştırma yöntemidir. (Şuan bunu yapabilirsiniz ama bu 
		haliyle bir işe yaramayacaktır. Birazdan formumuz biraz daha işlevsel 
		hale getireceğiz.)</p>
		<p>
		Ama öncesinde formu nihai kullanıcının çalıştıracağı yöntemlere bakalım. 
		Kullanıcılar,</p>
				<ul>
					<li>Bir Worksheet butonuna</li>
					<li>Bir ActiveX butonuna</li>
					<li>Bir Add-in'deki butona </li>
					<li>Ribbona/QAT'ye yerleştirilen bir makro butonu<span style="mso-spacerun: yes">na</span></li>
				</ul>
				<p>bastıklarında Formlar açılırlar. 
		Biz basit ve sık kullanılan bir yöntem olması adına bir Worksheet 
		butonuna tıklandığında aktive edecek kodu yazalım. ActiveX'teki mantık 
				da aynı olacaktır. Sonraki sayfada ActiveX kontrollerin detayını 
				göreceğiz. </p>
		<p>
		Şimdi sayfamıza bir buton ekleyelim. Ekleyince çıkan dialog kutusunda 
		New'e tıklayalım ve otomatik açılan modüle aşağıdaki kodu yazalım</p>
		<pre class="brush:vb">Sub Button1_Click()
   frmDeneme.Show
End Sub</pre>
				<p>Evet düğmemizin adını değiştirdikten sonra tıklayalım ve 
				Formumuzu açalım.</p>
<p><img alt="FormAç" src="../../images/vbauserformtemel3.jpg"></p>
		<h4>
		Modal vs Modeless seçenekleri ile sayfa erişimi</h4>
		<p>
		Form bu şekilde açıldığında arkadaki sayfaya erişimimiz engellenmiştir. Sayfayla 
		Form 
		arasında serbest geçiş yapabilmek istiyorsam formu ya <strong>Design 
		aşamasında </strong>properties'ten Modeless<strong>(ShowModal=False)
		</strong>tanımlarım ya da Button1'e tıkladığımda <strong>Runtime 
		sırasında </strong>Modeless açılmasını sağlarım.</p>
		<h5>
		Design sırasında</h5>
		<p>
		<img src="../../images/vbauserformmodeless.jpg"></p>
		<h5>Runtime sırasında</h5>
		<pre class="brush:vb">
Sub Button1_Click()
	frmDeneme.Show vbModeless
End Sub</pre>
		<p>
		Design modundayken <strong>Modeless </strong>tanımlanmış bir formu duruma göre modal 
		açmak için ise aşağıdaki kodu yazarız.</p>
		<pre class="brush:vb">
Sub Button1_Click()
	frmDeneme.Show vbModal
End Sub</pre>
<h3>
		UserForm başlangıç ayarları</h3>
		<p>
		Bir form ilk açıldığında çeşitli başlangıç ayarları yapmak iyi bir 
		fikirdir. Bunu formların <span class="keywordler">Initialize </span>eventi ile yapıyoruz. Web sitesi 
		tasarlayanlar bilir, bu biraz javascripitin <strong>onload</strong> veya 
		ASP'nin <strong>Page_Load 
		</strong>veya .Net formlarındaki <strong>Form_Load</strong> eventlerine benzer.</p>
		<p>
		Bunun için forma sağ tıklayıp <strong>View Code </strong>deyin. Açılan kod sayfasında 
		formu seçince ilk etapta click eventi gelecektir. Siz <strong>Initialize 
		</strong>eventini seçip Click eventini de silin.</p>
		<img alt="" src="../../images/vbauserformtemel5.jpg">
		<p>
		Mesela, formumuz açıldığında cbYıl combobox'ını dolduralım. Şimdi 
		aşağıdaki kodu yazalım.(Detay açıklamalara sonraki Kontroller sayfasında 
		gireceğiz)</p>
		<pre class="brush:vb">Private Sub UserForm_Initialize()
&nbsp;&nbsp;&nbsp; cbYıllar.List = Array("2017", "2018", "2019")
End Sub</pre>
		<p>
		ComboBox'tan seçim yapılacağı sırada da gri olan bilgilendirme yazısı yok olup 
		yılların rengi de siyaha dönsün istiyoruz diyelim.</p>
		<pre class="brush:vb">Private Sub cbYıllar_DropButtonClick()
&nbsp;&nbsp;&nbsp; cbYıllar.ForeColor = vbBlack
&nbsp;&nbsp;&nbsp; cbYıllar.Text = ""
End Sub</pre>
		<p>Bunun dışında yine çeşitli başlangıç değeri atamaları, 
		kutuların temizlenmesi, varsa statik değişken tanımlamaları v.s bu event 
		içinde yapılabilir.
		</p>
				<h3>Formları Gizleme ve Kapama</h3>
				<p>Bazen bir formu geçici olarak gizlemek bazen de tamamen 
				kapatıp başka bir form açmak isteriz. Bunlar için ihtiyacımız 
				olan kodlar şöyle:</p>
				<pre class="brush:vb">Me.Hide 'o an aktif olan formu gizler
Form1.Hide 'Form1i gizler

Unload Me 'aktif formu kapatır
Unload Form1 'Form1'i kapatır</pre>
				<p>Bir formu kapatmak için sağ üstteki X düğmesine de 
				basılabilir tabiki ama <span class="keywordler">Unload</span> 
				fonksiyonu daha çok kapatma işleminin arkasından başka bir iş(ler) yapmak(mesela başka 
				bir formu açmak) istediğimiz zamanlarda kullanılır.</p>
				<p>Gizlediğimiz bir formu tekrar aktive etmek için yine
				<span class="keywordler">Show</span> metodu kullanılır, yani bir
				<strong>Load</strong> fonksiyonu bulunmamaktadır. Özetle 
				elimizdekiler şöyle: İlk kez açma ve yeniden gösterme için
				<strong>Show</strong> metodu, gizleme için <strong>Hide</strong> 
				metodu, kapatma için <strong>Unload</strong> fonksiyonu.</p>
				<h4>Esc tuşu ile çıkış</h4>
				<p><span>ESC tuşuyla çıkış yapmak isterseniz, Form üzerinde bir 
				buton koyun ve bunun <strong>Cancel</strong> özelliğine <strong>
				True</strong> atayın. Bu sayede Esc tuşuna basıldığında bu düğme 
				odağı almış olur, yani sanki seçilmiş gibi olur, ki bu da Enter 
				eventini tetikler. Şimdi ikinci olarak bu düğmenin Enter 
				eventine formu kapatan kodu yazalım.</span></p>
				<pre class="brush:vb">Private Sub CommandButton3_Enter()
  Unload Me
End Sub</pre>
				<p><strong>NOT: </strong>Modal formlarda, bir form ilk kez 
				açıldığında önce <strong>Initialize</strong> sonra <strong>
				Activate</strong> eventleri meydana gelir. Sonra bu form 
				gizlenip tekrar açıldığında sadece Activate meydana gelir. 
				Modeless formlarda ise, Initializedan sonra Activate meydana 
				gelmez, iki modeless forma arasında gidip gelince veya gizli 
				olan bir modeless form tekrar gösterildiğinde meydana gelir. Eğer 
				ki yeniden aktive olma durumlarına yaptırmak istediğiniz bir 
				işlem varsa, bu ayrıma dikkat etmelisiniz.</p>
		<hr>
				<p>Evet, Formlara kısa bir giriş yaptıktan sonra artık formlar 
				üzerindeki kontrollerin detaylı kullanımına geçebiliriz.</p>
		</div>

</asp:Content>
