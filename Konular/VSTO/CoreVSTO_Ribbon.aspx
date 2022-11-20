<%@ Page Title='Ribbon' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'>
    <div id='gizliforkonu'>
        <table>
            <tr>
                <td>
                    <asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td>
                <td>
                    <asp:Label ID='Label2' runat='server' Text='Görsel Araçlar'></asp:Label></td>
                <td>
                    <asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td>
            </tr>
        </table>
    </div>
    <h1>Ribbon</h1>
    <p>Öncelikle şunu söylemekte faydar var. VS, VSTO kodlaması adına birçok işi bizim yerimize yapmış olacak. Bizim bilmemiz gereken 3 temel şey var</p>
    <ul>
        <li>Excel object model&#39;e nasıl erişilir</li>
        <li>.Net kodu nasıl yazılır</li>
        <li>Ribbon yönetimi nasıl yapılır</li>
    </ul>
    <p>
        Excel Object Modelle uğraşmayı bi önceki sayfada gördük.</p>
    <p>
        .Net çok büyük bir dünya. Aynı anda hem VS&#39;yu kullanmayı öğrenmeniz hem de c#(veya Vb.Net) kullanmayı öğrenmeniz gerekecek. Bu uzun bir yolculuk olacak. Ben başlangıç için gerekli olan kısmı bu sitede vereceğim, kendinizi daha ileri götürmek size kalmış.</p>
    <p>
        Geriye bence en önemli kısım, çalışmanızın vitrini olan Ribbon yönetimi kalıyor. İşte bu sayfada bu konuya detaylarıyla(başka hiçbir kaynakta olmayan detaylarıyla) bakacağız.</p>
    <h2 class='baslik'>Sekmeler</h2>
    <div class='konu'>
        <p>Ribbon yaratımını<a href="VSTOOzel_IlkVSTOAddin.aspx#ribbon"> ilk vsto addin</a> sayfasında görmüştük. Kaldığımız yerden devam edelim. Yalnız devam etmeden önce küçük bir bilgilendirme yapmak isterim. Biz, ribbonumuzu Visual Designer modunda oluşturmuştuk ve bu yöntemin XML yöntemine göre daha basit bir fonksiyonalite sunduğundan bahsetmiştik. <a href="https://social.msdn.microsoft.com/Forums/vstudio/en-US/e3a68e06-9e27-4d6c-bd1e-e566ab8b7506/ribbon-xml-vs-ribbon-designer?forum=vsto">MSDN&#39;de</a>, üstatlardan Cindy Meister&#39;ın verdiği cevapta detayları görebilirsiniz.</p>
        <h3>Sekmelerin konumu</h3>
        <p>Hatırlayacak olursak Ribbonumuz Add-in menüsüne yerleşmişti.</p>
        <p><img alt="" src="/images/vsto_ribbonshow.jpg" /></p>
        <p>Halbuki ribbonun yerleşimi için daha şık olan iki seçenek daha var:</p>
        <ul>
            <li>Bağımsız bir sekme. Ör:Excel menülerinden bambaşka bir fonksiyonaliteye sahipse, veya ticari amacı olan bir çalışma ise.</li>
            <li>Mevcut sekmelerden birinin içine. Ör:Add-indeki kodlarımızı temsil eden ribbonumuz veri ağırlıklı bir çalışma ise Data menüsüne yerleştirebiliriz.</li>
        </ul>
        <p>
            Şimdi VS ekranındaki halini tekrar bi hatırlayalım.</p>
        <p>
             <img alt="" src="/images/vsto_ribbonadd2.jpg" /></p>
        <p>
            Gördüğünüz üzere ilk yaratımda <strong>TabAddIns(Built-In)</strong> şeklinde isimlendirildi. 
            Bunu şöyle okumak lazım. Excel&#39;in built -in tablarından(sekmelerinden) Add-in&#39;s sekmesi içinde görünecek. Şimdi toolboxtan yeni bir Tab elemanını ribbona sürüklersek bunda parantez içinde <strong>built-in </strong>yazmadığnı görürüz. Bu, yeni sekme ayrı bir sekme şeklinde görünecek demektir. Bakalım gerçekten öyle mi, projemizi çalıştıralım ve bakalım:</p>
        <p>
             <img alt="" src="/images/vsto_ribbonnewtab.jpg" /></p>
        <p>
            Bu şekilde, tab2 için yukarıdaki iki seçenekten ilkini uyarlamış olduk, yani bağımsız bir sekme yaptık. Şimdi diyelim ki biz bu sekmeyi Data menüsü içinde görmek istiyoruz. Bunun için ilgili sekmeyi seçtikten sonra Properties menüsünden&nbsp;<strong>ControlIdType</strong> değerini “<strong>Custom</strong>”dan “<strong>Office</strong>”e dönüştürelim.</p>
        <p>
            tab2&#39;nin ilk hali böyle iken,</p>
        <p>
            <img alt="" src="/images/vsto_tabkonum1.jpg" /></p>
        <p>
            Şimdi şöyle olur, bu arada <strong>OfficeId</strong> seçeneğine de hangi built-in menüde görüneckese onu yazarız; data için <strong>TabData</strong>.</p>
        <p>
            <img alt="" src="/images/vsto_tabkonum2.jpg" /></p>
        <p>Sonuçta designerda sekmelerimiz aşağıdaki gibi görünür.</p>
        <p>
            <img alt="" src="/images/vsto_tabkonum3.jpg" /></p>
        <p>
            Şimdi bunun içine bir <strong>Group</strong> kontrolü koyup, içine de bir iki buton koyalım ve çalıştırıp Excel&#39;de nasıl göründüğne bakalım.</p>
        <p>
            <img alt="" src="/images/vsto_tabkonum4.jpg" /></p>
        <p>
            Gördüğünüz gibi, her ne kadar bu sekmenin adı tab2 idiyse de, tab2 ifadesini herhangi bir yerde görmüyorsunuz, zira bu sekmemiz Data sekmesi içine yedirilmiş oldu.</p>
        <h4>
            Built-in menüleri gizleme</h4>
        <p>
            Excel&#39;in built-in menüleri görünmesin, sadece kendi yazdığımız ribbon sekmesi görünsün istiyosak, Ribbon nesnesinin <strong>StartFromScratch</strong> seçeneği True yapılır.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonscratch.jpg" /></p>
        <h3>
            File menüsü</h3>
        <p>
            Çok ihtiyacınız olur mu bilemedim ancak istenirse File menüsüne bile bir buton ekleyebiliyorsunuz. Mesela aşağıdaki gibi genel ribbon ayarlarınızı içeren formu açan bir düğme koyabilirsiniz. Gerçi ben bunu ana sekmemin en sonuna koymayı tercih ederdim. Tercih sizin.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonfilemenu.jpg" /></p>

        </div>


        <h2 class="baslik">Ribbon özelleştirme</h2>
    <div class="konu">
        <h3>
            Ribbon Controlleri</h3>
        <h4>
            Kapsayıcılar(Container elementler)</h4>
        <p>
            Ana kapsayıcımız Group controlüydü, bunu yukarıda gördük, 
            tüm controlleri bunların içine koymak durumundayız. Bunun dışında bir de alt kapsayıcılar var, bunlar Group nesnesi içindeki contolleri gruplamaya yararlar. Özellikle mantıksal(logical) bir gruplama amacıyla kullanılırlar.</p>
        <p>
            İki tür kapsayıcı kontrol var. Biri <strong>Buttongroup</strong>, diğeri <strong>Box</strong>. <strong>Buttongroup</strong> sadece button benzeri controlleri(button, toggle buton, split button) ve içine buton konulan popup menüleri(menü ve gallery) alır ve bunları dizerken, <strong>Box</strong> controlü ise her tür controlü içine alır ve<strong> hem yatay hem dikey</strong> hizalayabilir. Aşağıdakilerden kırmızı olan Buttongroup, mavi olanlar Box controlüdür. Box&#39;ın BoxStyle propertysi ile yatay/dikey ayrımı belirlenir. Bu arada her ne kadar design görünümünde etraflarında bi border çizigisi varmış gibi görünse de derlendiğinde bu çizgiler görünmez, o yüzden bi alttaki separatorleri kullanmak gerekebilir.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonboxandgroup.jpg" /></p>
        <p>
            Box nesnesinin Items adlı bir collectionı vardır. Bu demektir ki, içindeki controlleri döngüyle dolaşabilirsiniz. İhtiyacınız olur diye söylüyorum.</p>
        <h4>
            Seperator</h4>
        <p>
            Anlamı açık diye tahmin ediyorum. Kontroller arasına ayraç koyar. Böylece aynı group içindeki kontrollerin birbirine çok yaklaşması engellenmiş olur. Yalnız, bazen design görünümünde taşma v.s görünse bile proje derlendiğinde normal görünüm alabiliyor. O yüzden nihai görünümü, çalıştırmadan bilemezsiniz diyebiliriz. </p>
        <p><img alt="" src="/images/vsto_ribbonseperator.jpg" />
            &nbsp;</p>
        <h4>Toggle button</h4>
        <p>VBA&#39;den biliyorsunuz.</p>
        <h4>Label ve EditBox</h4>
        <p>Label&#39;ı açıklamaya gerek yok, VBA&#39;den biliyorsunuz. EditBox da bildiğimiz TextBox aslında. VBA&#39;de add-in&#39;ler ve menüler konusunda da bahsetmiştim, bir textboxın(editboxın) ribbonda/menüde olması çok pratik bi uygulama değil bence. O yüzden şimdilik es geçiyorum. </p>
        <h4>Checkbox</h4>
        <p>Checkbox da VBA&#39;deki gibi, açıklamaya pek gerek yok, gerçi pratikte bir kullanımı olur mu ondan da emin değilim. Bu arada checkbox&#39;lara benzeyen radiobutton&#39;lar Ribbonlarda kullanılamıyor.</p>
        <h4>Pop-up menüler</h4>
        <h5>Menü</h5>
        <p>Menülere VBA&#39;den aşinayız ama yine de nasıl kullanılacağını bi görelim.</p>
        <p>Bunun için burdan itibaren <strong>ilk tab</strong>&#39;ımızı kullanmaya başlayalım. Buna yeni bir grup ekleyelim ve adına Dosyalar diyelim. Buraya sık kullandığımız bazı dosyaları koyacağız. Normalde zaten&nbsp; Windows&#39;un böyle bir özelliği var, mesela Excele sağ tıkladığınızda sık kulalndıklarınızı üst tarafa pinleyebiliyorsunuz ama biz böyle bir özellik olmadığını varsayalım ve sık kullanılan dosyalara ait bir menü yapalım. Bu arada buraya farklı dosya tiplerini de ekleyebilirsiniz(word, pdf, ppt v.s). Hatta ileride göreceğimiz gibi bunlar başka programları çalıştıran butonlar da olabilir.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonmenu1.jpg" /></p>
        <p>Bunların click eventine de ilgili dosyaları açan kodu yazarız.</p>
        <pre class="brush:csharp">
private void button13_Click(object sender, RibbonControlEventArgs e)
{
    string adres = @"E:\OneDrive\Dökümanlar\bütçem.xlsx";
    MyStatik.app.Workbooks.Open(adres);
}            &nbsp;</pre>
        <p>Menü içine farklı resim-yazı kombinasyonunda butonlar ve başka menüler de konabilir.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonmenu2.jpg" /></p>
        <p>Burada karşımızda birkaç problem durabilir</p>
        <ul>
            <li>10 tane dosya varsa, 10&#39;u için de ayrı ayrı mı click eventi gireceğiz.(Hayır tabiki, detaylar için <a href="OrnekProjeler_Excelent.aspx">Excelent</a> proje incelemesinde)</li>
            <li>Dosyaların bazısı Excel dosyası değilse?(Teoriği <a href="Netislemleri_TemelNetislemleri.aspx">.Net işlemlerinde</a>, pratiği Excelent projesinde)</li>
            <li>Biz şuan sabit bir dosya listesi girdik, ya Kullanıcılar bunları kendi isteği gibi özelleştirmek isterse? (Aşağıda dialog launcher bölümünde settings işlemlerini yapmayı göreceğiz, açıklamalı örneği ise yine Excelent Projesinde göreceğiz)</li>
        </ul>
        <p>
            <strong>ItemSize</strong> özelliği ile menünün içindeki elemanların boyutunun küçük mü büyük mü olacağı belirlenir. Diğer özellikleri aşağıdaki ortak özellikler içinde göreceğiz.</p>
        <p>
            NOT: Bu arada butonlara atadığınız resimler/iconlar design modda görünmez. Bunların nasıl göründüğünü test etmek için projeyi çalıştırmanız gerekiyor. Ayrıca bazen buton orada olduğu halde yokmuş gibi de görünebiliyor. Bu tür durumlarda Ribbon&#39;u kapatıp açın veya komple projeyi kapatıp açın.</p>
        <h5>Gallery</h5>
        <p>Menülere göre daha şık bir gruplama sağlarlar. Aşağıda Excel&#39;in kendi menülerinden birinde Gallery kullanımı görüyoruz.</p>
        <p>
            <img alt="" src="/images/vsto_ribbongallery.jpg" /></p>
        <p>Gallery içine buton sürüklemek yerine, properties&#39;ten Items collection&#39;ına elemanlar ekliyoruz. Sonra da bunların resimleri için değer gireceğiz. Bunları az aşağıda ortak özellikler bölümünde göreceğiz. Biz şimdi resim eklemeden kodumuzu yazalım. Burada farklı butonlar olmadığı için aslında tek bir event var, o da gallerynin tıklanma eventi.</p>
        <p>Yapacağımız şey, seçilen elemanını indeksine bakmak olacaktır.</p>
        <pre class="brush:csharp">private void gallery1_Click(object sender, RibbonControlEventArgs e)
{
    switch (this.gallery1.SelectedItemIndex)
    {
        case 0:
            MessageBox.Show("ilk item seçildi");
            break;
        case 1:
            MessageBox.Show("ikinci item sçeildi");
            break;
        default:
            MessageBox.Show("Buraya gelmemeli");
            break;
    }
}</pre>
        <p>Burada ayrıca switch-case yapısını da görmüş olduk. VBA&#39;deki seleckt case yapısının aynısıdır. Kullanımındaki küçük ayrımı keşfetmeyi size bırakıyorum.</p>
        <p><strong>ItemImageSize</strong> özelliği ile itemların ebatları rakam verilerek belirtilir. Menülerde ise küçük/büyük seçenekleri vardı.</p>
        <h5>ComboBox ve DropDown</h5>
        <p>ComboBox&#39;ların ne olduğunu biliyoruz. Dropdown&#39;lar bunlara çok benzer; tek farkı, Dropdownlara tıkladığımızda metnin içini seç(e)meyiz, Combolarda ise metni seçmiş oluruz ve gerekirse elle bişeyler yazabiliriz. Ben dropdownları seviyorum, neresine tıklarsam tıklayayım seçenekler geliyor, comboda ise metinli bölgeye tıklarsanız oraya girmiş oluyorsunuz, seçimlerin açılması için illa yandaki oka tıklamak gerekiyor.</p>
        <p>
            <img alt="" src="/images/vsto_ribbondropcombo.jpg" /></p>
        <p>
            Bu anlamda
            dropdownlar gallery&#39;lere benziyor. Aradaki fark, dropdowndakiler hep altalta bir liste şeklinde açılırken galeryde ise bi grid(ızgara) şeklinde yerleşim sözkonusu.</p>
        <h5>
            Splitbutton</h5>
        <p>
            ButtonType=Button atanırsa aşağıdakilerden ilki gibi, Toggle seçilirse ortadaki gibi görünür. 
            Diğer popuplardan farkı şu: Diğerlerinde seçeneklerden birini seçmek için açılır kutuya basmamız&nbsp;gerekir, bunda ise görünen metnin kendisi de bir seçenektir, görünür metin default seçenek olup buraya genelde en olası seçenek yazılır, böylece hiç açılır kısmı açmadan doğrudan seçim yapılabilir. Diğer seçenekler için açılır kutuyu açmak gerekir. Dolayısıyla splitbuttonun ve içindeki butonların eventleri birbirinden ayrıdır(split ifadesi bundan dolayıdır).</p>
        <p>
            &nbsp;</p>
        <p>
            <img alt="" src="/images/vsto_ribbonsplittoggle.jpg" /></p>
        <p>
            Bu arada Excel&#39;in built-in menülerinden örnek vermek gerekirse, Home menüsündeki Paste butonu bi Splitbuttondur. Doğrunda Paste&#39;e tıkladığınızda bir seçimi olduğu gibi yapıştırır, altındaki açılır kutuya tıklayıp başka bir seçim de yapabilirsiniz.</p>

        <h3>Resource Ekleme</h3>
        <p>Gerek ribbondaki controllere, gerek formlarımızda kullanacağımız resimleri projenin içine dahil etmek istiyorsak aşağıdaki adımları ekleyerek ilerelyebilriz.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonresources.jpg" /></p>
        <p>
            Buraya icon dosyaları(ico), ses ve video dosyaları da ekleyebiliyorsunuz. </p>
        <p>
            Sonrasında, eklediğimiz dosyaları herhangi bir kontrolün Image properties&#39;ine kolaylıkla ekleyebiliyoruz.</p>
        <p>
            <img alt="" src="/images/vsto_resourceadd.jpg" /></p>
        <p>
            Hatta bu imajlara daha sonra runtime sırasında aşağıdaki gibi de ulaşabiliyoruz.</p>
        <p>
            <img alt="" src="/images/vsto_resourceprop.jpg" /></p>
        <p>
            Bu arada ico uzantılı dosyaları Image porperty&#39;sine atamak için aşağıdaki gibi bir kod yazmak gerekiyor. 
            Resource&#39;dan ico dosyasını(_0.ico) seçiyoruz, sonra onu Bitmap&#39;e çeviriyoruz. Ben bunu Ribbon_Load eventine yazdım, böylece ribbon yüklenir yüklenmez imaj ataması yapılmış oluyor.</p>
        <pre class="brush:csharp">
button20.Image = Properties.Resources._0.ToBitmap(); </pre>
        <p>
            Vb.Net&#39;te bu kısım biraz farklı yazılıyor.</p>
        <pre class="brush:vbnet">
button20.Image = My.Resources._0.ToBitmap()</pre>
        
        <h3>Ortak Properties</h3>
        <p>Burda birçok nesnenin ortak property&#39;sine bakacağız.</p>
        <h4>ControlSize</h4>
        <p>
            İlgili controlün küçük mü yoksa büyük mü görüneceğini belirler. Hali hazırda Excel&#39;de iki türü de görüyorsunuz. Mesela Home menüsündeki Paste butonu <strong>Large</strong> bir controldür. Hemen yanındaki Cut ise <strong>Regular(küçük) </strong>bir controldür. Yalnız dikkat, bir control ButtonGroup veya bir Pop-up menü içindeyken bu özellik görünmez(erişilebilir değildir) ve mecburen küçük(regular) olur. Bu boyutu, ilgili kontrol için seçilen resmin pixelinden bağımsız olarak ekranda kapladığı alan olarak düşünün. Bu yüzden küçük ölçülerde boyutlanmış bir resim Large bir controlün Image property&#39;sine atandığında pixel pixel(tabiri caizse Minecraft karakterleri gibi) görünürken, büyük boyutlu bir resim ise küçük(regular) bir controle atandığında buruşuk(pixellerin üstüste binmesinden dolayı) görünecektir.</p>
        <h4>Image, OfficeImageId, ShowImage, ShowLabel</h4>
        <p>Bu 4 özellik birbiriyle bağıntılı olduğu için bir arada aldım.</p>
        <p>
            Bir control için görsel bir etiket istiyorsak, bunun için <strong>Image </strong>özelliği için bir resim seçip atarız veya <strong>OfficeImageId </strong>özelliğine bir değer atarız. OfficeImageId değerlerini internette bulabilirsiniz. Gerçi Microsoft&#39;un kendi sitesindeki link hata veriyor, belki ileride düzeltirler diye ben yine de bu <a href="http://go.microsoft.com/fwlink/?LinkID=220760">linki</a> buraya koyuyorum. Alternatif siteler ise şöyle:</p>
        <ul>
            <li><a href="https://www.spreadsheet1.com/office-excel-ribbon-imagemso-icons-gallery-page-01.html">https://www.spreadsheet1.com/office-excel-ribbon-imagemso-icons-gallery-page-01.html</a></li>
            <li><a href="https://bert-toolkit.com/imagemso-list.html">https://bert-toolkit.com/imagemso-list.html</a></li>
        </ul>
        <p>
            Bir şekilde hem Image hem OfficeImageId belirlediyseniz Image baz alınır.</p>
        <p>
            <strong>ShowLabel</strong>: Controlün etiketi görünecek mi görünmeyecek mi, bunu belirler.&nbsp; ControlSize, Large set edilmişken False atanamıyor.</p>
        <p><strong>ShowImage</strong>:Control'e atanmış resmin gösterilp gösterilmeyeceği belirlenir. ControlSize'da olduğu gibi, bir control bir buttongroup veya menü içindeyken bu özellik erişlebilir durumda değildir.</p>
        <p>Pop-up menülerin bir de <strong>ShowItemLabel </strong>ve <strong>ShowItemImage </strong>özellikleri var ki, bunlar içindeki elemanlar için geçerlidir.</p>
        <p>Image değeri verilmişken ShowImage=False denirse resim görünmez. OfficeImageID vermişseniz ShowImage=False yapamazsınız. Bunlarla oynayıp denemeniz gerek. Çeşitli kombinasyonlar çıkabiliyor, resimli-yazısız, resimli-yazılı, resimsiz-yazılı, küçük resimli, büyük resimli v.s</p>
        <p>Aşağıda benim çeşitli denemelerimi görebilrisiniz.</p>
        <p>
            <img alt="" src="/images/vsto_ribbongallery2.jpg" class="zoomla" /></p>
        <h4>Screentip ve Supertip</h4>
        <p>Sırayla, bir controlün üstüne gelindiğindeki başlık ve detay bilgileri verdiğiniz propertylerdir.</p>
        <p>
            <img alt="" src="/images/vsto_ribbontip.jpg" /></p>
        <p>Gallery itemlarında bunlar her item için ayrı ayrı ayarlanabiliyor.</p>
        <p>Bu arada bu screentip içine resim de eklenebiliyor ancak bunu yapmak için ribbonu Visual Designer ile değil Xml olarak yaratmamız lazımdı. Aradaki fark için en başta yazdığım açıklamaya tekrar bakın.</p>
        <h3>
            Dialog Launcher ekleme</h3>
        <p>
            Aşağıdaki Home menüsünün köşesindeki gibi küçük butonlara Dialog Launcher deniliyor. Bunlara tıklandığında ya bir dialog kutusu çıkar veya bir Taskpane, ve bunlarda çeşitli ayarlar yapılır. Bu ayarların bir kısmı tek seferliktir, bir kısmı ise File&gt;Options&#39;taki gibi kalıcı ayarlardır. Şimdi bunlar nasıl yaplır, ona bakaczğız.</p>
                <p>
            <img alt="" src="/images/vsto_ribbondialoglauncher0.jpg" /></p>
        <p>
            Ribbonumuzda tab2 sekmemizinde ilk group nesnesini(group2) seçtikten sonra properties&#39;ten aşağıdaki gibi Add DialogLauncher diyelim.</p>
        <p>
            <img alt="" src="/images/vsto_ribbondialoglauncher1.jpg" /></p>
        <p>
            Buna tıklayınca group nesnesinin köşesine bu butoncuk eklenir. Şimdi sıra, buna tıklandığında ne yapılacağını belirlemeye geldi. Aşağıdaki gibi yıldırım butonuna tıklayarak açılan eventler kısmı boşken, çift tıklayınca otomatik bir event handler yaratılır ve kod sayfasında bunu size gösterir. </p>
        <p>
            <img alt="" src="/images/vsto_ribbondialoglauncher2.jpg" /></p>
        <p>
            Nasıl göründüğüne bakalım:</p>
        <p>
            <img alt="" src="/images/vsto_ribbondialoglauncher3.jpg" /></p>
        <p>
            Şimdi frmSettings1 adında bir form oluşturalım ve dialog eventimize aşağıdaki kodu yazalım</p>
        <pre class="brush:csharp">
private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)
{
    frmSetting1 f = new frmSetting1();
    f.Show();
}   </pre>
        
        <h3 id="settings">
            Settings formları olşuturma</h3>
        <p>
            Dialog Launcher&#39;da açtığımız settings formunu düzenleyeceğiz.</p>
        <p>
            Solution Explorer&#39;da Properties&#39;e çift tıklayalım.(VB.Net&#39;te MyProject) ve açılan pencerede Settings&#39;e gelelim. Şimdi iki ayar gireceğiz. Bunlardan biri bazı makrolar için kullancıdan kullanıcıya değişebilecek klasor adresi girmek, bir diğeri de yeni bir dosya açıldığında kaç sayfa açılacak, bunu kullanıcının belirlemesine izin vermek, bunu taibiki <strong>Files&gt;Options&#39;</strong>tan da yapabiliyoruz ancak maksat pratik olsun. Bu arada VBA&#39;de iken bu tür ayarları txt dosyalarına yazıp okuyarak hallediyorduk, VSTO&#39;da ise bu şekilde hem daha pratik hallediyoruz hem de şifre tarzı bilgileri de güvenlik altına almış oluyoruz. </p>
        <p>
            Şimdi aşağıdaki iki ayarı girelim.(İlki için siz de uygun klasörü girin)</p>
         <p>
            <img alt="" src="/images/vsto_ribbonsettings.jpg" /></p>
        <p>
            Şimdi Settings formumuz açılıdığnda ayarlarda kayıtlı ne varsa onu form içindeki textboxlarda gösterelim. Bunun için formda iki kutu yaratalım.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonsettings2.jpg" /></p>
        <p>
            Bu formun Load eventine aşağıdaki kodu yazarak ayarları alıyoruz.</p>
        <pre class="brush:csharp">
private void frmSetting1_Load(object sender, EventArgs e)
{
    this.txtKlasor.Text = Properties.Settings.Default.Anaklasor;
    this.txtSheetAded.Text = Properties.Settings.Default.YeniWbSheetAdedi.ToString();
}  </pre>
        <p>
            Kullanıcı ayarları değiştirip kaydetmek isteyecektir. Bunun için bir kaydet butonu yaratmak yerine tüm MS Office uygulamlarının da yaptığı gibi, form kapandığında otomatik kayıt yapma mantığıyla hareket edelim. O yüzden formun eventlerine gidip FormClasing&#39;e çift tıklayarak ilgili event handlerı oluşturuyoruz ve aşağıdaki kodu yazıyoruz.</p>
        <pre class="brush:csharp">
private void frmSetting1_FormClosing(object sender, FormClosingEventArgs e)
{
    Properties.Settings.Default.Anaklasor = this.txtKlasor.Text;
    Properties.Settings.Default.YeniWbSheetAdedi = int.Parse(this.txtSheetAded.Text); //parsing, string bir ifadeden sayısal metin çıkarma işlemidir
    Properties.Settings.Default.Save(); //bunu yazmazsak ayarlar kaydolmaz
}</pre>
<p>
    Kodumuzu çalıştıralım, dialog launcherı çalıştıralım, çıkan pencerede değerleri değiştirip, 
    tekrar açtığımızda değişmiş olduğunu görüyüoruz.</p>
        <p>
            <img alt="" src="/images/vsto_ribbonsettings3.jpg" /></p>

<h4>
    Settings kontrollerini özelleştirme</h4>
        <p>
            Bu arada bazı kutuların ise ayarlara kaydolmasını istemez ama kullanıcının yine de ilgili oturum içinde 
            bunların içeriğini değiştirebilmesine izin vermek istersiniz. Kalıcı olmayan bu bilgilerle kalıcı yaptığınız ayarlı olanların birbirinden farklılaşması için ayarlı textboxların kırmızı fontlu olmasını sağlayabilirsiniz. Bunun için her ayarlı kutu için tek tek uğraşmak yerine TextBox sınıfından türetilmiş RenkliTextBox sınıfını yaratabiliriz. Bu kavrama Inheritance(Kalıtım) denir. Çok fazla örneğini görmeyeceğiz ancak özetle mantığı şu: Bir sınıfı baz alarak başka bir sınıf tanımlıyoruz. Böylece o sınıfın bütün özellikleri yeni sınıfımızda hazır olarak bulunuyor, biz sadece değiştirmek istediğimiz kısımları değiştiriyoruz veya eklemeler yapıyoruz. Hadi gelin şimdi de ona bakalım.</p>
        <p>
            Projemize <strong>Add&gt;New</strong> item diyerek <strong>UserControl</strong> ekliyoruz, adına RenkliTextBox diyelim. Açılan pencerede controlün içine bir tane textbox sürükleyelim. Sonraki kod sayfasına geçip constrcutor metod içindeki InitializeComponent satırını altına renklendirme kodumuzu yazalım.</p>

    
        <pre class="brush:csharp">
public RenkliTextBox() //constructor metod
{
    InitializeComponent();
    this.textBox1.ForeColor = Color.Red;
}           &nbsp;</pre>
        <p>
            Projeyi <strong>Rebuild </strong>edelim ve controlümüzün Toolboxa geldiğini görelim.(<strong>Rebuild= Clean + Build</strong>. Bazı işlemler için Rebuild gereklidir, yeni eklenen controllerin toolboxa eklenmesi de onlardan biri)</p>
        <p>
            <img alt="" src="/images/vsto_usercontrol.jpg" /></p>
        <p>
            Bakalım istediğimiz gibi çalışıyor mu?</p>
         <p>
            <img alt="" src="/images/vsto_usercontrol2.jpg" /></p>
        <p>
            &nbsp;</p>

                <h3>Sekmeler arası geçişler</h3>
        <p>Çalışmanızda birden fazla sekme yaratmış olabilirsiniz. Böyle bir durumda hepsini tek seferde göstermek yerine, ilk açılışta ana sekmeyi açıp, butonlar aracılığı ile de diğer sekmeleri açıp kapatabilirsiniz. </p>
        <p>Mesela ana sekmemiz olan tab1&#39;e koyacağımız bir toggle_buton&#39;a aşağıdaki gibi kod yazabiliriz. Bunun için bir sekme daha yaratmanız gerekiyor, yani tab3. Butona tıklıyken tab3 görünecek ve aktif olacak, butona tıklı değilken tab3 gizlenecek.</p>
        <pre class="brush:csharp">
private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
{
    Globals.Ribbons.Ribbon1.tab3.Visible = toggleButton1.Checked; //ribbonu görünür kılıyoruz
    if (toggleButton1.Checked)
    {
        Globals.Ribbons.Ribbon1.RibbonUI.ActivateTab("tab3");//sonra da aktif hale getiriyoruz            
    }            
}        </pre>

        <h3>Ribbonu kaldırma</h3>
        <p>Bir üstte, Ribbon sekmelerini nasıl gösterip gizlediğimizi gördük. Ancak bazen komple Ribbonu kaldırmak isteyebiliriz. Mesela 
            belli bir tarihe geldiğinizde Add-in'in kullanıcılardan kalkmasını isteyebilirsiniz. Bunun için Trial Version&#39;u olan bir proje yapmak da çözümdür ancak bunun daha karmaşık olduğu aşikardır, onun yerine Ribbon&#39;u kaldırarak kullanıcının arayüze dolayıyısla add-ine erişimini engellemiş olursunuz. Aşağıdaki kodu Ribbon_Load veya ThisAddIn_Startup içine yazabilirsiniz.</p>
        <pre class="brush:csharp">
if (DateTime.Today>Convert.ToDateTime("31.12.2019"))
{
    MessageBox.Show("Add'in in süresi dolmuştur.");
    Globals.Ribbons.Ribbon1.Dispose();
}   
        </pre>
    </div>


</asp:Content>
