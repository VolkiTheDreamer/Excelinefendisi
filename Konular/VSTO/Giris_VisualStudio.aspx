<%@ Page Title='Giris VisualStudio' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='3'></asp:Label></td></tr></table></div>
    <h1>Visual Studio</h1>
    <h2 class="baslik">Nedir?</h2>
    <div class='konu'>
        <p> Visual Studio(VS), Microsoft'un yazılım geliştiriciler için sunduğu IDE'dir.(Integrated Development Environment). Burada hem masaüstü, hem web hem mobil uygulama geliştirebiliyoruz demiştik. Hatta bu web sitesini de şuan VS içinde hazırlıyorum.</p>
        <p> İşte VSTO sözkonusu olduğunda, VBA&#39;den farklı olarak kodlarımız Excel içinde değil VS içinde yazacağız. Kodlamayı bitirdiğimizde yine Excel&#39;in içinde görünür şekilde bulamayacaksınız, ki bu aynı zamanda kodlarınızı başkalarından korumak adına iyi de birşey. VBA&#39;e göre daha zahmetli olacağı kesin ancak getireceği avantajlar düşünüldüğünde bu kadar zahmete katlanılır diye düşünüyorum. </p>
        <p> Bu IDE&#39;nin genel görünümü aşağıdaki gibidir.</p>
        <p> 
            <img alt="" src="/images/VSTO_vs1.jpg" class="zoomla" /></p>
        <p> Bi önceki sayfada belirttiğimiz gibi Visual Studio’nun da sürekli yeni versiyonu çıkmaktadır. İşyerindeki program kurulumlarını BT ekibiniz yapıyorsa en güncel ellerinde ne varsa onu kuracaklardır. Siz evde Community versiyonunu <a href="https://visualstudio.microsoft.com/tr/">buradan</a> kurabilirsiniz. Versiyonların sağladıkları imkanlar için <a href="https://visualstudio.microsoft.com/tr/vs/compare/?rr=https%3A%2F%2Fwww.google.com%2F">bu sayfaya</a> (bilahare) bakabilirsiniz.</p>
        </div>

        <h2 class="baslik">Kurulum</h2>
        <div class="konu">
        <h3>Kurulum aşamaları</h3>
        <p> Öncelikle güzel haber: Visual Studio’nun eski versiyonlarında, Community(Express) sürümlerinde doğrudan VSTO yazamıyorduk, bunun için ek başka programlar(Önce "Web Platform Installer", sonra bunun içinden de "Office Developer Tools for VS") yüklemek gerekiyordu. Visual Studio 2017&#39;den itibaren ise direkt VS Installer ekranından kurulacak componentleri seçebiliyorsunuz, ve bunlardan biri de VSTO seçeneğidir.</p>
            <h4> Ara Not</h4>
            <p> Bu noktada şunu belirtmek isterim ki, bazı kişiler garip bir şekilde artık VSTO Add-in&#39;lerin öldüğünü, artık desteklenmediğini söyleyebiliyor. O yüzden mi MS, bunu VS içine koydu? Yani insan yorum yaparken mantıklı yorum yapmalı.</p>
            <p> &nbsp;Şu bir gerçek ki, VSTO ile özdeşlemiş olan Interop kütüphanesi(ilerde detaylıca göreceğiz) çok süper bir kütüphane değil. VBA&#39;deki Excel Object Model&#39;in hemen hemen bir kopyası gibi. Kolaylık sağlayan unsurlar eklenmemiş. Bu doğru, ancak bu genel olarak VSTO&#39;yı yetersiz yapmaz, Interop&#39;u yetersiz yapar. Biz Interop kütüphanesine ek olarak 3rd Party kütüphaneleri de göreceğiz. Ama tüm faaliyetimiz bir VSTO Add-in yaratarak başlayacak. Interop&#39;tan kısmen yararlanacağız, Ribbon arayüzü geliştireceğiz ve 3rd Party kütüphaneleri de kodumuza entegre edeceğiz. Terimin açık adı gayet açık ve net: Office uygulaması geliştirmekl için Visual Studiop&#39;dan yararlandığım her şey bence VSTO&#39;dur. O yüzden VSTO ifadesini çekinmeden kullanacağım. Hiç de eski bir teknoloji kullanmış olmayacağız dostlar, endişeniz olmasın.</p>
        <h4> Kuruluma Devam</h4>
            <p> Visual Studio&#39;yu indirdiğinizde VS Installer diye bir kurucu devreye girer ve size hangi bileşenleri kuracağınızı sorar. Eğer mevcutta VS&#39;nuz varsa Başlat menüsünden VS Installer’ı çalışıtırın ve Modify deyin.</p>
        <p> 
            <img alt="" src="/images/vsto_installer1.jpg" /></p>
        <p> Bileşen seçim ekranı ise aşağıdaki gibidir ve VSTO seçimi görseldeki gibi yapılır.</p>
        <p> 
            <img alt="" src="/images/vsto_installer2.jpg" class="zoomla" /></p>
        <h3> 
            İlk VSTO Add&#39;inimiz</h3>
            <p>VSTO'yu başlattığınızda ilk gelen ekran <strong>Create New Project</strong> diyelim ve şu seçenekleri işaretleyelim. (Şu anda kod yazmayacağız, o yüzden dil seçimine takılmayın)</p>           
        <p><img alt="new project" src="/images/vsto_newcs1.jpg" /></p>
        <p> Next diyip sonraki ekrana geçelim ve aşağıdaki gibi dolduralım. (Siz tabiki kendinize uygun bir klasör seçmelisiniz) </p>
        <p><img alt="new project" src="/images/vsto_newcs2.jpg" /></p>
            <p>Projeniz açıldığında bizi 3 bölmeli bir pencere karşılar. Sol bölmede Toolbox, sağ bölmede Solution Explorer ve ortada Kod yazma alanımız bulunur. Aslında hem sağ hem sol kısımda sekmeler halinde farklı pencereler yer alabilir. Mesela bir veritabanı bağlantınız da olacaksa Toolbox ile Server Explorer sol tarafı paylaşır, sekmelerle birinden diğerine geçeriz. Bu kısımları keşfetmeyi size bırakıyorum.</p>
            <p>Visual Studio olağanüstü güzel bir program arkadaşlar. Bunu iyice keşfetmenizi tavsiye ederim. Bunun için internette bol miktarda kaynak var. Bir ikisini buraya bırakabilirim:</p>
            <ul>
                <li><a href="https://tutorials.visualstudio.com/vs-get-started/intro">https://tutorials.visualstudio.com/vs-get-started/intro</a></li>
                <li><a href="https://www.edureka.co/blog/visual-studio-tutorial/">https://www.edureka.co/blog/visual-studio-tutorial/</a></li>
            </ul>
            <p>
                Her zamanki gibi önerim. Bu linklere hem şimdi bakın, hem paralelde kod yazarken bakın, hem de ilerde deneyim kazandığınızda yine bakın. Her defasında farklı şeyler alacaksınız.</p>
        </div>
</asp:Content>
