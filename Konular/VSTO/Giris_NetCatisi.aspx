<%@ Page Title='Giris NetCatisi' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Giriş'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>

<h1>.Net Framework(Çatısı)</h1>
<h3>.Net nedir?</h3>
<div><p> .NET("dat net" diye okunur) , Microsoft'un yazılım geliştirme(ve aynı zamanda bu platformda geliştirilen yazılımların çalıştırma) ortamıdır.</p>
    <p> Bu platformda hem masaüstü, hem web hem de mobil uygulamaları geliştirilebilmektedir. </p>
    <p> Platformdan, Microsoft ailesinin dilleri olan C#, VB.Net ve F# dillerine ek olarak, C++, ve Python gibi dillerde de program yazılabilmektedir.</p>
    <h3> Bileşenler</h3>
    <p> Bu platformun 3 bileşeni vardır. Bunlar;</p>
    <ul>
        <li>Programlama dili</li>
        <li>CLR(Common Language Runtime): </li>
        <li>Kütüphaneler</li>
    </ul>
    <p>
        Özetleyecek olursak, Programlama dili çeşitli kütüphanelerdeki sınıfları kullanarak yazdığımız program CLR&#39;de önce derlenir, sonra just-in-time yorumlayıcı adı verilen programa gönderilir ve programlar çalıştırılır. Yani CLR, programların çalıştığı ortamdır. Meşhur JVM de CLR&#39;nin Java karşılığıdır. Yani CLR ve JVM iki farklı şirketin aynı görevi gören araçlarıdır.</p>
    <p>
        Detayları merak edenler google araştırması yapabilir, ancak bizim şu aşamada bu detayları bilmemize gerek yok.</p>
    <p>Bu arada daha önce bahsettiğimiz gibi, bizim oluşturacağımız add-in projeleri bir dll dosya üretirler. Dosya ve klasör yapılarına bilahare bakacağız.</p>
</div>
<h3>.NET versiyonları</h3>
<div><p> Her uygulamada olduğu gibi .Net'in de zaman içince yeni sürümleri çıkmaktadır. </p>
    <p> Aslında, farklı versiyonları bulunan tek şey .Net framework değil. Kullandığınız dilin(vb.net veya c#), Visual Studio&#39;nun, CLR&#39;nin de versiyonları bulunmaktadır. Özellikle yeni versiyonda hangi yeniliklerin geldiğini bilmek bu anlamda önem arz etmektedir. VBA&#39;de şanslıydık, zira dilin kendisinde 2000lerin başından beri hiçbir geliştirme olmamakta, sadece Excel&#39;in yeni nesnelerine ait object modele eklentiler olmaktadır. Ancak burada versiyon değişikliklerini iyi takip etmek durumundayız. Çünkü çeşitli forumlarda araştırma yaparken sunulan çözümün hangi versiyona uygun olduğunu da bilmeniz gerekiyor. Örneğin c# 7&#39;ye ait bir çözüm sizde c# 6 varsa çalışmaz.</p>
    <p> <a href="https://www.guru99.com/c-sharp-dot-net-version-history.html">Bu siteden</a> özet, <a href="https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/versions-and-dependencies">bu siteden</a> ise detaylı tarihçeyi bulabilirsiniz. </p>
    <p> Siz arzu ederseniz internette bol miktarda bulunan .Net tanıtım dokümanlarını da inceleyebilirsiniz. Biz şimdi kodlarımızı yazacağımız ortam olan Visual Studio&#39;yu incelemeye geçelim. </p>
   
</div>

</asp:Content>
