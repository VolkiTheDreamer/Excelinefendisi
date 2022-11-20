<%@ Page Title='ProjeyiYayinlama ClickOncedeployment' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'>
        <table>
            <tr>
                <td>
                    <asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td>
                <td>
                    <asp:Label ID='Label2' runat='server' Text='Projeyi Yayınlama'></asp:Label></td>
                <td>
                    <asp:Label ID='Label3' runat='server' Text='1'></asp:Label></td>
            </tr>
        </table>
    </div>
    <h1>ClickOnce Deployment</h1>
    
        <p>Uzun uğraşlar sonrasında add-in&#39;imiz bitti ve kullanıma hazır. Başka kişilerin VSTO add-in&#39;inizi kullanmaya başlaması için iki yol var.</p>
    <ul>
        <li>ClickOnce Deployment</li>
        <li>MSI Installer(Windows Installer)</li>
    </ul>
    <p>
        Biz burada ilkini göreceğiz, ikincisini de bir sonraki sayfada. Farkları şöyle:
    </p>
    <ul>
        <li>ClickOnce, bir bilgisayarda birden fazla kullanıcı çalışıyorsa sadece geçerli kullanıcının hesabına kurulum yapar, diğer kişiler de kullanacaksa onlar da tek tek kurmalıdır. MSI Installer ise herkesin hesabına kurmuş olur. </li>
        <li>ClickOnce kurulumu basit bir arayüze sahiptir, profesyonel bir program görüntüsü vermez. Üstelik ara stepler bulunmaz. MSI Installer ise bunları yapar.</li>
        <li>Projenizde güncelleme yaptığınızda ClickOnce ile kuranlar bu güncelleştirmeleri otomatik alırlar. Sizin tekrardan bi setup dosyası göndermenize gerek olmaz. MSI&#39;la kuranlar ise her yeni güncelleme sonunda yeni kurulum yapmak durumundalar.</li>
    </ul>
    <p>
        Kurulum yaptığınızda setup dosyaları Registry&#39;nizde kayıt oluşturur. Bu teknik bir detay olup ilgilenenler <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2019">buraya</a> bakabilir.</p>
    <p>
        Şimdi kurulum aşamalarına geçelim.</p>
    <h2>Kurulum aşamaları</h2>
    <p>
        ClickOnce deployment ile ilgili ana referans dokümanımız <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=vs-2019">şurada</a> olup ben size bu sayfada özetini vermeye çalışacağım.</p>
    <ul>
        <li>İlk olarak projemizin configuration tipi Debug&#39;tayken Build menüsünde <strong>Clean</strong> yapalım. Böylece registry kayıtlarını silelim ki çakışma problemleri yaşamayalım. Şuan mesela Registry&#39;de HKEY_CURRENT_USER&gt;Software&gt;Microsoft&gt;office altındaki görüntü şöyle:
            <img alt="registry1" src="../../images/VSTO_reg1.jpg"  />
            <p>Ancak 2 tanesini Clean ettikten sonra baktığımda şöyle:</p>
            <img alt="registry1" src="../../images/VSTO_reg2.jpg"  />
        </li>
       
        <li>Şimdi ikinci olarak configuration tipini&nbsp;<strong>Debug&#39;dan</strong> <strong>Release&#39;e</strong> alalım ve projemizi build edelim. Şimdiye kadar farkettiğiniz üzere projemizi parça parça oluşturup denemeler yaparken F5 tuşuna bastığında, yani Debug modda çalıştırdığımızda projemizin alt tarafta build edildiğini görmüşsünüzdür. Ve aslında bu tüm .Net projelerindeki sürecin aynısıdır. Çok kritik değil ancak gerek genel kültür gerek bir sorun yaşamanız durumunda Build süreci hakkında detay bilgi almak adına <a href="https://docs.microsoft.com/en-us/visualstudio/ide/compiling-and-building-in-visual-studio?view=vs-2019">buraya</a> bakabilirsiniz. Açılan linkte ve altındaki linklerde çok detaylı bilgiler edinebilirsiniz. Office programlarına özel olarak build işlemleri için de <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/building-office-solutions?view=vs-2019">buraya</a> bakabilirsiniz</li>
        <li>Solution Explorer&#39;da Properties&#39;e(VB.Net&#39;te MyProject) çift tıklayın ve proje pencerenizde <strong>Publish</strong> sekmesine gelin</li>
        <li>Açılan yerde <strong>Publish Folder Location</strong>&#39;da, kurulum dosyaları nereye kopyalanacaksa orayı seçin. Bi altındaki Installation Folder kısmı genelde boş bırakırız, böylece kullanıcılar da yine publish edilen folderdan kurmaya çalışır. Tabi siz buraya kurumunuzdaki herkesin erişeceği ortak alanı(Ör:\\M00002\\OrtakFolder\) seçmelisiniz. Veyahut, bir web alanı da girebilirsiniz.(<span style="text-decoration: underline"><strong>Türçke karakter olmamalı</strong></span>)</li>
    </ul>
    <p>
        <img alt="publish" src="../../images/vsto_publish1.jpg" style="width: 819px; height: 505px" /></p>
    <ul>
        <li>Prerequisites seçimleri defaulttaki gibi kalsın</li>
        <li>Updates butonunudan, kullanıcıların kurulum paketleri ne sıklıkta yeni versiyon arasın, onu seçiyoruz</li>
        <li>Options&#39;dan da projenizle ilgili genel bilgileri girebiliyorsunuz</li>
    </ul>
    <p>
        Son olarak <strong>Publish Now </strong>diyoruz. Publish Version alanında yazan numaralar projenizin versiyon numarasını takip etmek için kullanılır. Sonuncusu her publish sırasında otomatik olarak 1 artar. Diğerlerini siz manuel değiştirebilirsiniz. Bunların ne anlama geldiğinde google’da arayabilirsiniz.
    </p>
    <p>
        Bu arada olur da kodunuzun bir yerinde add-in&#39;inizin versiyonunu göstermek isterseniz aşağıdaki kod ile gösterebilirsiniz. Bunu yazabilmek için <strong>System.Deployment</strong> namespace&#39;ini reference olarak eklemelisiniz.</p>
    <pre class="brush:csharp">using System.Deployment.Application;
this.label1.Label = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);</pre>
    <p>
        Ama bu kod, <strong>debug </strong>modda değil, <strong>release </strong>modda kurulum yaptıktan sonra çalışır.</p>
    <h2>Kurulum</h2>
    <p>Kurulum dosyanız hazır ve kullanıcılarınıza bunu dağıttıktan sonra sadece <strong>setup.exe</strong>&#39;yi çalıştırmaları yeterli. Kendi PC&#39;nizde denerken herhangi bir sorunla karşılaşmazsınız, direkt aşağıdaki gibi bir pencere çıkacak, Install diyeceksiniz, o kadar. (Sonraki deneme yanılmalarınızda sorun yaşamanız ise muhtemel. Böyle bir durumda, projeyi <strong>Clean</strong> etmeyi ve/veya Control Panelden kurulmuş programınızı kaldırmayı deneyebilirsiniz.)</p>
    <p>
        <img alt="setup" src="../../images/vsto_setup1.jpg" style="width: 567px; height: 269px" /></p>
        <p>Ancak diğer kullanıcılarda sorun çıkabilir. Excelent add-in&#39;imde benim yaptığım gibi siz de ücretli sertifika almadıysanız <a href="https://www.excelinefendisi.com/Excelent/Download.aspx">https://www.excelinefendisi.com/Excelent/Download.aspx</a> sayfasında belirttiğim adımların işletilmesi gerekmektedir.</p>
    <p>Kurulum yapıldıktan sonra Registry&#39;de ayrıca <strong>VSTA</strong> folder&#39;ı altında bir de bunu application olarak görürüz.&nbsp;</p>
    <p>
        <img alt="registry" src="../../images/vsto_reg3.jpg" style="width: 1074px; height: 228px" /></p>
    <p>Sadece orada değil, aynı zamanda <strong>Control Panel>Programs</strong> altına da gelir. Zira artık bu bilgisayarımıza kurulmuş bi programdır. Bunu kaldırmak tıpkı normal programları kaldırmak gibidir. Yani Developer menüsünde COM add-ins altından check işaretin kaldırmak yetmez, o sadece &quot;Excel&#39;de yüklenmesini istemiyorum&quot; demektir ama PC&#39;nizde hala yüklüdür. Komple kaldırmak için Control Panelden işlem yaparız.</p>
    <p>
        <img alt="programfiles" src="../../images/vsto_denetimmasa.jpg" style="width: 431px; height: 132px" /></p>
    <p>Şimdi bu arada fark ettiyseniz, VSTOWorkbook çalışmamız da burda bir program gibi görünüyor. Evet, o da publish edildiği için bir program olarak görünecektir. Bunlarla ilgili işlemler de aşağı yukarı aynı olup ilave yapılması gerekenler için <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=vs-2019">ClickOnce linkinden</a> bilgi alabilirsiniz.</p>
    <h3>Performans</h3>
    <p>VSTO Add-in&#39;lerin performanısını iyileştirme konusundaki <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/improving-the-performance-of-a-vsto-add-in?view=vs-2019">şurada</a> ClickOnce yerine Windows Installer kullanılmasının bazı aşamların pypass edilmesi nedeniyle daha hızlı olduğu söylenmekte. Burda karar size kalmış, diğer performans artırıcı yöntemleri denediğiniz halde hala performans sorunu yaşıyorsanız sizi sonraki sayfaya alalım.</p>
</asp:Content>

