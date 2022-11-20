<%@ Page Title='ProjeyiYayinlama WindowsInstallerolusturma' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'>
        <table>
            <tr>
                <td>
                    <asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td>
                <td>
                    <asp:Label ID='Label2' runat='server' Text='Projeyi Yayınlama'></asp:Label></td>
                <td>
                    <asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td>
            </tr>
        </table>
    </div>
    <h1>Windows Installer Oluşturma</h1>
    <p>ClickOnce deployment sayfasında belirttiğimiz gibi daha profesyonel görünümlü, ara stepleri olan ve tek seferde bir PC&#39;deki tüm kullanıcılara kurulum imkanı veren setup yöntemi <strong>MSI</strong> veya <strong>Windows</strong> <strong>Installer</strong> yöntemi olarak geçiyor. (Şirketinizin genelinde, adminler tarafından tüm PC&#39;lere dağıtım yapılması işi ise başka bir süreç olup, ben şahsen bunu hiç yapmadım ancak ihtiyacınız bu ise <a href="https://docs.microsoft.com/en-us/office/dev/add-ins/publish/centralized-deployment">bu sayfadan</a> bilgi edinebilirsiniz)</p>
    <p>Bu kurulum seçeneğinde tüm PC&#39;ye kurulum yapılacağı için kurulum yapan kişinin <strong>Admin</strong> yetkilerine sahip olması gerektiği de aşikardır. Şahsi PC&#39;lerde bu çok sorun olmayacaktır ancak kurumsal ortamlarda BT(IT) politikaları nedeniyle bu pek mümkün olmayabilir. Böyle bir durumda ClickOnce seçeneği değerlendirilebilir. </p>
    <p>Şimdi burda da geçmişe bakıldığında iki farklı yöntem görünüyor. Ben birine kısaca <strong>ISLE</strong> yöntemi, diğerine <strong>Installer</strong> yöntemi diyeceğim. <a href="https://stackoverflow.com/questions/31888465/visual-studio-2015-community-isle-setup-and-deployment-doesnt-appear/39717246"><span style="">Burada</span></a> ve <a href="https://stackoverflow.com/questions/12378125/create-msi-or-setup-project-with-visual-studio-2012">b<span style="">urada</span></a> görüleceği üzer ISLE yöntemi VS&#39;nun Community versiyonunda desteklenmiyormuş, sadece ücretli versiyonlarında destekleniyormuş, ancak bunun yerine sonra Installer yöntemi gelmiş, biz de bunu kullanacağız. Olur da ISLE yöntemini de denemek istersneiz <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-windows-installer?view=vs-2019">burdan</a> bilgi alabilirsiniz.</p>
    <p>Ana referans dokümanımız şu: <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2019">https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2019</a></p>
    <p>Referans dokümanında bahsedilen installer’ı indirmek için <a href="https://marketplace.visualstudio.com/items?itemName=visualstudioclient.MicrosoftVisualStudio2017InstallerProjects"><span style="">buraya</span></a> tıklayınız. Bu link VS 2017 ve 2019 için geçerli. Daha eski(veya yeni) versiyonlar için sayfadaki yönlendirmeleri takip ediniz.</p>
    <p>Bunu indirip çalıştırdıktan sonra(VS kapalı olsun bu arada) VS&#39;yu çalıştıralım. Ve pencereye yeni seçenekler geldiğini görelim.</p>
    <p>
        <img alt="msi" src="../../images/vsto_msisetup.jpg" style="width: 587px; height: 438px" /></p>
    <p>Şimdi ben VSTOvb projem içine bir setup programı oluşturacağım. O yüzden bu solution’ı açıyoruz.</p>
    <ul>
        <li>File menüsünden New Project diyip yeni eklenenlerden ilkini(<strong>Setup Project</strong>) seçiyoruz</li>
        <li>Solution kısmında &quot;<strong>Add do Solution</strong>&quot; diyip mevcut solution’a ekleyeceğimizi belirtelim, ben adını MySetup dedim.(Kötü bi isim ama, ana projeden kolay ayırdetmek adına şimdilik böyle diyebiliriz). <strong>Create </strong>dedikten sonra karşımıza aşağıdaki gibi bir pencere çıkacak.
            <img alt="msi setup1" src="../../images/vsto_msi_installer1.jpg" style="width: 612px; height: 160px" /></li>
        <li>Solutionda MySetup&#39;a sağ tıklayıp <strong>Add&gt;Project Output</strong> diyoruz, ve Primary Output seçiyoruz.
            <img alt="output" src="../../images/vsto_msi_installer2.jpg" style="width: 371px; height: 441px" /></li>
        <li>Yine MySetup&#39;a sağ tıkalyıp <strong>Add&gt; File</strong>, diyip ana proje klasörümüzdeki <strong>bin\Release</strong> yolunda vsto ve manifest uzantılı dosyaları ekliyoruz.
            <img alt="" src="../../images/vsto_msi_installer3.jpg" style="width: 753px; height: 470px" /></li>
        <li>Sonra, Solution Explorerda Detected Dependencies altındaki <strong>.Net Framework ve Utilites&#39;le bitenler hariç </strong>diğerlerine<strong> </strong>sağ tıklayıp Properties diyoruz ve <strong>Exclude=True </strong>diyoruz. Bunun sebebi özetle şu: Bunlar add-in&#39;imize pre-requisite olarak eklenmeli, yani kurulum sırasında öncelikli olarak kurulmalılar.&nbsp;
            <img alt="" src="../../images/vsto_msi_installer4.jpg" style="width: 298px; height: 218px" /><br />
            Hemen akabinde yine MySetup&#39;a sağ tıklayıp Properties&#39;e tıklayalım</li>
        <li>
            <img alt="" src="../../images/vsto_msi_installer5.jpg" style="width: 785px; height: 544px" /> Burda da Prerequisites&#39;e tıklayıp gerekli .Net Framework&#39;ü ve VSTO Runtime&#39;ı seçelim<img alt="" src="../../images/vsto_msi_installer6.jpg" style="width: 580px; height: 453px" /></li>
        <li>Bu aşamadan sonra uzunca bir Registry ayarlama süreci var, onu referans dokümanından bakabilirsiniz, ben aşağıya oluşması gereken resmi koydum. Bunun için MySetup&#39;a sağ tıklayıp <strong>View&gt;Registry </strong>diyorsunuz. CURRENTUSER ve LOCALMACHIINE altındaki <strong>Software&gt;Manufacturer </strong>Key&#39;lerini siliyorsunuz. Sonra en alttaki Key&#39;e resimdeki hiyerarşide olacak şekilde sırayla bu Key&#39;leri ekliyorsunuz. En alttaki Key&#39;e de yandaki 4 Value(ilk ikisi ve sonuncusu String, üçüncüsü DWORD tipinde) değerini ekliyoruz, bunların değerleri de yine yanlarında yazılı.
        <img alt="" src="../../images/vsto_msi_installer7.jpg" style="width: 965px; height: 325px" /></li>
        <li>Solution Explorer&#39;da MySetup seçiliyken aşağıdaki Properties penceresinde <strong>TargetPlatform </strong>belirliyoruz. 64 bit kullanacaklar içn 64, 32&#39;ler için 86 seçmek gerekiyor. İki konfigürasyonu olan kişiler de kullanacaksa iki ayrı setup dosyası oluşturacaksınız demektir. Hazır Properties açıkken TargetPlatform&#39;un biraz yukarısında Manufacturer alanında &quot;Default Company Name&quot; yazıyor, bunu da değiştirebilir, kendi adınızı veya şirket adını verebilirsiniz.</li>
        <li>Kurulum yapan kişide VSTO Runtime olup olmadığını kontrol eden bir Launch Conditon ekliyoruz.(Bu adım zaruri değil, başka bir projemde böyle denedim, sorun yaşamadım, olur da sorun yaşarsanız bu adımı işletin). Adımları yine referans dokümanda bulabilirsiniz, çıktısı aşağıdaki gibi olacak.<img alt="" src="../../images/vsto_msi_installer8.jpg" style="width: 848px; height: 186px" /></li>
        <li>Bu condition belirleme sürecinden bir tane daha var(Bu aşamayı hiçbir projemde yapmadım, sorun yaşamadım. Sorun yaşayan kullanıcılarınız olması durumunda işletebilirsiniz)</li>
        <li>Artık sona geldik, şimdi projemizi <strong>build </strong>edebiliriz. İşlem bitince File Explorer&#39;dan ilgili klasöre gidip <strong>bin\Release</strong> altındaki 2 dosyayı da alıp insanların erişebileceği bir klasöre kopyalayalım. Herkes buradan setup dosyasını çalıştırarak programı(add-ini) kurabilir.</li>
    </ul>
    <p>
        Gördüğünüz üzere, MSI kurulum yöntemi biraz daha teferruatlı. Tüm PC kullanıcılarına kurma gibi bir derdiniz yoksa, veya bir lisans key gibi birşey girdirmeyecekseniz ClickOnce tercih edin derim. Özellikle ticari amaçlı bir proje yapmıyor, kendiniz ve/veya birkaç kullanıcı için bir add-in hazırlıyorsanız MSI&#39;a hiç gerek yok.</p>
    <h2>Kurulum</h2>
    <p>
        Kurulum aşaması yine ClickOnce&#39;daki gibi basit:</p>
    <p>
        <img alt="" src="../../images/vsto_msi_installer9.jpg" style="width: 499px; height: 409px" /></p>
    <p>
        Next dedikten sonra çıkan ekranda nereye kurulum olacaksa orası seçilir. User&#39;a mı tüm PC&#39;ye mi kurulum yapacağını da soruyor, Everyone dersek tüm PC&#39;ye yani tüm kullanıcılara tek seferde kurulmuş olur.</p>
    <p>
        <img alt="" src="../../images/vsto_msi_installer10.jpg" style="width: 499px; height: 409px" /></p>
    <p>
        Ben Everyone dedim, sonra çıkıp PC&#39;dei Aile hesaplarından birine girdim, baktım ve gerçekten orda da kurulumuş olduğunu gördüm.</p>
    <p>
        Tabi Control Paneldeki Program &amp; Features&#39;an bakınca VSTOvb olarak değil MySetup olarak görünüyor. O yüzden setup projesinin isimlendirmesini daha anlaşılır yapmakta fayda var.</p>
    
</asp:Content>

