<%@ Page Title='Office Uygulamlarıyla Çalışmak' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' runat='Server'>
    <div id='gizliforkonu'></div>

    <h1>Office Uygulamalarıyla Çalışmak</h1>
    <h2 class='baslik'>Giriş</h2>
    <div class='konu'>
        <h3>Genel Bilgiler</h3>
        <p>VSTO projelerinde diğer ofis uygulamalarıyla çalışmanın VBA&#39;den çok büyük bi farkı yok. Burada da yine ilgili kütüphaneyi referans olarak ekleme işi var, sonra yine ilgili programın obje modeline erişme işi var. Sadece kullandığımız objelerin işi bittiğinde yok edilmesiyle ilgili olarak bir iki ufak detay var, onları göreceğiz.</p>


        <p>Yine VBA&#39;de dikkat ettiğimiz gibi, seçtiğiniz Office versiyonuna ait referansı içeren kodunuz, daha düşük bir Office versiyonu olan&nbsp;bir PC&#39;de çalışacaksa, sizdeki versiyonda yeni gelmiş özellikleri kullanmadığınızdan emin olmanız gerekiyor. Ör: Slicerlar 2010&#39;da geldi, Office 2007&#39;si olan bir kişi bunu kullanırken hata alır. O yüzden böyle bir ihtimal varsa kodunuzda kullanıcının Office versiyon kontrolünü yaptırabilirsiniz.</p>
        <p>Bu arada tabiki VSTO&#39;da da yine Late Binding tekniğini kullanabilir ve versiyon probleminden kaçınabilirsiniz, ancak VBA&#39;den hatırlayacağımız üzere Intellisense&#39;den yararlanamayacağınız gibi performans kaybı da olacaktır. Late Binding, özellikle c# için biraz farklı olabilmekte. Bunun için <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/late-binding-in-office-solutions?view=vs-2019">şu</a>, <a href="https://stackoverflow.com/questions/23873825/how-to-achieve-late-binding-in-c-sharp-for-microsoft-office-interop-word">şu</a>, <a href="https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/binding-type-available-to-automation-clients">şu</a> ve <a href="https://stackoverflow.com/questions/34195650/c-sharp-create-excel-sheet-late-bound">şu</a> linklere bakabilirsiniz</p>
        <p>
            Bu arada çok ihtiyacınız olacağını sanmam ama olur da bir nedenle sizdeki office versiyonundan daha eski bir Office versiyonunda çalışacak bir add-in hazırlamanız gerekiyorsa ilgili versiyona ait PIA denen referenceleri indirmeniz gerekir. Olası bir neden şu olabilir: Intellisenseten faydalanmak istiyorsunuz, yani Late Binding yapmayacaksınız, ilgili kütüphaneyi referans olarak vereceksiniz. Ancak sizdeki Office versiyonu, add-ini kullanacak diğer kişilerden düşükse onlarda kodunuz hata alır. Böyle bir durumda PIA yöntemi işinizi görür. Bunları da <a href="https://docs.microsoft.com/tr-tr/visualstudio/vsto/how-to-install-office-primary-interop-assemblies?redirectedfrom=MSDN&amp;view=vs-2019">buradan</a> indirebilirsiniz.(Ama ben olsam, early binding ile hazırlar, intellisensin nimetlerinden yararlanır, sonra tüm kodumu Late Binding&#39;e çeviririm, tabi performans sorunu olmayacaksa)</p>
        <h3>İşi biten nesnelerin yok edilmesi</h3>
        <p>.Net framework&#39;te diğer office uygulamalarına ve içlerindeki nesnelere <strong>COM nesneleri </strong>denir. Bir COM nesnesi yaratılıp kullanıldıktan sonra işi bitince bellekten atılması için özel işlemlere tabi tutulur. VBA&#39;deki gibi kısaca null atamak yeterli değildir.</p>
        <p>Outlook, Word gibi Office uygulamalarını add-inimizde kullandıktan sonra onlara ait nesneleri bellekten silmemiz gerekiyor. Bunun için onlara sadece null atamak yeterli olmuyor, ayrıca onlara ait referansları serbest bırakmamız(<strong>releasing</strong>) gerekiyor. Bahsettiğim şey aşağıdaki linklerden bazısında bahsedilen <strong>RCW</strong>(Runtime Callable Wrapper) nesnesi ile ilgili. Bu RCW nesnesi, bizim COM nesnesinin etrafını saran başka bi nesnedir ve COM nesnesine her başvuru yaptığınızda bu RCW nesnesine ait olan referans sayısı 1 artar. Ve COM nesnesi ile işimiz bittiğinde RCW&#39;ye yapılan tüm referansların sıfırlanması gerekir. Aksi halde bu RCW nesnesi serbest kalmadığı için ilgili COM nesnesine tekrar ulaşmaya çalıştığımızda hata alırız. Aşağıdaki örneklerden göreceğimiz üzere, bu nesne serbest kalmadığında ilgili office uygulaması <strong>Task Manager</strong>&#39;da hala yaşıyor görünecektir.</p>
        <p>Peki bu her zaman dikkat edilmesi gereken bir durum mudur? Aslında değil. Özellikle Excel&#39;den Outlook&#39;a ulaşırken çok sorun olacağını düşünmüyorum. Zira Outlook&#39;umuz genelde açıktır, yani bellekte ayrı bir Outlook nesnesi yaratmak yerine açık olana başvururuz. Zaten maillerimizin gitmesi için Outlook açık da olmalı, aksi halde maillerimiz outbox&#39;ta kalacaktır, biz de istediğimiz işlemi yapamamış olacağızdır. Ama yine de &quot;Ben Outlook kapalıyken de bunu kullanacağım, mailler de outboxta kalırsa kalsın, mailin acil gitmesi gerekmiyor, veya ben sadece mail gönderimi için Outlook&#39;a erişmiyorum, calendar işlemleri gibi işler de yapıyorum v.s&quot; diyorsanız da buradaki yöntemleri kullanabilirsiniz. Evet burda özellikle bahsedilen Office uyguılamasının kapalı olduğu durumdaki sıkıntıdan bahsediyoruz. Bu uygulamaların açık olması ise onları release etmek gibi bir derdimizin olmaması anlamına gelir. </p>
        <p>Fakat, diğer Office uygulamalarında bu endişe devam eder. Yani Word ile çalışırken bu işlemleri mutlaka yapmalısınız. Veya bir Outlook add-in&#39;i yapıp oradan Excel&#39;e ulaştığınızda Excel nesneleri için bu adımları işletmeniz gerekir.</p>
        <p>Bir diğer sıkıntı olmayacak durum, 3rd parti uygulamaları kullanarak işlemleri yapabiliyor olduğumuz zamanlardır. Mesela yine Outlook&#39;ta bir add-in yaptık diyelim; Excel&#39;e bilgi yazacağız, bunun için Interop&#39;a gerek yok. <a href="ThirdPartyKutuphanler_ClosedXML.aspx">ClosedXML</a> ile release derdi olmadan işlerimizi halledebiliriz.</p>
        <p>Ama ille de Interop kullandığınızda bu aşamaları geçmeniz lazım. Göreceksiniz ki çok farklı yöntemler/öneriler var. Biraz aşağıdaki linklerde verilen çözüm önerilerinin hepsini inceleyip denedim. Buna göre farklı caselerde farklı davranışlar söz konusudur, hepsini de göreceğiz.</p>
        <h4>Yöntemler&nbsp;</h4>
        <h5>Two Dot prensibi ve kodla ulaşılamayan ara nesneler</h5>
        <p>Yöntemlere geçmeden önce bi prensip hakkında bilgi edinelim. Bu prensip şunu der: Bir nesneyi kullanırken onu dolaylı olarak kullanmayın, ona ait değişkeni de mutlaka yaratıp öyle kullanın. Mesela bir A nesnesinin A1Prop propertysini kullandığımızda A1 nesnesini elde ediyoruz, A1 nesnesinin de A1aProp propertysini kullanarak da A1a nesnesi yaratabiliyoruz diyelim. Bu prensibe göre yapma<span style="text-decoration: underline"><strong>ma</strong></span>mız gereken kötü şeyler ve yapmamız gereken iyi şeyler şöyle:</p>
        <pre class="brush:csharp"> //kötü yöntem
var A= new A();
var A1a=A.A1Prop.A1aProp; //*Açıklama için aşağı bakın

//Doğru yöntem
var A= new A();
var A1= A.A1Prop;
var A1a = A1.A1aProp;
</pre>
        <p>
            *<strong>Açıklama</strong>: işte tam burada iki nokta(two dot) kullanılıyor. Bu noktada hem A1Prop nesnesi hem de A1aProp nesnesi olmak üzere iki COM nesnesi yaratılıyor. Bu iki nesne için de ayrı ayrı 3er tane nesne yaratılmış oluyor. Birisi bildiğimiz .Net nesnesi(A1Prop ve A1aProp), birisi bunların referans aldığı RCW nesnesi(buna biz kod sırasında dokunamayız), bir diğeri de bunların referans aldığı COM nesneleri(bunlara da kodumuzda dokunamayız)
        </p>
        <p>
            Şimdi nesneleri yarattık, kullandık ve artık işimiz bitti diyelim. İşi biten nesnelerden kurtulmak için 3 temel yöntem var.
        </p>
        <ul>
            <li><strong>Release</strong> yöntemleri. Bu yöntemlerin başarılı olabilmesi için <strong>two dot prensibine uygun hareket edilmesi</strong> gerekmektedir</li>
            <li><strong>Garbage Collector(GC)&#39;</strong>ı çağırmak(release modda işe yarar, debug modda manuel kill lazım). Two dot prensibine gerek duyulmuyor.</li>
            <li>İlgili office uygulamasını öldürmek(<strong>Task manager&#39;</strong>daki Outlook.exe, word.exe prosesini sonlandırmak): Biz buna hiç değinmeyeceğiz bile, çünkü bu yöntem ne şık ne de verimli.</li>
        </ul>
        <h5>Release Yöntemi</h5>
        <p>
            Hem debug hem release modda işe yarar. (İki release ifadesi birbirine karışmasın. Release yöntemi derken nesneyi serbest bırakma, bellekten atma anlamında kullanıyoruz; release mod derken ise VS&#39;nun release modu yani projenizi yaygınlaştırma modunu kastediyoruz).
        </p>
        <ul>
            <li>Bu yöntemde, ilk olarak nesnemize referansta bulunan tüm referanslar özel bir metod(<strong>System.Runtime.InteropServices</strong> içinde <strong>Marshal.ReleaseComObject)</strong> ile yok edilir. Aksi halde ilgili nesne hala arka planda yaşamaya devam eder, ta ki Garbage Collector alıp onu yok edene kadar, sonra da null(Vb.Net&#39;te Nothing) ataması yapılır. </li>
            <li>En alt nesneden başlayarak yukarıya&nbsp;doğru sırayla önce release etme ve nesne değişkenlerine null atama</li>
            <li>En son Applicatio'ndan çıkma(Quit) ve buna ait nesneyi release edip null yapma. </li>
        </ul>
        <p>
            Bu yöntemde ise 3 farklı alt yöntem bulunuyor.</p>
        <p>
            <strong>Tek release</strong>: ilgili nesne bi kere release edilir. Genelde bu yeterlidir, bu metod aslında o nesneye olan başvuruların sayısını bir azaltıyor. Çoğu yerde birden fazla referans olması durumundan bahsedilmiş ama ben bunu simüle edemedim, aynı nesneden ikinci defa yarattım, o nesnenin bir metodundan bir başka nesne yarattım, o nesneyi iki ayrı List&#39;e ekledim, her defasında tek release yeterli oldu.
        </p>
        <p>
            <strong>Loop release</strong>: Birden fazla referans varsa, bu kullanılır deniyor, ama dediğim gibi buna hiç gerek olmadı</p>
        <p>
            <strong>FinalRelease</strong>: Sonradan gelen bi metod, loop release yerine bu da kullanılır ama başka bir add-in de kullanıyorsanız ve ilgili diğer com nesnelerine bu add-inler de ulaşıyorsa sıkıntı olur deniyor.</p>
        <h5>
            GC Yöntemi</h5>
        <p>
            Bu, sadece release modda çalışır, debug modda etkisini göremezsiniz. Collect ve WaitForPendingFinalizers metodlarını çağırırız. Yanlız bunları iki kere yapıyoruz, akabinde ilgili ofis uygulamasından çıkıyoruz(Quit metodu)</p>
        <h5>Hangisi tercih edilmeli?</h5>
        <p>Aşağıdaki kaynaklara bakıldığında kimi diyor ki, finalrelease kullanmayın; kimi diyor hiç release etmeyin, GC kullanmak yeterli, sadece bunu debug modda denemeyin; kimsi de GC&#39;yi hiç önermiyor. Tam deli işi. O kadar delice ki, birinin en güvenilir dediğine diğer kaynak çok tehlikeli diyebiliyor ve bunu her iki taraf da diyor. O yüzden üşenmedim hepsini denedim. Şimdi benim testlerime bakalım, sonra sonuçları görelim.</p>
        <h4>Benim testlerim</h4>
        <p><strong>oMail&#39;e(mail nesnesi) tek referans varken</strong>: Her iki yöntem de işi yaradı. Ancak sadece oApp&#39;i release etmek yeterli değil, her halükarda oMail de release edilmeli, yoksa bunun üzerinden hala oApp'e erişim kalıyor.</p>
        <p><strong>oMail için iki referans:</strong></p>
        <ul>
            <li>Sadece release ile: Gerçekten de yaşamaya devam etti, hala bi referans var.
            <ul>
                <li>Finalrelaese omail:tek omailde yapınca yine işe yaramadı, zira recipient hem outmaile hem oappe referansta bulunuyor, oAppte de finalrelease yapmak lazım. </li>
                <li>Finalrelease oAppte de yapınca:yine olmadı, çünkü two dot kuralını ihlal ettim, </li>
                <li>Looplu ReleaseCom: yine olmadı, çünkü two dot ihlali var </li>
                <li>Release modda GC:OK</li>
            </ul>
            </li>
            <li>
                Two dot kuralını uygulayınca:
                <ul>
                    <li>Finalrelease omail+oapp: olmadı</li>
                    <li>Finelreasle receivers, receiver, oMail ve oApp : oldu</li>
                    <li>Looplu release receivers, receiver, oMail ve oApp : oldu</li>
                    <li>GC ok, ama two dot ile satır sayısını artırmaya gerek yok</li>
                </ul>
            </li>
        </ul>
        <p>Test kodlarının hepsini buraya koymaya gerek görmedim, sadece bir başarılı olan bir de başarısız olanı koyuyorum. Diğerlerine örnek dosyadan bakabilirsiniz.</p>
        <p>UYARI: Testin anlamlı olabilmesi için Outlook&#39;umuzun kapalı olmasını tekrar hatırlatmak isterim</p>
        <pre class="brush:csharp">
private void btnTekRelease_Click(object sender, RibbonControlEventArgs e)
{
    ///Tek release yapıyoruz. Two dot prensibine uyuyoruz. Birden fazla referans bulunmadığı için sıkıntı yaşamıyoruz
    try
    {
        outlook.Application oApp;
        outlook.MailItem oMail;
        Outlooktip ot;
        if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
        {
            ot = Outlooktip.Varolan;
            oApp = (outlook.Application)Marshal.GetActiveObject("Outlook.Application");
        }
        else
        {
            ot = Outlooktip.Yeni;
            oApp = new outlook.Application();
        }
        oMail = oApp.CreateItem(outlook.OlItemType.olMailItem);  
        //two dot prensibini ihlal etmeden iki ayrı değişken oluşturuyoruz
        outlook.Recipients receivers = oMail.Recipients; 
        outlook.Recipient receiver = receivers.Add("volkan.yurtseven@hotmail.com");
        receiver.Type = (int)outlook.OlMailRecipientType.olCC;

        oMail.To = "mvolkanyurtseven@gmail.com";
        oMail.Subject = "VSTO-Tek Release";
        oMail.Body = "Bu bir deneme mailidir";
        oMail.Send();
        System.Windows.Forms.MessageBox.Show("İşlem tamam");

        //Releasing objects                                 
        Marshal.ReleaseComObject(receiver);
        receiver = null;
        Marshal.ReleaseComObject(receivers);
        receivers = null;
        Marshal.ReleaseComObject(oMail);
        oMail = null;

        if (ot == Outlooktip.Varolan)
        {
            //Burda program açık kalmalı, o yüzden Quit metodu yok. Marshal.Release de yapmıyoruz,zira program hala açık
            oApp = null;
        }
        else
        {
            oApp.Quit();
            Marshal.ReleaseComObject(oApp);
            oApp = null;
        }
    }
    catch (Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(ex.Message);
    }
}            
</pre>
        <p>
            Two dot prensibini ihlal ettiğimiz kod ise şöyle</p>
        <pre class="brush:csharp">
private void btnTwoDotIhlal_Click(object sender, RibbonControlEventArgs e)
{
    ///two dot prensibini ihlal edeceğiz, LoopRelease ile yapacağız ama FinalRelease ile yapsaydık da değişmeyecekti, çünkü two dot ihlali iki yöntemi de affetmiyor. o yüzden hala arkada yaşamaya davem edecek
    try
    {
        outlook.Application oApp;
        outlook.MailItem oMail;
        Outlooktip ot;
        if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
        {
            ot = Outlooktip.Varolan;
            oApp = (outlook.Application)Marshal.GetActiveObject("Outlook.Application");
        }
        else
        {
            ot = Outlooktip.Yeni;
            oApp = new outlook.Application();
        }
        oMail = oApp.CreateItem(outlook.OlItemType.olMailItem);
        outlook.Recipient receiver = oMail.Recipients.Add("volkan.yurtseven@hotmail.com"); //two dot
        receiver.Type = (int)outlook.OlMailRecipientType.olCC;

        var act = oMail.Actions;//omail için başka bi referans                

        oMail.To = "mvolkanyurtseven@gmail.com";
        oMail.Subject = "VSTO-Release ama two dot ihlali";
        oMail.Body = "Bu bir deneme mailidir";
        oMail.Send();
        System.Windows.Forms.MessageBox.Show("İşlem tamam");


        //Releasing objects                     
        ReleaseWithOrder(act, "receiver");
        ReleaseWithOrder(receiver, "receiver");
        ReleaseWithOrder(oMail, "oMail");                

        if (ot == Outlooktip.Varolan)
        {
            oApp = null;
        }
        else
        {
            oApp.Quit();                    
            ReleaseWithOrder(oApp, "oApp");
            oApp = null;
        }
    }
    catch (Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(ex.Message);
    }
}</pre>
        <p>
            Bu kodda geçen RelesaseWithOrder isimli yardımcı fonksiyonun koda da şöyle:</p>
        <pre class="brush:csharp">
private void ReleaseWithOrder(object o,string ad)
{
    try
    {
        //while (Marshal.ReleaseComObject(o)>0) { }; //normalde bu satır yeterli, ben test amaçlı aşağıdaki bloğu oluşturdum
        int cnt = Marshal.ReleaseComObject(o);
        while (cnt > 0)
        {
            //buraya hiç girmedi, o yüzden çoklu referans durumunu hiç simüle edememiş oldum. Bu durumda ilk yöntemle bu yöntem arasında fark bulunmuyor
            Trace.WriteLine(ad + " nesnesi için " + cnt.ToString() + ". release");
            cnt =Marshal.ReleaseComObject(o);
        }
    }
    catch
    {
        //
    }
    finally
    {
        o = null;
    }
}
           &nbsp;</pre>
        <p>NOT: Hepsinin başlarında &quot;sıkıntı yaşamıyoruz&quot; derken kastım kodun ilk çalışmasından değil, sonraki çalıştırmaları kastediyorum, zira zaten ilk çalıştırmada problem olmaz, önemli olan sonraki çalıştırmalarda arkada yaşamaya devam ediyor mu, onu tespit etmek.</p>
        <p>
            Tabi bu arada sorunsuz çalışan yöntemlerde task managerdan gözlemekmek lazım, hemen gitmiyor, 2-3 sn gitmesini bekleyin, ondan sonra deneyin. 
        </p>
        <h4>Sonuç</h4>
        <p>Örnek doysyayı da incelediyseniz, göreceğiniz gibi FinalRelease, looplu release yerine rahatlıklık kullanılabilir. Sadece başka add-inler de devredeyse o zaman sıkıntı çıkarabilir. O zaman looplu release en güvenilir olanı. Bunda da zincirleme çok nesne varsa her biri için çok satır yazmak gerekiyor, pek şık değil. 
            <strong>Bence en pratik yöntem, two dot uygulamadan(satır sayısı artıramaya gerek yok) release modda GC yapmak. </strong>
            Release metodları hem çok satır alıyor, hem büyük projelerde gözden kaçan bir ara nesne varsa(two dot nedeniyle) onu bulmak çok vakit alabilir.
            <strong>Özetle GC candır.</strong>
            Two dots prensibini de unutun gitsin, ama araştırdığınız kod örneklerinde görebilirsiniz diye buraya dahil ettim.</p>
        <h4 class="baslik">Çeşitli linkler</h4>
        <div class="konu">
        <p>
            Konuyla ilgili birçok kaynak var. Benim bu yukarıda anlattıklarım genel olarak yeterli olmakla birlikte daha teknik detayları öğrenmek isteyenler bu linklere bakabilir. Ama yukarıda belirttiğim gibi bunlar kafanızı çok da karıştırabilir,&nbsp;zira&nbsp;birinin dediği ile diğerininki tutmayabiliyor.
        </p>
        
            <ul>
                <li><a href="https://www.add-in-express.com/creating-addins-blog/2013/11/05/release-excel-com-objects/">https://www.add-in-express.com/creating-addins-blog/2013/11/05/release-excel-com-objects/</a></li>
                <li><a href="https://www.add-in-express.com/creating-addins-blog/2011/11/04/why-doesnt-excel-quit/">https://www.add-in-express.com/creating-addins-blog/2011/11/04/why-doesnt-excel-quit/</a></li>
                <li><a href="https://www.add-in-express.com/creating-addins-blog/2008/10/30/releasing-office-objects-net/">https://www.add-in-express.com/creating-addins-blog/2008/10/30/releasing-office-objects-net/</a> </li>
                <li><a href="https://stackoverflow.com/questions/158706/how-do-i-properly-clean-up-excel-interop-objects">https://stackoverflow.com/questions/158706/how-do-i-properly-clean-up-excel-interop-objects</a>
                </li>
                <li><a href="https://stackoverflow.com/questions/17130382/understanding-garbage-collection-in-net/17131389#17131389">https://stackoverflow.com/questions/17130382/understanding-garbage-collection-in-net/17131389#17131389</a> </li>
                <li><a href="https://stackoverflow.com/questions/25134024/clean-up-excel-interop-objects-with-idisposable">https://stackoverflow.com/questions/25134024/clean-up-excel-interop-objects-with-idisposable</a></li>
                <li><a href="https://stackoverflow.com/questions/29067714/vsto-manipulating-com-objects-one-dot-good-two-dots-bad">https://stackoverflow.com/questions/29067714/vsto-manipulating-com-objects-one-dot-good-two-dots-bad</a> </li>
                <li><a href="http://www.siddharthrout.com/index.php/2018/01/12/vb-net-two-dot-rule/">http://www.siddharthrout.com/index.php/2018/01/12/vb-net-two-dot-rule/</a> </li>
                <li><a href="https://www.breezetree.com/blog/common-mistakes-programming-excel-with-c-sharp">https://www.breezetree.com/blog/common-mistakes-programming-excel-with-c-sharp</a> </li>
                <li><a href="https://www.codeproject.com/Tips/162691/Proper-Way-of-Releasing-COM-Objects-in-NET">https://www.codeproject.com/Tips/162691/Proper-Way-of-Releasing-COM-Objects-in-NET</a>  </li>
                <li><a href="https://stackoverflow.com/questions/3937181/when-to-use-releasecomobject-vs-finalreleasecomobject">https://stackoverflow.com/questions/3937181/when-to-use-releasecomobject-vs-finalreleasecomobject/a> </li>
                <li><a href="https://stackoverflow.com/questions/1827059/why-use-finalreleasecomobject-instead-of-releasecomobject">https://stackoverflow.com/questions/1827059/why-use-finalreleasecomobject-instead-of-releasecomobject</a> </li>
                <li><a href="https://support.microsoft.com/tr-tr/help/317109/office-application-does-not-exit-after-automation-from-visual-studio-n">https://support.microsoft.com/tr-tr/help/317109/office-application-does-not-exit-after-automation-from-visual-studio-n</a> </li>
            </ul>
        </div>
    </div>
    <h2 class="baslik">Outlook ile Çalışma</h2>
    <div class='konu'>
        <p>Öncelikle şunu belirteyim. .Net&#39;in mail göndermeyle alakalı <strong>System.Net.Mail</strong> isimli bir namespace&#39;i var. Mail işlemleri genelde bununla yapılır. Ancak olur da bir şekilde(Ör:BT politikalarının izin vermemesi) bu namespace&#39;i kullanmak yerine Outlook kullanmanız gerekirse, veya Outlook folderlarına, calendar&#39;a v.s erişmeniz gerekiyorsa o zaman mecburen Outlook kütüphanesiyle çalışmak gerekecek.</p>
        <h3>Referans ekleme</h3>
        <p>Early binding yapacağız, o yüzden referans ekleriz.</p>
        <p>
            <img alt="" src="../../images/vsto_digerofficeoutlook1.jpg" style="width: 760px; height: 435px" />
        </p>
        <p>
            Bunu eklediğimizde <strong>stdole</strong>&#39;den bi tane daha ekliyor, onu silelim.
        </p>
        <p>Kodları yukarıda gördüğümüz için ayrıca burada gerek görmüyorum. Sadece yukarıda detaylı bahsetmediğimiz bir iki yer var, onları açıklayalım.</p>
        <p>Öncelikle ulaştığımız Office uygulamasının o anda açık olup olmadığını kontrol ediyoruz. Bunun için bi enumaration yarattım, istersek string bir &quot;Yeni&quot;/&quot;Varolan&quot; parametresi de gönderebilirdik, ama enum yaratmak daha havalı.</p>
        <pre class="brush:csharp">
private enum Outlooktip
{
    Varolan,
    Yeni
}
</pre>
        <p>
            Bunu da aşağıdaki gibi kullanıyoruz. O anda açık bir Outlook uygulaması olup olmadığını <strong>Process.GetProcessesByName("OUTLOOK").Count()&gt;0</strong> kodu ile sorguluyoruz. Evetse yine <strong>Marshall</strong> sınıfının <strong>GetActiveObject</strong> metodu ile bunu oApp nesnemize atıyoruz. Eğer Outlook o an açık değilse bellekte yeni bir Outlook nesnesi yaratıyoruz.
        </p>
        <pre class="brush:csharp">
if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
{
    ot = Outlooktip.Varolan;
    oApp = (outlook.Application)Marshal.GetActiveObject("Outlook.Application");
}
else
{
    ot = Outlooktip.Yeni;
    oApp = new outlook.Application();
}
</pre>
        <p>
            Aradaki diğer işlemleri da yaptıktan sonra Application nesnesini de bellekten atmamız gerekiyor, burada da yine Outlook&#39;a erişim tipine göre ya sadece nesneyi null yapıyoruz(Varolan durumu) ya da uygulamadan da çıkış gerektiren diğer kodlar(Yeni durumu)
        </p>
        <pre class="brush:csharp">
//ara kodlar

if (ot == Outlooktip.Varolan)
{
    oApp = null;
}
else
{
    //GC veya Release yöntemlerine göre uygun kodlar
}
</pre>
        <p>Bunların dışında VBA&#39;de gördüğümüz yöntemler aynen kullanılabilir.</p>
    </div>
    <h2 class="baslik">Access</h2>
    <div class="konu">
        <p>VBA&#39;de veritabanı işlemlerini anlatırken yaptığımız gibi burada da Access Object modeline girmeyeceğiz. O yüzden bir COM nesnesi yaratmış olmayacağız. Sadece Access'ten data nasıl çekilir, ona bakacağız.</p>
        <h3>ADO.Net</h3>
        <p>VBA&#39;deyken DAO ve ADO kullanmıştık. Burada ise .Net dünyasıyla gelen ADO.Net kullanacağız. Ancak bir nedenle klasik ADO da kullanmak isteyebilirsiniz. Bunun için en geçerli sebep, Range nesnesine ait CopyFromRecordset metodunu ADO.Net ile kullanamıyoruz, zira ADO.Net&#39;te böyle <strong>Recordset</strong> diye bi nesne yok. Onun yerine bir datatable yaratıp bir loop ile buradaki kayıtları yazdırıyoruz(Başka yöntemler de var, sonra bahsedieceğim). 'Ben yeni bir şey öğrenmek istemiyorum, zaten ADO&#39;yu zor öğrendim' derseniz ve ADO kullanmak isterseniz References&#39;te COM altından ekleyebilirsiniz.</p>
        <p>
            <img alt="" src="../../images/vsto_ado.jpg" style="width: 569px; height: 170px" /></p>
        <p>Biz şimdi ADO.Net ile gidelim. Bunun için ayrı bir reference eklemeye gerek yok. <strong>System.Data </strong>altındaki namespaceleri kullanacağız. Access için <strong>OleDb</strong> kullanabiliriz. Using bloğuna bunu eklememiz yeterli. Ayrıca çekeceğimiz veriyi bir DataTable içinde tutacaksak <strong>System.Data</strong> namespaceini de projeye dahil ederiz</p>
        <p>Bu arada veritabanı işlemleri için şöyle <a href="https://www.buraksenyurt.com/post/DataAdapter-Kavramc4b1-ve-OleDbDataAdapter-Sc4b1nc4b1fc4b1na-Giris-bsenyurt-com-dan">güzel bir bir kaynak</a> var. Artık yayında değilse de küçük bir google araştırması ile güncel kaynaklara da ulaşabilirsiniz.</p>
        <h3>Başka bir Excel dosyasından veri okuma(Dosyayı açmadan) </h3>
        <h4>ADO.Net ile</h4>
        <p>VBA&#39;de de bunu yapmıştık. ADO yerine ADO.Net kullanarak da aynısını yapabiliriz. Yalnız burada 255 karakter sorunu vardır, buna dikkat edilmesi gerekir. Yani bir kolondaki ilk metin 255 karakterden az ise, bu kolon normal text gibi algılanır ve o kolondaki ondan sonraki tüm satırlar da text gibi düşünülür ve 255'ten uzun olan metinleri 255'te keser. İlk satır 255'ten büyükse bunların tipini memo(long string) olarak algılar, ama buna güvenemezsiniz. O yüzden bazı registry ayarları yapılması gerekiyor. Ben bu detayı buraya koymadım, isteyen araştırabilir. Size önerim şu: Eğer 3rd party paket yükleyerek projenizin boyutunu yükseltmek istemiyorsanız ve metinsel bilgilerin her zaman 255ten kısa olduğundan eminseniz bunu kullanın. Aksi halde bir alt başlıktaki çözümü kullanın.</p>

        <p>Yukarıda demiştim ki Ado.Net içinde CopyReceodrset metodu yok, onun yerine loop yapıyoruz. Bu loop, biraz uzunca olabiliyor, özellikle başlıkları falan da alacaksak. Bir de bu çok sık kullanılabilecek bir kod, o yüzden bunu static bic class içine koyup ihtiyaç oldukça çağırmak en iyisi. Ve hatta bu fonksiyon başka birçok projede kullanılabilecek bir fonksiyon, o yüzden böyle sık kullanılacak kodlar için kendinize ait bir utility kütüphanesi yazmak çok daha mantıklı olabilir. Mesela benim de VolkansUtility diye böyle bir paketim var, githubdan indirebilirsiniz. Hala geliştirme aşamasında olduğunu söylemeliyim. (Siz de böyle bir paket yapacaksanız, projenin Output Type&#39;ını Class Library olarak seçmeniz gerekiyor)</p>
        <p>
            <img alt="" src="../../images/vsto_utility.jpg" style="width: 634px; height: 165px" /></p>


        <p>Aşağıdaki kodda  <strong>ReadFromExcelIntoDTWithOledDB</strong> metodunun içini görmek için utility paketine göz atabilirsiniz.</p>
        <pre class="brush:csharp">
private void btnExcelOledb_Click(object sender, RibbonControlEventArgs e)
{
    //okuma
    string file = @"E:\OneDrive\Uygulama Geliştirme\web sitelerim\Yeni Efendi\Ornek_dosyalar\pivotdata.xlsx";
    DataTable dt = ExcelRW.ReadFromExcelIntoDTWithOledDB(file, "select Kalem, Sum(Rakam) from [Sheet1$] Group by Kalem");
    //yazma
    ExcelRW.WriteDataTableContentToActiveWBWithInterop(dt, ExcelRW.TargetLocation.ActiveCell);
}            &nbsp;</pre>
        
        
        <h4>3rd Party uygulamalar ile</h4>
        <p>Kendi paketleriniz yazmak yerine hazır paketleri da kullanabilirsiniz. Daha önce <a href="ThirdPartyKutuphanler_ClosedXML.aspx">ClosedXML</a>&#39;i görmüştük. Burada datayı Excel&#39;e yazma işi oldukça pratik. Adından da anlaşılacağı üzere bu paket ile kapalı dosyalarla işlem yapıyoruz. Interop ile yaparken açık dosyada işlem yapıyoruz. Daha önce diğer işlemleri gördüğümüz için ben şimdi sadece Excel'e yazma kısmından bahsedeceğim.</p>
        <p>Veri okuma için de bi paket var. ClosedXML&#39;i anlatırken <strong>ExcelDataReader</strong> paketinden en sonda bahsetmiştim. Şimdi bunla ilgili bi örnek yapalım. Her ne kadar bu paketler pratiklik sağlasa da bunları daha da pratik kullanmak için yine Utility paketime veri okumayla ilgili bir fonksiyon koydum. Bunu kullanalım. Detaya gerek Utility paketinden gerek geliştiricinin kendi sitesinden bakabilirsiniz.</p>
        <p>Burada kullanıcıdan bi seçim yapmasını bekliyorum; kullanıcı 1 derse bunu sayfa numarası gibi algılayacağız, 2 derse Tarihsel Data sayfasına gideceğiz. Burda c#&#39;ı yeni öğrenenler için iki güzel özellik var. İlk olarak, static sınıfımdaki metod dynamic tipte parametre alıyor, yani sayfa bilgisini index olarak da yani integer gönderebilirim, isim olarak da yani string de. Bir diğer&nbsp; özellik de varolan bir sınıfa kendimiz de metodlar yazabiliyoruz, bunlara extension metodları deniyor. Ben de mesela string sınıfı için ConvertIoInt sınıfını yazdım. string bir ifadeyi yazdıktan hemen sonra ConvertIoInt diyebiliyorum. Alternatifi, Convert.ToInt32 diye en başa yazmaktı, ki extension hali çok daha pratiktir.</p>
        
        <pre class="brush:csharp">
private void btnExcelReader_Click(object sender, RibbonControlEventArgs e)
{
    //okuma
    string cevap = Interaction.InputBox("Seçim yapın:1,2");
    string file = @"E:\OneDrive\Uygulama Geliştirme\web sitelerim\Yeni Efendi\Ornek_dosyalar\pivotdata.xlsx";
    DataTable dt;
    if (cevap == "1")
    {
        dt = ExcelRW.ReadFromExcelIntoDTWithExcelReader(file, cevap.ConvertIoInt()); //extension metod
    }
    else
    {
        dt = ExcelRW.ReadFromExcelIntoDTWithExcelReader(file, "Tarihsel Data", "Ürün='Ürün1'", new string[] { "Bölge Kodu", "Ay", "Aylık Gerç" });
    }
    //yazma
    ExcelRW.WriteDataTableContentToActiveWBWithInterop(dt, ExcelRW.TargetLocation.ActiveCell);
}
            &nbsp;</pre>
    </div>

    <h2 class="baslik">VSTO kodumuzu VBA&#39;de kullanma</h2>
    <div class="konu">
    <p>

        Diyelim ki .Net ile yazdığınız, Excel&#39;le işlem yapan bir sınıf yazdınız veya bir 3rd party kodun pratik özelliklerini çok beğeniyorsunuz. &quot;Keşke bu sınıfı VBA&#39;de de kullanabilsem&quot; diyorsunuz. İşte bu kısımda bu hayaliniz gerçek oluyor.</p>
        <p>

            Örnek olarak, yukarıda bahsettiğimiz OledDB ile bir Excel dosyadan, o kapalıyken veri okuma problemini ele alalım. Yukarıda bahsettiğimiz gibi biz bunu VBA ile sadece ADO kullanarak yapabiliyoruz, ancak orada da 255 karakter sorunu var. Biz bunun yerine .Net&#39;te ne yaptık, ExcelReader kütüphanesini kullandık. İşte bu kütüphaneyi VBA&#39;de kullanacağız. Zira bazen ana kodumuz VBA içinde olabiliyor ve ADO ile bi workbookdan veri okuması yapmamıaz gerekebiliyor. 255 karakter problemi ise problem yaratıyor. Hadi şimdi bu sorunu çözelim. </p>
        <p>

            Burda ana başvuru kaynağımız <a href="https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-calling-code-in-a-vsto-add-in-from-vba?view=vs-2019">şu sayfa</a> olacak.Tamamen bu işe özel bir add-in yaratacağınız gibi mevcut add-ininiz içine de entegre edebilirsiniz, ki biz burada öyle yapacağız.</p>
        <ul>
            <li>İlk olarak yeni bir class yaratalım, adına <strong>VBAClass </strong>diyelim.</li>
            <li>Class&#39;ımızın içindeki kodu şu şekilde değiştirelim.<br />
                <br />
            <pre class="brush:csharp">
using System.Data;
using System.Runtime.InteropServices;
using VolkansUtility;

[ComVisible(true)]
public interface IVBAClass
{
    void ExceldenOkuveAktiveCelleYaz(string dosya, string sayfa);
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public class VBAClass : IVBAClass
{    
    public void ExceldenOkuveAktiveCelleYaz(string dosya, string sayfa)
    {
        DataTable dt = ExcelRW.ReadFromExcelIntoDTWithExcelReader(dosya, sayfa);
        ExcelRW.WriteDataTableContentToActiveWBWithInterop(dt, ExcelRW.TargetLocation.ActiveCell);
    }
}
</pre>
</li>
            <li><p>Bu kod ile VBAClass isimli sınıfımızı COM bileşenlerine görünür hale geitriyoruz. Yani bir COM nesnesi kodumuza ulaşaiblir diyoruz, ki VBA de bir COM bileşenidir. Aynı zamanda sınıfımızın <strong>IDispatch </strong>interface'ini de implemente etmesi gerekmektedir, bunu sağlayan da <strong>ClassInterface</strong> attribute'ümüz oluyor. Interface konusu başlı başına bir konu olup burada detayına girmeyeceğiz. Burada bir interface kullanımı da <strong>IVBAClass</strong> isimli kendi Interface'imiz. Bununla özetle şunu demiş oluyoruz: &quot;Yaratacağımız sınıfta mutlaka falanca metod olsun&quot;. Yani, interface'ler belli metodların/propertylerin bir sınıfta mutlaka bulunmasını garantini altına alan yapılar olarak düşünülebilir, ve genelde sınıf adının önüne I harfinin getirilmesiyle tanımlanırlar.</p>
            </li>
            <li>
                <p>Son olarak, sınıfımızı diğer office uygulamalarına açabilmek için ThisAddin sınıfının <strong>RequestComAddInAutomationService </strong>metodu override edilmelidir. Overriding de yine kapsam dışı bir konu olup, özetle mevcut bir sınfın metodunun içeriğinin değiştirilmesidir.<br />
                <pre class="brush:csharp">
private VBAClass vbc;
protected override object RequestComAddInAutomationService()
{
    if (vbc == null)
        vbc = new VBAClass();

    return vbc;
}                </pre>
                </li>
            <li>Artık .Net tarafındaki işmiz bitmiştir. Projemizi <strong>build </strong>edebiliriz.</li>
            <li><p>Son olarak Excel&#39;e geçeriz ve herhangi bir modülde aşağıdaki kodu çalıştırırız. Burda sadece VSTO projesinin adını ve ilgili metodun parametrelerinin değiştirmek yeterlidir.<br />
            <pre class="brush:vb">
Sub ExcelOkuveYaz()
    Dim addIn As COMAddIn
    Dim automationObject As Object
    Set addIn = Application.COMAddIns("VSTO_DigerOffice") 'Proje adı yazılıyor, sınıf adı olan VBAClass değil
    Set automationObject = addIn.Object
    automationObject.ExceldenOkuveAktiveCelleYaz "E:\OneDrive\Uygulama Geliştirme\web sitelerim\Yeni Efendi\Ornek_dosyalar\pivotdata.xlsx", "Sheet1"
End Sub</pre>
            </li>
            
        </ul>
        <p>Maalesef Late Binding ile ilerlemek durumundayız ve bu yüzden Intellisense&#39;ten de faydalanamıyoruz. İlgili metodu kullanmayı ezbere bilmek zorundayız. Ama tek kusuru da bu olsun. Bence yeri geldiğinde çok faydalı bi fonksiyonalite.</p>
     </div>
    <h2 class="baslik">Outlook add-in</h2>
    <div class="konu">
        <p>Excel'den sonra en çok add-in yapılan uygulama sanırım Outlook'tur. Biz de burda küçük bir Outlook Add-in yapacağız. Ancak yine Excel&#39;le ilişkiyi koparmayacağız. Yani bu add-inle Excel&#39;e kayıt atıyor olacağız.</p>
        
        <p>Pek yakında...</p>
    </div>
</asp:Content>
