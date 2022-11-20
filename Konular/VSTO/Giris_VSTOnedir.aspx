<%@ Page Title='VSTO nedir' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %><asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'></div><h1>VSTO nedir?</h1>
    
    <h3>Nedir?</h3>
    <p>VSTO, açık adıyla Visual Studio Tools for Office, .Net framework&#39;ten yararlanarak Excel Add-in&#39;lerine kıyasla daha profesyonel görünümlü ve daha performanslı add-inler yapmamızı sağlayan teknolojidir. <a href="../VBAMakro/Ileriseviyekonular_Add-InlerveCustomMenuler.aspx">Excel Add-inler sayfasında</a> üç tür add-inden bahsetmiştik. VSTO add-inler, COM Add-in olarak da geçer, eğer javascriptle haşır neşirliğiniz yoksa bence gelmeniz gereken son noktadır. Yine de son teknoloji olan web add-inlerle VSTO arasındaki farkları görmek isterseniz şöyle <a href="http://techgenix.com/comparing-vsto-and-office-web-add-ins-video/">güzel bir karşılaştırma sayfası</a> var, buraya bakabilirsiniz.</p>

    <h3>Neden VSTO?</h3>
    <p>Ne zaman Excel Add-in ne zaman VSTO Add-in hazırlanır diye sorarsanız, bunun net bir cevabı olmamakla birlikte, şöyle özetleyebilirim:</p>
    <ul>
        <li>Add-in&#39;inizi herkese tek tek dağıtımla uğraşmak ve nasıl kurması gerektiğini anlatmak istemiyorsanız</li>
        <li>Yaptığınız bir güncellemenin herkeste otomatikman güncellenmesini istiyorsanız</li>
        <li>.Net framework&#39;ün gücünden faydalanmak istiyorsanız(Bunları zamanla keşfedeceğiz)</li>
        <li>Yazdığınız kodların kimse tarafından görülmesini istemiyorsanız</li>
        <li>Ticari amaçlı birşeyler yazmak istiyorsanız</li>
    </ul>

    <p>
        Daha detay için <a href="https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa192490(v=office.11)">buraya</a> bakabilirsiniz.</p>

    <h3>Ne lazım?</h3>
    <ul>
        <li>Excel Nesne Modeline hakimiyet ve tercihen VBA bilgisi</li>
        <li>Visual Studio(Nasıl kuracağımız göreceğiz)</li>
        <li>VB.Net veya C#(işimize yarayacak kadar olan kısmını burada öğreneceğiz, geliştirmek size kalmış)</li>
    </ul>
 
    <h3>Neye benziyor?</h3>

        <p>Aslında, sitemde daha önce gezindiyseniz Excelent sayfama da uğramış olmanız mümkündür. Yapacağımız add-inler (genelde) buna benzer bir ribbon üzerinden kullanılmakta. Ribbondan açılan bir form ile sıradan bir masaüstü program gibi bir add-in'iniz bile olabilir(hatta bunu ribbondan bağımsız da yapabilirsiniz, Excel başlar başlamaz da açılabilir ama genelde ribbona adreslenir). Aslında, genel bakıldığında fonksiyonalite açısından VSTO add-in'inizin normal bir masaüstü programdan hiçbir farkı yoktur, küçük bir detay hariç: Add-inler, çalışması için Excel'e ihtiyaç duyar. (Teknik dille ifade etmek gerekirse yazacağımız add-inler bir <strong>dll</strong> dosyası üretir, normal masaüstü programlar ise exe dosyası üretirler)</p>
        <p>Bu add-inimizi bir kurulum dosyası olacak şekilde derleyip, bunu bir web sitesi üzerinden, şirketinizin networkü üzerinden herkesin erişebildiği bir alandan veya daha geleneksel yöntemler olan, CD/DVD&#39;den kurulacak şekilde dağıtabilirsiniz. Mesela benim Excelent, web sitesi üzerinden servis edilmektedir.</p>
        <p>Şimdi sırayla .Net çatısını tanıyalım, akabinde Visual Studio&#39;yu kuralım.</p>
   
</asp:Content>
