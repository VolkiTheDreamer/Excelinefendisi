<%@ Page Title='Taskpane' Language='C#' MasterPageFile='~/MasterPage.master' AutoEventWireup='true' %>

<asp:Content ID='Content1' ContentPlaceHolderID='SayfaIcerik' Runat='Server'><div id='gizliforkonu'><table><tr><td><asp:Label ID='Label1' runat='server' Text='VSTO'></asp:Label></td><td><asp:Label ID='Label2' runat='server' Text='Görsel Araçlar'></asp:Label></td><td><asp:Label ID='Label3' runat='server' Text='2'></asp:Label></td></tr></table></div>
    <h1>Taskpane</h1>
    
    <p>Taskpane kullanımı oldukça basittir. Peki nerede kullanırız. Aslında 
	aklınıza gelebilecek herşey için kullanabilirsiniz. Excel'in kendi built-in 
	taskpanelerini düşünün. Onlardan herhangi birine benzer bir amacınız 
	olabileceği gibi, ribbonlardaki çeşitli seçenekleri Taskpane'den de vermeyi 
	tercih edebilirsiniz. Mesela HTML dünyasındaki CSS'lere benzer bir 
	yapıyı burada yaratabilrisiinz. Böylece Excel&#39;in hazır styling şablonlarını 
	uygulamak yerine kendinize ait yeni şablonlar yaratabilir ve bunları hızlıca 
	tablolarınıza uygulayabilrsiniz. Veya Excel açıldığı sırada bir veritabanından çeşitli değerleri okuyabilir ve taskpane üzerinden
	bunlara ait işlemler yapabilirsiniz. Taskpaneden hızlı copy-paste yapmak gibi.</p>
	
	<p>Dediğim gibi, aklınıza gelecek herşeyi yapabilirsiniz. Hadi hızlıca ne yapmak gerekiyor ona bakalım.</p>
	
	<h2><strong>Yaratım</strong></h2>
	<p>Malesef Toolbox içinden sürükle bırak şeklinde veya New Project Item deyip Ribbon yaratır gibi TaskPane yaratamıyoruz. Bunun
	için birkaç parça kod yazmamız gerekiyor.</p>
	<h3><span style="font-weight: normal">User Control</span></h3>
	<p>Öncelikle, Project menüsüne sağ tıklayıp New Item diyerek bir adet
	<strong>User Control</strong> ekliyoruz. Bunun adına MyUserControl diyelim. Bu nesne, aslında bir Form nesnesi gibidir. İçine her tür form kontrolü konabilir.</p>
	<p>Şimdi bu usercontrolün içeriğini istediğimiz kontrollerle dolduralım. 
	Basit örnek olması adına şimdilik sadece bir buton ve combobox koyalım.</p>
	<p><img alt="Usercontrol" src="../../images/Vsto_taskpaneusercontrol.jpg"></p>
	<h3><span style="font-weight: normal">Kod</span></h3>
	<p>İlk olarak ister <strong>ThisAddin </strong>içine ister <strong>Ribbon </strong>içine(hangisi ihtiyacımızı görüyorsa)
	ana nesne değişkenlerini tanımlarız, sonra da nerden tetikleyeceksek orada taskpane yaratım kodlarını. Ben Ribbondan açmayı düşündüğüm için
	ribbona bir buton ekledim, global değişkenlerimi de ona göre olşuturacağım.</p>
	
	<img src="../../images/taskpane1.jpg">
	<pre class="brush:csharp">
public partial class Ribbon1
    {
        public MyUserControl myusercontrol1; //taskpane için
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane; //taskpane için
    
        //Aradaki diğer kodlar
	
	private void button30_Click(object sender, RibbonControlEventArgs e)
        {
            this.myusercontrol1 = new MyUserControl();
            this.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(this.myusercontrol1, "İlk Task Pane");
            this.myCustomTaskPane.Visible = true;
            this.myCustomTaskPane.Width = 200;
        }	
    }  </pre>
	
<p>Bu butona tıkladığımızda sağ tarafta TaskPane'imiz açılır.</p>	

<img src="../../images/taskpane2.jpg">

<h2><strong>Kullanım</strong></h2>
<p>Şimdi bu TaskPane ile basit birkaç iş yapalım.</p>
<p>Öncelikle diyelim ki, taskpane açılır açılmaz, combobox&#39;ın içeriği dolsun. Butona basılınca da comboboxta seçili olan değeri aktif hücreye yazdıralım.</p>
<pre class="brush:csharp">
    public partial class MyUserControl : UserControl
    {
        public Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
        public MyUserControl()
        {
            InitializeComponent();
        }

        private void MyUserControl_Load(object sender, EventArgs e)
        {
            this.comboBox1.Items.Add(1);
            this.comboBox1.Items.Add(2);
            this.comboBox1.SelectedIndex = 0; //ilk değeri seçiyoruz
        }

        private void button1_Click(object sender, EventArgs e)
        {
            app.ActiveCell.Value = this.comboBox1.SelectedItem.ToString();
        }
    }</pre>

<h2>Daha kompleks bir örnek</h2>
<p>Bu örnekte Taskpane'imizi hem içerik olarak zenginleştireceğiz. Hem de .Net dünyasının nimetlerinden faydalanacağız. </p>
    <p><strong>EDIT:</strong>Bu örneğin yerini değiştirip Örnek Projeler içine almaya karar verdim. Zira burada aşağıdaki ileri c# konularının kullanımı da sözkonusu.</p>
	<ul>
		<li>Async/await keywordleri</li>
		<li>DataGridView</li>
		<li>LINQ syntaxı</li>
	</ul>
	</asp:Content>
