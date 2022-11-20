<%@ Page Title='DortTemelNesne Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Dört Temel Nesne</h1>
<p>Excel, Nesne Modeline dayanan bir yapıya sahiptir demiştik. Bu, şu 
demek:Exceldeki herşey bir nesnedir ve bu nesneler hiyerarşik bir yapı içinde 
bulunurlar. Çok sayıda nesne olmakla birlikte en sık kulanacağımız nesneler 
aşağıdaki dörtlüdür. </p>

	<p>NOT:Eğer bölümlerde sırayla ilerleyen biriyseniz, artık gözünüz aydın 
	diyebilirim, çünkü bu bölümden itibaren kod yazımında oldukça ivme kazanacağız. 
	O yüzden 
	nerdeyse tüm makrolarda bulunan koşullu yapıları ve döngüleri de sık sık kullanacağız. VBA yerine VB.NET anlatıyor olsaydım bunları daha önce anlatırdım. Ancak VBA olunca işin rengi biraz dğeişiyor. O yüzden önce <a href="KosulifadeleriveDonguler_Konular.aspx">bu konulara</a> bakın demek istemiyorum, çünkü ordaki şeyleri anlamak için de bu dört nesneyi de az çok tanımak gerekiyor. O yüzden bi oraya bi buraya bakarak da ilerleme yolunu tercih edebilirsiniz.</p>

<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Dört Temel Nesne') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
