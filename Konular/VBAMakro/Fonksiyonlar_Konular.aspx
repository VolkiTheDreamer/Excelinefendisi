<%@ Page Title='Fonksiyonlar Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Fonksiyonlar</h1>
<p>Fonksiyonlar, köken olarak matematikteki fonksiyonlara dayanır. f(x)=3x+1 gibi bir fonksiyonda x bir parametredir ve bu fonksiyon parametre olarak aldığı sayının 3 katından 1 fazlasını döndürür. İşte programlama dillerindeki fonksiyonlar da bu şekilde, genelde bir parametre alırlar ve size birşey döndürürler.</p>
    <p>
Algoritma kavramından sonra en çok önem verilmesi gereken konudur. </p>
    <p>
Bu bölümde gerek VBA’in built-in fonksiyonlarını gerek kendi VBA fonksiyonlarımızı nasıl yazacağımızı göreceğiz. Beni en çok heyecenlandıran, kendimize ait Excel fonksiyonlarımızı nasıl yazacağımızı da yine bu bölümde göreceğiz.
</p>
<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Fonksiyonlar') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
