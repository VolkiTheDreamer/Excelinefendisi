<%@ Page Title='FormulasMenusu1 Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Formulas Menüsü-Fonksiyonlar</h1>
<p>Fonksiyonlar, Excelin en temel araçlarıdır. Onlarsız bir Excel düşünmek mümkün değildir. Ayrı bir uzmanlık gerektirdiği için burada önemli olduğunu düşündüğüm tüm fonksiyonları detaylı bir şekilde ele almaya çalıştım. Verdiğim örnekleri, tüm site boyunca benimsemeye çalıştığım genel yaklaşıma göre yani gerçek dünya örneklerine göre kurgulamaya çalıştım. </p>

<p>Bu bölümde, Finansal ve Matematiksel fonksiyoları düşünmezsek kalan tüm fonksiyonların en az %90ını ele almışımdır(Saymadım ama buna yakındır). Bununla birlikte Excel'in yerleşik fonksiyonlarının yeterli olmadığı veya çok uzun formüller yazmanız gerektiği durumlarda gerek kendi yazdığım UDF'leri verdim, gerekse kendi UDF'lerinizi yazma konusunda sizi teşvik etmeye çalıştım. Umarım bu teşviklerim işe yarar ve UDF yazmayı öğrenirsiniz. UDF yazmak, VBA bilmeyi gerektirdiği için öncelikle temel VBA konularını öğrenmenizi tavsiye ederim.</p>

<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Formulas Menüsü(Fonksiyonlar)') and (a.AnaKonu='Excel') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
