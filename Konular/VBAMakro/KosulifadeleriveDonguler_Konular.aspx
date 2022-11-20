<%@ Page Title='KosulifadeleriveDonguler Konular' Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true"%>
<asp:Content ID="Content1" ContentPlaceHolderID="SayfaIcerik" Runat="Server">
         

<h1>Koşul ifadeleri ve Döngüler</h1>

<p>Koşullu ifadeler ve Döngüsel yapılar tüm programlama dünyasının olduğu gibi 
VBA dünyasının da kilit noktalarından ikisidir. Bunlar aslında <strong>
<a href="http://www.yazilimcilardunyasi.com/p/temel-bilgisayar-programlama.html">
Algoritma</a></strong> dediğimiz kavramın omurgasını oluştururlar. Bu bölümde "x=y ise şunu yap, 
değilse bunu yap" tarzı sorgulamaları ve "birşeyi yapmaya başla, bunu x defa yap 
veya şu şey olana kadar yap" gibi döngüsel yapıları göreceğiz.</p>

<div id="headerDatalist">Konular</div>
    <asp:DataList ID="DataList1" runat="server" DataSourceID="SqlDataSource1">
        <ItemTemplate>
            <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl='<%# String.Format("~/Konular/{0}/{1}_{2}.aspx", Eval("AnaKonu"), MyStatik.harfdonustur_cs(Eval("AltKonu").ToString()), MyStatik.harfdonustur_cs(Eval("Konu").ToString()))%>' Text='<%# Eval("Konu")%>'> 
            </asp:HyperLink>
        </ItemTemplate>
    </asp:DataList>

    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:YeniEE %>" SelectCommand="SELECT a.AnaKonu, t.AltKonu, k.Konu FROM Altkonular AS t INNER JOIN Konular AS k ON k.AltKonuID = t.AltKonuID INNER JOIN Anakonular AS a ON t.AnaKonuID = a.AnaKonuID 
where (t.AltKonu='Koşul ifadeleri ve Döngüler') and (a.AnaKonu='VBAMakro') order by KonuSiraNo"></asp:SqlDataSource></asp:Content>
