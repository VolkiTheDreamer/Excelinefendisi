var mypath="/"; //ziyaret için. buraya konular klasörü gelmeli, localhostta şişin bitince replace et.

function dateOfToday() {
var today = new Date();
var dd = today.getDate();
var mm = today.getMonth()+1; //January is 0!
var yyyy = today.getFullYear();

if(dd<10) {
    dd = '0'+dd
} 

if(mm<10) {
    mm = '0'+mm
} 

today = dd+'.'+mm+'.'+yyyy;
return today; //kısaca şöyle de olaabilir: new Date().toLocaleDateString()
}

function showHiddenTag(etiket, t) {
	if (t == "id") {
		document.getElementById(etiket).style.display = 'block';
		document.getElementById(etiket).style.visibility = 'visible';
	}
	else {
		var x = document.getElementsByClassName(etiket);
		var i;
		for (i = 0; i < x.length; i++) {
			x[i].style.display = 'block';
			x[i].style.visibility = 'visible';
		}
	}
}
    

function wrapperLink()
{
//bunlar işe yaramadı çünkü tüm event lsitenerları yokediyorlar, jqueryde yaptım
//document.body.innerHTML=document.body.innerHTML.split('UDF').join('<a href=../VBAMakro/Fonksiyonlar_UDFKullaniciTanimliFonksiyonlar.aspx>UDF</a>');
//cument.body.innerHTML=document.body.innerHTML.split('Workbook_Open').join('<a href=../VBAMakro/Olaylar_WorkbookOlaylarievent.aspx>Workbook_Open</a>');
}


/*function rezerve_kelime_renk_mavi() 
{
    var kac = document.getElementsByTagName("pre").length;
    for (var m = 0; m < kac; m++) {
        var rezv = ["And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Catch", "CBool", "CByte", "CChar", "CDate", "CDbl", "CDec", "Char", "CInt", "CLng", "CObj", "Const", "CStr", "Date", "Decimal", "Debug.Print", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "End", "Error", "Event", "Exit", "For", "Function", "GoTo", "If", "Integer", "Is", "IsNot", "IsNull","Let", "Like", "Long", "Loop", "Me", "Module", "New", "Next", "Not", "Nothing",  "Null", "Object", "On", "Option", "Optional", "Or", "OrElse", "ParamArray", "Private", "Property", "Public", "RaiseEvent", "ReadOnly", "ReDim", "REM", "Resume", "Select", "Set",   "Single", "Step", "String", "Sub", "Then", "Variant", "Wend", "When", "While", "With", "False", "True"];


        //alert(document.getElementsByClassName("pre1")[0].innerHTML.replace("\n"," \n"));
        var donusum = document.getElementsByTagName("pre")[m].innerHTML.replace("("," ").replace(")"," ").replace(/\n/g, " \n ").replace(/'/g, '<span class="yorum">\'').replace(/-ys/g, "</span>")
        //var donusum = document.getElementsByClassName("pre1")[0].innerHTML.replace("'", '<span class="yorum">\'').replace("-yorumsonu", "</span>")
        //var kelime = donusum.split(/\n| /);
        var kelime = donusum.split(" ");
        //alert(donusum);

        var yeni = "";


        for (var i = 0; i < kelime.length; i++) {
            var x = kelime[i];
            var a = (x == /\n/) ? "" : " ";
            //alert(kelime[i] + "-" + aranan.length);
            if (rezv.indexOf(x) > -1) {
                yeni = yeni.concat(a, '<span class="mavi">', x, '</span>');
            }
            else {
                yeni = yeni.concat(a, x);
            }
            //alert("kelimemiz:" + kelime[i]+ " ve boyu:" + kelime[i].length);  
            //alert("yeni dizimiz:" + yeni);

        }

        document.getElementsByTagName("pre")[m].innerHTML = yeni.substring(1, yeni.length - 1).replace(/\n /g, "\n");//baştaki 1 tane boşluğu atmak için
        //document.getElementById("yaz").innerHTML = yeni.substring(1, yeni.length - 1);
        //alert(yeni);
    }
}
*/
function html_tag_ekle(once, sonra)
{

    //var textarea = document.getElementById("ctl00_SayfaIcerik_txtEdit");
    var textarea = document.getElementById("SayfaIcerik_txtEdit");
    var bas = textarea.selectionStart;
    var bts = textarea.selectionEnd;
    var adet = textarea.value.length;

    var sabit_bas = textarea.value.substring(0, bas);
    var sabit_bts = textarea.value.substr(bts, adet);
    var secim = textarea.value.substring(bas, bts);
    textarea.value = sabit_bas + once + secim + sonra + sabit_bts;

    textarea.focus();
    textarea.setSelectionRange(bas+once.length,bas+once.length);

}

function iframeGuncelle(a,b,c) {
    //var ana = document.getElementById("#ctl00_SayfaIcerik_drpAnakonu").textContent;
    //var alt = document.getElementById("#ctl00_SayfaIcerik_drpAltkonu").textContent;
    //var konu= document.getElementById("#ctl00_SayfaIcerik_drpkonu").textContent;
    
    var yenisayfa = "../Konular/" + a + "/" + b+ "_"+ c +".aspx";
    document.getElementById('myIframe').src = yenisayfa;
}

//function storeinput(value) {
//    document.getElementById("<%=hidValue.ClientID%>").value = value;
//}

//function promptyap() {
//    var cevap = prompt("eski alt konu neydi");
//    document.getElementById("<%=hidValue.ClientID%>").innerText = cevap;
//}

function setCaretPosition(ctrl, pos)
{
 
	if(ctrl.setSelectionRange)
	{
		ctrl.focus();
		ctrl.setSelectionRange(pos,pos);
	}
	else if (ctrl.createTextRange) {
		var range = ctrl.createTextRange();
		range.collapse(true);
		range.moveEnd('character', pos);
		range.moveStart('character', pos);
		range.select();
	}
}
 
function process(txtid,pos)
{
	setCaretPosition(document.getElementById(txtid),bas);
}


// When the user clicks on the button, scroll to the top of the document
function topFunction() {
    document.body.scrollTop = 0; // For Chrome, Safari and Opera 
    document.documentElement.scrollTop = 0; // For IE and Firefox
}

function bottomFunction() {
    //document.body.scrollBottom = 0; // For Chrome, Safari and Opera 
    //document.documentElement.scrollBottom = 0; // For IE and Firefox
	window.scrollTo(0,document.body.scrollHeight);
}

function udemyFunction() {
	ga('send', 'event', 'Button', 'Click');
	alert("VBA/Makro-2 eğitimi için indirim kuponunu almak için iletişim sayfasından bana ulaşın");
	window.open("https://www.udemy.com/course/excel-makro-vba-egitimi-2ileri-seviye/?referralCode=29DFE807E63EBAA41CF1","_blank");
}

function mobilMenu() {
	var element = document.getElementById('mobil_anamenu'),
    style = window.getComputedStyle(element),
    deger = style.getPropertyValue('display');
    if (deger=='none'){
	    document.getElementById("mobil_anamenu").style.display = 'block';
	    }
  	else {
		document.getElementById("mobil_anamenu").style.display = 'none';
  	}
}

function mobilSideMenu() {

	var element = document.getElementById('sidebar');
    style = window.getComputedStyle(element);
    deger = style.getPropertyValue('display');
    if (deger=='none'){
	    document.getElementById("sidebar").style.display = 'block';
	    document.getElementById("sidebar").style.zIndex="999";
   	    document.getElementById("sidebar").style.width="30%"
   		document.getElementById("content").style.visibility = 'hidden';
	    }
  	else {
		document.getElementById("sidebar").style.display = 'none';
   		document.getElementById("content").style.visibility = 'visible';
  	}

}

//cookiler
function getCookieVal (offset) {
	var endstr = document.cookie.indexOf (";", offset);
	if (endstr == -1)
	endstr = document.cookie.length;
	return unescape(document.cookie.substring(offset, endstr));
}

function GetCookie (name) {
var arg = name + "=";
var alen = arg.length;
var clen = document.cookie.length;
var i = 0;
while (i < clen) 
	{
	var j = i + alen;
	if (document.cookie.substring(i, j) == arg)
		return getCookieVal (j);
	i = document.cookie.indexOf(" ", i) + 1;
	if (i == 0) 
		break; 
	}
return null;
}

function SetCookie (name, value) {
	var argv = SetCookie.arguments;
	var argc = SetCookie.arguments.length;
	var expires = (2 < argc) ? argv[2] : null;
	var path = (3 < argc) ? argv[3] : null;
	var domain = (4 < argc) ? argv[4] : null;
	var secure = (5 < argc) ? argv[5] : false;
	document.cookie = name + "=" + escape (value) +
	((expires == null) ? "" : ("; expires=" + expires.toGMTString())) +
	((path == null) ? "" : ("; path=" + path)) +
	((domain == null) ? "" : ("; domain=" + domain)) +
	((secure == true) ? "; secure" : "");
}

function espiricookie() {
///BURADA KLADIM, diğer bilgileri de yazdıralım, eğer daha önce birkez info bilgisini close butonua tıkladınysa bunu artık göstermesin. if espiri>0 then ilgili id visible ve display değiştir. ikinci buton için de aynısı olsun.
    var espiri;
    espiri=GetCookie("espiritest")
    espiri++;
    sitesontarih = new Date().toLocaleDateString(); //dateOfToday();
    SetCookie("espiritest", sitevisit, expdate, mypath, null, false);
}
   	                    
   	                    
function cookieUyariGoster() {
var sitevisit;
sitevisit = GetCookie("sitevisit");
//console.log("sitevisit değeri:"+sitevisit);
    //alert("hey");
    //alert(sitevisit);
    if (sitevisit == "0") //|| sitevisit == null 
	{			
		$("#cerez").show();
	}
}
   	                    
function DisplaySiteVisitInfo() {
//son tarihe bakarak, fark büyükse, bayağıdır yoktun nerelerdeyin, özledik seni....
	var expdate = new Date();
	var sitevisit;
	var sitesontarih;
	var mesaj;
	
	expdate.setTime(expdate.getTime() +  (24 * 60 * 60 * 1000 * 365));
	sitevisit = GetCookie("sitevisit");
	sitesontarih=GetCookie("sitesontarih");

	if (sitevisit == "0") 
		{			
			mesaj="Siteme hoşgeldin. Umarım senin için faydalı bir kaynak olur.";
		}
	else
		{
			var bas = new Date(sitesontarih).getTime();
			var timeDiff = Math.abs(new Date().getTime()-bas);
			var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
			if (diffDays>30)
				{
					mesaj="Hey, dostum, nerdesin, epeydir yoktun ortalıkta, özlettin kendini";
				}
			else
			{
				switch(parseInt(sitevisit)) 
				{
				  case 1:
				    mesaj="Siteme tekrar hoşgeldin.";
				    break;
				  case 2:
				    mesaj="Bakıyorum yine buralardasın. Güzeel, sitemi beğenmene sevindim.";
				    break;
				  case 3:
				    mesaj="Yine mi sen! E hadi biraz da çalış";
				    break;
				  case 4:
				    mesaj="Ooo müdavim, hoşgelmişsin";
				    break;
				  case 5:
				    mesaj="Sen bayağı Excel sevdalısı çıktın ya, hoşgeldin bakalım";
				    break;
				  case 6:
				    mesaj="Hey biliyo musun, bu dünyada benim sitemden başka siteler de var, ama ne yalan söyliyim bayılıyorum ya seni buralarda görmeye..";
				    break;
				  case 7:
				    mesaj="Buldun beleş siteyi tabi, kapağı attın buralara. Ohh değmesinler keyfine.";
				    break;
				  case 8:
				    mesaj="Yav dünyanın bilgisini edindin hala doymadın. Doy eccuk doy..";
				    break;
				  case 9:
				    mesaj="Bilgiyi aldıkça coşup şenleniyosun dimi. Lösev'e küçük bi bağış yap da çocuklar da şenlensin.";
				    break;				    
				  case 10:
					mesaj="Artık seni ortak yapabilirim bu siteye. Söylesene, hosting kiramın yarısını sen verir misin?";
				    break;
				  default:
					mesaj="Hoşgeldin ortak!";
				    break;
				}	
			}		
		}

	sitevisit++;
	sitesontarih=new Date().toLocaleDateString(); //dateOfToday();
	SetCookie("sitevisit", sitevisit, expdate, mypath , null, false);
	SetCookie("sitesontarih", sitesontarih, expdate, mypath , null, false);	

	return mesaj;
}

function DisplayPageVisitInfo() {
	var expdate = new Date();
	var urlAdres;
	var lastVisit;
	var visitAdet;
	var mesaj;
	var süre;
	var urlindeks;

	expdate.setTime(expdate.getTime() +  (24 * 60 * 60 * 1000 * 365));
	
	var bas = new Date("12/21/2018").getTime();
	var timeDiff = Math.abs(new Date().getTime()-bas);
	var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
	if (diffDays>300)
		{süre="son 1 yıldaki";}
	else if (diffDays>30)
		{süre="son "+ parseInt(diffDays/30+1)+ " aydaki";}
	else
		{süre="son "+ parseInt(diffDays/7+1)+ " haftadaki";}
	
	
	var pagevisit=GetCookie("pagevisit");
	if (!pagevisit) 
		{
			urlindeks=-1;
		}
	else
	{	
		var pagevisitArray=pagevisit.split('-');
		urlindeks=pagevisitArray.indexOf(window.location.href);
	}	
	
	//alert("url indeks:"+urlindeks);
	if (urlindeks>=0) //varsa
	{
		pagevisitArray[urlindeks+1]=new Date().toLocaleString(); //son visiti update, toLocaleDateString de var.
		var ziyaret=parseInt(pagevisitArray[urlindeks+2])+1;
		pagevisitArray[urlindeks+2]= ziyaret//ziyareti 1 artır
		//şimdi diziyi tekrar string yapalım
		pagevisit=pagevisitArray.join("-");

		mesaj=mesajGoster(ziyaret) + "\n\n Bu sayfayı en son " + pagevisitArray[urlindeks+1]+ " saatinde ziyaret etmişsin. Şimdikiyle birlikte " + süre + " ziyaret sayın " + ziyaret + " olmuş." ;
	}
	else //ilk kez cookie yaratılacaksa veya cookiler dolmuş ama bu sayfa içeride yoksa, yeni eklenecek
	{
		urlAdres=window.location.href;
		lastVisit=new Date().toLocaleDateString();
		visitAdet=1;
		if (pagevisit=="")
			{pagevisit=urlAdres+"-"+lastVisit+"-"+visitAdet;}
		else
			{pagevisit=pagevisit+"-"+urlAdres+"-"+lastVisit+"-"+visitAdet;}
		mesaj="Bu sayfaya ilk ziyaretin, hoşgelmişsin, sefalar getirmişsin.";			
	}

	//set işleminden önce değerleri almak lazım, o yüzden bu aşama en son
	SetCookie("pagevisit", pagevisit, expdate, mypath , null, false); 
	//alert(mesaj);
	return mesaj;
}

function mesajGoster(adet) {	
	var message;
/*	if(adet == 1) 
	message="Sayfaya hoşgeldin!";
*/	if(adet == 2) 
	message="Bakıyorum da yine gelmişsin!";
	if(adet == 3) 
	message="Vaay, yine mi sen!";
	if(adet == 4)
	message="Sen bu konuda oldukça meraklısın galiba!"; 
	if(adet == 5) 
	message="Sen resmen bu sayfanın müdavimi oldun çıktın ha!";
	if(adet == 6) 
	message="Hey, git kendine bi hobi bul artık!";
	if(adet == 7)
	message="Yapacak başka işin yok mu, niye sürekli bu sayfadasın? Git başka sayfalarıma bak, onlar da çok güzel."; 
	if(adet == 8) 
	message="Hiç uyuduğun oluyor mu? Rüyanda da mı bu konuları görüyorsun yoksa?";
	if(adet == 9)
	message="Git hayatını yaşa dostum, Excel'den daha mühim şeyler de var bu hayatta!"; 
	if(adet >= 10) 
	message="Hey, her ayın 1'nde kiranı yatırman gerekiyor, buralarda takılmaktan bunu unutma sakın!";
	if(adet >= 11) 
	message="Güzel güzel takılıyorsun, bari biraz yorum yaz da site şenlensin!";
	if(adet >= 11) 
	message="Galiba, senden kira almam gerekecek, nedir bu sayfanın senden çektiği!!!";
	if(adet >= 12) 
	message="Ya çok kalın kafalısın da konuyu hala anlamadın, ya da bu sayfaya aşık oldun. Hangisi?";
	if(adet >= 13) 
	message="Benim sana diyecek bi lafım yok artık. İstiyosan sayfayı üzerine yapayım.";
	if(adet >= 14) 
	message="Hoşgeldin ortak!";
		
	return message;
}

function ResetCountsSite() {
var expdate = new Date();
expdate.setTime(expdate.getTime() +  (24 * 60 * 60 * 1000 * 365)); 
sitevisit = 0;
sitesontarih="";
SetCookie("sitevisit ", sitevisit , expdate , mypath, null, false);
SetCookie("sitesontarih", sitesontarih, expdate , mypath, null, false);
history.go(0);
}

function ResetCountsPage() {
var expdate = new Date();
expdate.setTime(expdate.getTime() +  (24 * 60 * 60 * 1000 * 365)); 
pagevisit="";
SetCookie("pagevisit", pagevisit, expdate , mypath, null, false);
history.go(0);
}


function junkKontrol() {
var mesaj="Az sonra, hem bu sitenin mail adresinden hem de şahsi mail adresim olan volkan.yurtseven@hotmail.com ";
mesaj=mesaj + "adresinden mail almış olacaksınız. Maillerin junk'a düşmediğinden emin olmak için posta kutunuzu kontrol edin.";
mesaj=mesaj + "NOT: Yeni sayfa bildirimlerini bir süre şahsi mailimden alacaksınız.";

alert(mesaj);
}

