var d = jQuery.noConflict();


d(document).ready(function () {

                //Herşeyden önce içindekiler jquerysi
			//d('#content').jqTOC();
		 	d(window).resize(function() {

			if (d(window).width() < 700) 
				{
				mobilEkran();
				}
			else 
				{
					PCEkran();
					location.reload();//resim jquerysi çalışmıor diye
				}
			});
		
			if (d(window).width() < 700) 
				{
				mobilEkran();
				}
			else 
				{
					PCEkran();
				}
					
		function mobilEkran() {			
			//d("#anacontainer").css("width", "100%");
			d("#content").css({ "width": "100%", "margin-left": "0", "padding-top": "2%"});		  	
			d(".nomobilsegizle").css("display", "block");
			d(".yankutuluarmobildegizle").css("display", "none");

			//d("#navhome").css("left", "1%");		  	
			d("#navhome").css("visibility", "hidden");		  	
			d("#navend").css("visibility", "hidden");		  	
			d("#udemy").css("visibility", "hidden");		  					  					  	
			d("#ziyaret").css("visibility", "hidden");		  					  					  	
			d("#yazilar").css("padding-left", "5%");
			d("#ortablok_footer").css("padding-left", "7%");
			
		}	  

		function PCEkran() {
			//tasarım nedeniyle başta gizlediklerimi gösterme işi
			d("#sidebar").css("display", "block");
			d("#anamenu").css("display", "block");
			d(".yankutuluarmobildegizle").css("display", "block");
			d("#udemydestek").css("display", "none"); //burayı bazen toggle yap
			//d("#yazilar").css("margin-left", "15%");

			d("#anacontainer").css("width", "70%");
			d("#content").css("width", "81%");
			d(".nomobilsegizle").css("display", "none");			  	
			d("#navhome").css("left", "10%");		  	
			d("#navend").css("right", "10%");		  			  	
			d("#udemy").css("top", "0%");	
				  	
			d("h1").showBalloon(
						    	        {contents:"Tüm konuları(koyu maroon arkaplan başlıklı) açmak/daraltmak için ana başlığa çift tıklayın",
										showDuration: "slow",
										position:"top",
										showAnimation: function(d, c) { this.fadeIn(d, c);}
										}
										);
			setTimeout(function()
	                {
			    d("h1").hideBalloon();
			    d("#duyuruimg").hide();
   	            d("#ziyaret").hide();
	                },10000);
	                       
			}
		 
		 function getContent(nereye,sayfa_ve_id) {
			 d(nereye).load(sayfa_ve_id);
		 }
///BAŞLIK, COLLAPSE, TOGGLE
		//ilk girişte kapalı gelen tüm h2 başlıklar 1 sn sonra açılsın ve h1 seviyesinde uyarı çıksın,ama anchora tıkalndıysa çalışmasın
               if (window.location.hash =='')
                {
					setTimeout(function()
						{
						d("h2.baslik").each(function()
							{
								d(this)
					            .next()
					            .slideToggle("slow");
					            
						    });
						}, 500);
			
                }
                


//d("a[id!=ctl]").attr("target", "_blank");   //bunu daha sonra şöyle yap, eğer id'sinde ctl varsa kendinde çasın değilse yeni sayfada açsın

		//başlık classına sahip tüm başlıklar(h2,h4) click olunca kendisinden sonra gelen(NEXT) divleri toggle yapsın
        d(".baslik").click(function () {
            d(this)
            .next()
            .slideToggle("slow");
            });

        d("#newComment").click(function () {
            d(this)
            .next()
            .slideToggle("slow");
            });
        
        /*d('.rate').on('change', function() {		
			var deger=d('input[name="rate"]:checked').val();
			alert(deger);
			d("#<%=hiddenalan.ClientID%>").val(deger);
			alert(deger+10);
			alert(d("#<%=hiddenalan.ClientID%>").val());
		});
*//* çalışmıyor, zira galiba jquery en başta yüklendiği için ve henüz mailtopla oluşmadığı için görmüyor.
olmaz, copycontent de masterpagede yükleniyor ama o çalışıyor, başka bir sorun olmalı.
       d(".mailTopla").click(function () {
        	alert("test");
            d(".mailTopla").hide();
            });            
*/
		d("#duyuruimg").click(function () {
            d(this).css("display", "none");
            d(this).prev().css("display", "inline");
            d(this).prev().fadeIn("slow");//ilan yavaşça görünsün, ama çalışmıyor
            d(this).prev().delay(20000).fadeOut("slow");//ilan yavaşa kaybolsun
/*            d(this).prev().show("slide", { direction: "right" }, 1000);*/
        });

		d("#duyuruilan").click(function () {
            d(this).css("display", "none");
		});
	d("#udemydestek").click(function () {
		d(this).css("display", "none");
	});
		d("#cerez").click(function () {
            d(this).css("display", "none");
        });

		d("#ziyaret").click(function () {
            d(this).css("display", "none");
        });

	
		//h1'lere çift tıklanınca aynı seviyedeki(SIBLING) tüm "konu" classlı divler toggle olsun
        d("h1").dblclick(function () {
            d(this)
            .siblings(".konu")
            .slideToggle(200);
            });
            
        //h1 üezrine gelince balon çıksın    
		d("h1").balloon({
		  contents:	"Tüm konuları(koyu maroon arkaplan başlıklı) açmak/daraltmak için buraya çift tıklayın",
		  showDuration: "slow",
		  position:"null",
		  showAnimation: function(d, c) { this.fadeIn(d, c); }
		});

        //akıllı başlık üezrine gelince balon çıksın    

		d(".pre_baslik").balloon({
		  contents:	"İçeriği wrap/unwrap yapmak için çift tıklayın",
		  showDuration: "slow",
		  position:"left",
		  showAnimation: function(d, c) { this.fadeIn(d, c); }
		});
		
		d(".CopyContent").balloon({
		  contents:	"Kodun içine çift tıklayın ve kopyalayın.",
		  showDuration: "fast",
		  position:"left",
		  showAnimation: function(d, c) { this.fadeIn(d, c); }
		});

		d(".syntaxhighlighter  vb").children().balloon({
		  contents:	"Kodun içine çift tıklayın ve kopyalayın.",
		  showDuration: "fast",
		  position:"left",
		  showAnimation: function(d, c) { this.fadeIn(d, c); }
		});
		
		

//div[class^="jander"]
//d("[id*=highlighter]").balloon({                
        d("div[class^=highlighter]").balloon({
		  contents:	"İçeriği wrap/unwrap yapmak için başlığa çift tıklayın",
		  showDuration: "slow",
		  position:"left",
		  showAnimation: function(d, c) { this.fadeIn(d, c); }
		});
		


/*		d('.SmartHeaderContainer').each(function()
					{
						d(this)
						.next()
						.find("vb spaces")
						.text().replace(new RegExp(String.fromCharCode(160),"g"),"xxxx");
					});
*/

/*		d('.vb .spaces').each(function()
					{
						d(this)
						.html().replace(/&nbsp;/g,"@");
					});
*/
		

/*		d('.vb.spaces').each(function()
			{
				d(this).text(d(this).html().replace(/&nbsp;/g," "));
			});
*/

/*		d('.marquee').marquee();*/
		
///REPLACER , WRAPPER
		//tüm PRE'lerdeki renklendirmeler, tüm P içindeki UDF v.s de wrapper olup hyperlink oluşsun		
		//ikisini ayrı ayrı yazınca ikincisi çalışmıyor, o yüzden tek fonk içine yazdım
		
		/*PRE içeriği
			tırnaklar
			formüller
			özel işaretler
			// ile başlayan yorumlar
		*/
		d('p, pre.formul, span').each(function(){ 
		   var p = d(this);
           
		if(p.prop("tagName")=='PRE')
			 {
			
				var fonk0=["TEXT","SUM","ROUND","COUNT","MIN","MAX","IF","IFS"] //bunların hem prefixlisi hem suffixlisi var diye ayrıca ele aldım				
				var fonk1=["YEARFRAC","YEAR","NETWORKDAYS","WORKDAY","ISOWEEKNUM","WEEKNUM","WEEKDAY","VLOOKUP","DATEVALUE","VALUE","UPPER","UNICODE","UNICHAR","TRUNC","TRIM","TREND","TRANSPOSE","TODAY","TEXTJOIN","SUMPRODUCT","SUMIFS","SUMIF","SUBTOTAL","SUBSTITUTE","STDEV","SMALL","SKEW","SIGN","SEARCHB","SEARCH","ROWS","ROW","ROUNDUP","ROUNDDOWN","RIGHTB","RIGHT","REPT","REPLACEB","REPLACE","RATE","RANK","RANDBETWEEN","RAND","NPV","PV","PROPER","PRODUCT","POWER","PMT","PERCENTILE","OFFSET","ODD","NPER","NOW","NORMDIST","MROUND","EOMONTH","MONTH","MODE","MOD","MINUTE","MINIFS","MINA","MIDB","MID","MEDIAN","MAXIFS","MAXA","MATCH","LOWER","HLOOKUP","LOOKUP","LENB","LEN","LEFTB","LEFT","LARGE","KURT","INT","INDIRECT","INDEX","HOUR","GETPIVOTDATA","FV","FORMULATEXT","FORECAST","FLOOR","FIXED","FINDS","FIND","EXACT","EVEN","EDATE","DSUM","DMIN","DMAX","DCOUNTA","DCOUNT","DAYS","DAY","DAVERAGE","DATEDIF","DATE","COUNTIFS","COUNTIF","COUNTBLANK","COUNTA","COUNT","CORREL","CONCATENATE","CONCAT","COLUMNS","COLUMN","CODE","CLEAN","CHOOSE","CHAR","CEILING","AVERAGEIFS","AVERAGEIF","AVERAGEA","AVERAGE","ASC","AREAS","ADDRESS","ABS"];				
				var fonk2 = ["SWITCH","NOT","ISTEXT","ISREF","ISODD","ISNUMBER","ISNONTEXT","ISNA","ISLOGICAL","ISFORMULA","ISEVEN","ISERROR","ISERR","ISBLANK","IFNA","IFERROR","CELL","AND"];
				var ozeller= ["{", "}"];
				
				//önce tırnaklar
	            p.html(p.html().replace(/"/g, "<span class='formulozelisaret'>\"</span>"));	
	
				//sonra IFERROR ve OR( ayrıca ele alınıyor
	            p.html(p.html().replace(/IFERROR\(/g, "<span class='formulrenk2'>IFERROR</span>\("));	
	            p.html(p.html().replace(/OR\(/g, "<span class='formulrenk2'>OR</span>\("));	
	
				//şimdi de diğerleri

	           for (var i = 0; i < fonk1.length; i++)           
	           {
	             var rf = new RegExp(fonk1[i],"g");
	             p.html(p.html().replace(rf, "<span class='formulrenk1'>"+fonk1[i]+"</span>"));	
	           }
	
	           for (var i = 0; i < fonk2.length; i++)           
	           {
	             var rf = new RegExp(fonk2[i],"g");
	             p.html(p.html().replace(rf, "<span class='formulrenk2'>"+fonk2[i]+"</span>"));	
	           }
	
	           for (var i = 0; i < ozeller.length; i++)           
	           {
	             var rf = new RegExp(ozeller[i],"g");
	             p.html(p.html().replace(rf, "<span class='formulozelisaret'>"+ozeller[i]+"</span>"));	
	           }
	
	           for (var i = 0; i < fonk0.length; i++)           
	           {
	             var rf = new RegExp(fonk0[i],"g");
	             p.html(p.html().replace(rf, "<span class='formulrenk1'>"+fonk0[i]+"</span>"));	
	           }	
	
				//şimdi de // ile başlayan yorumları halledelim
	       
				var str = p.html();
				var s= /\n/
				if (str.search(s)==-1) //satırsonu yoksa
				{
					var bas=str.indexOf("\/\/")
					if (bas!=-1){
				        var str2= str.slice(bas);
						p.html(p.html().replace(str2, "<span class='formulaciklama'>"+str2+"</span>"));}
				}
				else //Enter yapıldıysa
				{
					var bas=str.indexOf("\/\/")
					if (bas!=-1){	
						var dizi = str.split(s);
						for (var i=0; i < dizi.length;i++) 
						{
		         		   var basy=dizi[i].indexOf("\/\/")
				           var str2= dizi[i].substr(basy);
						   p.html(p.html().replace(str2, "<span class='formulaciklama'>"+str2+"</span>"));
						}
					}			
				}	
	        }
		else if (p.html().indexOf("aspx") == -1 && p.html().indexOf(".pdf") == -1 && p.html().indexOf("https://github.com/VolkiTheDreamer") == -1) //github kısmını 10.04.21de ekledim
			{
			//wrapperlar-------------	
//sadece p içinde değil tüm elementlerde çalışsın
			  
//dictionary tanımlıyoz
/*bilgi:şu aşağıdaki eirişim şekilleri aynıdır
myObj["SomeProperty"];
myObj.SomeProperty;*/
	  
			  var spesifikler = [
			      { k: "UDF", v: '<a href=/Konular/VBAMakro/Fonksiyonlar_ExcelicinUDFKullaniciTanimliFonksiyonlar.aspx>UDF</a>' },
			      { k: "Workbook_Open", v: '<a href=/Konular/VBAMakro/Olaylar_WorkbookOlaylari.aspx>Workbook_Open</a>'},
				  { k: " Excelent", v: ' <a href=/Excelent.aspx>Excelent</a>' },
			      { k: "Personal.xlsb", v: '<a href=/Konular/VBAMakro/Giris_MakroNedir.aspx#personal>Personal.xlsb</a>' },
			      //{ k: "UserForm", v: '<a href=/Konular/VBAMakro/Formlar_Konular.aspx>UserForm</a>' },
			      { k: "Hata Yakalama", v: '<a href=/Konular/VBAMakro/DebuggingveHataYonetimi_HataYakalama.aspx>Hata Yakalama</a>' },			      			      			      
			      { k: "Filling", v: '<a href=/Konular/Excel/HomeMenusu_Doldurma.aspx>Filling</a>' },			      			      			      
			      { k: "NamedRange", v: '<a href=/Konular/Excel/FormulasMenusuDiger_NameManager.aspx>Name</a>' },	      			      			      			        
				  { k: "Early Binding", v: '<a href=/Konular/VBAMakro/Ileriseviyekonular_ObjelerDunyasi.aspx#binding>Early Binding</a>' }	,		  
				  { k: "Late Binding", v: '<a href=/Konular/VBAMakro/Ileriseviyekonular_ObjelerDunyasi.aspx#binding>Late Binding</a>' },
				  { k: "VolkansUtility", v: '<a href="https://github.com/VolkiTheDreamer/dotnet/tree/master/Ugulamalar/VolkansUtility">VolkansUtility</a>' }
				];

				var sayfa=location.pathname;				
				for (var i = 0; i < spesifikler.length; i++) 
				{
					var kaynak=new RegExp(spesifikler[i].k,'g');						
					var hedef=spesifikler[i].v;
					//aynı sayfadaysam wrap yapma, saçma olur, ama aynı sayfada bi anchora göndereceksem yap
					if (hedef.indexOf(sayfa) == -1 || (hedef.indexOf(sayfa) > -1 && hedef.indexOf("#") > -1)) 
					{
						p.html(p.html().replace(kaynak,hedef));  
					}
				}
				
				//detay açıklama ve abbrevatioanlar
				var abbrler=[
				{k:"ZLS", v:'<abbr title="Zero Length String-Sıfır Uzunluklu metin">ZLS</abbr>'},
				{k: "default", v: '<abbr title="Varsayılan">default</abbr>' },
				//{k: "VS", v: '<abbr title="Visual Studio">VS</abbr>' },
				{k:"recursive", v:'<abbr title="İteratif(yinelemeli)">recursive</abbr>'},
				{k:"scheduled", v:'<abbr title="Zamanlanmış, planlanmış">scheduled</abbr>'},				
				{k:"syntax", v:'<abbr title="Fonksiyon veya metodun yazım şekli">syntax</abbr>'},				
				{k:"case-sensitive", v:'<abbr title="Küçük/büyük harf ayrımına duyarlı">case-sensitive</abbr>'},						
				{k:"Değişim derecesi:", v:'<abbr title="Düşük seviye ya hiç değişim olmamasını veya birkaç karakterlik düzeltme işlemini ifade eder; orta seviye 100 karaktere kadar olan değişimleri ifade eder; daha fazla değişim olması Yüksek seviye ile ifade edilir">Değişim derecesi:</abbr>'},						
				{k:"indirebilirsiniz", v:'<abbr title="Uzantısı xlsx olan tüm dosyaları güvenle indirebilirsiniz. Bunlar makro içer(e)mediği için içiniz rahat olsun">indirebilirsiniz</abbr>'}												
				];

				for (var i = 0; i < abbrler.length; i++) 
				{
					var kaynak=new RegExp(abbrler[i].k,'g');						
					var hedef=abbrler[i].v;
					p.html(p.html().replace(kaynak,hedef));  					
				}

	  
				//aşağıdaki gibi yapınca, bodydeki tüm evetn listenelrları yok ediyor, ama sadece p'lerde yapınca h1lere, basliklara bişey olmuyor		
				//d("body").html(d("body").html().replace(/UDF/g,'<a href=../VBAMakro/Fonksiyonlar_UDFKullaniciTanimliFonksiyonlar.aspx>UDF</a>'));
				//d("body").html(d("body").text().replace(/UDF/g,'zuzu'));
				//d("body").html(d("body").html().replace(/Workbook_Open/g,'<a href=../VBAMakro/Olaylar_WorkbookOlaylarievent.aspx>Workbook_Open</a>'));
			}								
		});

///DİĞER
		//akıllı başlığa çift tıkalyınca alttaki PRE wordwrap olayı toggle
        d(".pre_baslik").dblclick(function () {
        if (d(this).parent().next().find(".syntaxhighlighter .line").css("white-space" )=="pre")
           {
             d(this).parent().next().find(".syntaxhighlighter .line").attr("style", "white-space: normal !important;");
           }
         else
           {
             d(this).parent().next().find(".syntaxhighlighter .line").attr("style", "white-space: pre !important;");
           }
        
        });
        
/*        d(".CopyContent").click(function () {
        	alert("Kodun içine çift tıklayın ve kopyalayın");
        });
        
*//*		d(".CopyContent").click(function () {
			 var elm = d(this).parent().next().find("td.code");
//			 var elm = d(this).parent().next().find(".line .number1 .index0 .alt2");
		//	 alert(elm.text());
			 //copyClip(elm.text()); gerek kalmadı, tam çalışmıyordu zaten. doubleclick ile çalışıyor.
//			 elm.dispatchEvent(new MouseEvent('dblclick', {'bubbles': true}));
			 //document.execCommand('copy');
        });
*/
/*		function copyClip(text) {
		    var input = document.createElement('textarea');
		
		    var k=String.fromCharCode(160);
			 var r="\n"; //String.fromCharCode(13)+String.fromCharCode(13);
			 var rg = new RegExp(k,"g");
			 var metin=text.replace(k,r).replace(rg, '');		//sadece ilkini enter, sonrakiler "".

		    input.value=metin;
		    document.body.appendChild(input);
		    input.select();
		    var result = document.execCommand('copy');
		    //document.body.removeChild(input)

		    return result;
		 }
*/		            
		//resimleri zoomlama        
        d('.zoomla').hover(function () {
            d(this).addClass('transition');
            d("#content").css("overflow", "visible");
        	}, function () {
            d(this).removeClass('transition');
            d("#content").css("overflow", "hidden");
        });            
        
        //content içindeki linklere tıklandığında boş sayfada açsın. Bunu en alta koydum çünkü, wrapperlardan sonra çalışması lazım
    d("#yazilar a").click(function () {
        //not-a-a-tag olanları muaf tutmamız lazımi bunlar gridviwalardaki sortable başlıklar ve paging elemanları için. 
        var parentEls = d(this).parents()
            .map(function () {
                return this.className;
            })
            .get()
            .join(", ");

        if (parentEls.indexOf("not-a-a-tag")>-1) {
            return;
        }

				var hedef = d(this).attr("href");				
				if (hedef.substring(0,1)=="#")//hedef link aynı sayfada değilse yeni tabda açsın. aynı sayfadaki linklerin hedefi direkt # ile başlıyor
				{
		        	d(this).attr('target','_self');
		        }
		        else
                {                    
		        	d(this).attr('target','_blank');
		        }
            });

		

		    //hover da denedim mouserover/out da denedim olmadı, css3-transiton yaptım
		    //d(".zoomla").hover(function () { giris(); }, function () { cikis(); })
		    
		    //function giris() {
		    //    //alert("girdim");
		    //    var h = d(".zoomla").height() * 1.6;
		    //    var w = d(".zoomla").width() * 1.6;
		    //    var l = d(".zoomla").position().left - 180;
		    //    var t = d(".zoomla").position().top - 20;
		    //    d(".zoomla").animate({ height: h, width: w, left: l, top: t }, "fast");
		    //    d("#content").css("overflow", "visible");
		    //}
		
		    //function cikis() {
		    //    //alert("çıktım");
		    //    var h = d(".zoomla").height() / 1.6;
		    //    var w = d(".zoomla").width() / 1.6;
		    //    var l = d(".zoomla").position().left + 180;
		    //    var t = d(".zoomla").position().top + 20;
		    //    d(".zoomla").animate({ height: h, width: w, left: l, top: t }, "fast");
		    //    d("#content").css("overflow", "hidden");    
		    //}  
});
