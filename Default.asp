<%@ Language=VBScript %>
<%

	Session("StringConexion") = "DSN=AFEXchange;UID=cambios;PWD=cambios;"
	
urldelavisita = request.servervariables("remote_addr")
'response.write urldelavisita
'response.end

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>AFEX Ltda.</TITLE>
</HEAD>



<%If urldelavisita <> "66.150.161.133" Then%>
	<frameset ROWS="5, *" BORDER="0">	
		<frame Name="Cabecera" SRC="top.html" scrolling="NO">
		<frame Name="Principal" SRC="Home.asp" scrolling="NO">	
	</frameset>

	<NOFRAMES>
		<P>Solo puede entrar si su navegador soporta marcos.</p>
	</NOFRAMES>

<%Else%>
	<FRAMESET COLS="190,*" BORDER=0>
	<FRAME NAME="Menu" SRC="Menu.asp">
	<FRAMESET ROWS="5, *" >
		<FRAME Name="Cabecera"  SRC="Cabecera.asp" scrolling="NO">
		<FRAME Name="Principal" SRC="Principal.asp">
		<NOFRAMES>
			<P>Solo puede entrar si su navegador soporta marcos.</p>
		</NOFRAMES>
	</FRAMESET>
</FRAMESET>

<%End If%>
<!--
<FRAMESET COLS="190,*" BORDER=0>
	<FRAME NAME="Menu" SRC="Menu.asp">
	<FRAMESET ROWS="80, *" >
		<FRAME Name="Cabecera"  SRC="Cabecera.asp" scrolling="NO">
		<FRAME Name="Principal" SRC="Principal.asp">
		<NOFRAMES>
			<P>Solo puede entrar si su navegador soporta marcos.</p>
		</NOFRAMES>
	</FRAMESET>
</FRAMESET>
-->
<body>
<!-- SiteCatalyst code version: H.15.1.
Copyright 1997-2008 Omniture, Inc. More info available at
http://www.omniture.com -->
<script language="JavaScript" type="text/javascript" src="/s_code.js"></script>
<script language="JavaScript" type="text/javascript"><!--
/* You may give each page an identifying name, server, and channel on
the next lines. */
s.pageName="Index Afex"
var pagina=document.referrer
function ServerInfo_URL2SecondLevelDomain(pagina) {
	if(pagina!=''){
		var first_split = pagina.split("//");
		var without_resource = first_split[1];
		var second_split = without_resource.split("/");
		var dominio = second_split[0];
	}
	if(pagina==''){
	     var dominio ='directo'
	}	 
		return dominio;
}
var anterior=ServerInfo_URL2SecondLevelDomain(pagina)
if (anterior!='www.afex.cl'&&anterior!='directo'){
var evento='event1'
var variable4='click sitio'
var variable5=s.pageName
var temprop1='click sitio'
var temprop2=s.pageName
}
if (anterior=='directo'){
var evento='event2'
var variable6='directo sitio'
var variable7=s.pageName
var temprop3='directo sitio'
var temprop4=s.pageName
}
s.server=""
s.channel="Contacto"
s.pageType=""
s.prop1=temprop1
s.prop2=temprop2
s.prop3=temprop3
s.prop4=temprop4
s.prop10=""
s.prop11=""
/* Conversion Variables */
s.campaign=""
s.state=""
s.zip=""
s.events=evento
s.products=""
s.purchaseID=""
s.eVar4=variable4
s.eVar5=variable5
s.eVar6=variable6
s.eVar7=variable7
s.eVar11=""
/************* DO NOT ALTER ANYTHING BELOW THIS LINE ! **************/
var s_code=s.t();if(s_code)document.write(s_code)//--></script>
<script language="JavaScript" type="text/javascript"><!--
if(navigator.appVersion.indexOf('MSIE')>=0)document.write(unescape('%3C')+'\!-'+'-')
//--></script><noscript><a href="http://www.omniture.com" title="Web Analytics"><img
src="http://wallafex.112.2O7.net/b/ss/wallafex/1/H.15.1--NS/0"
height="1" width="1" border="0" alt="" /></a></noscript><!--/DO NOT REMOVE/-->
<!-- End SiteCatalyst code version: H.15.1. -->



<!-- JFMG 08-06-2011 MARCACIÓN WALLABY -->
<script type="text/javascript">

    var _gaq = _gaq || [];
    _gaq.push(['_setAccount', 'UA-21601361-1']);
    _gaq.push(['_trackPageview']);

    (function() {
        var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
        ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
        var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
    })();

</script>
<!-- FIN JFMG 08-06-2011 MARCACIÓN WALLABY -->

</body>
</HTML>
