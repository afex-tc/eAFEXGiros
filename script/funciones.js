// JavaScript Document
function compara_fechas(este)
{	if (este.value=="") return;
	var feste = Number(este.value.substr(6,4))*10000+Number(este.value.substr(3,2))*100+Number(este.value.substr(0,2));
	var fmin = "";
	var fmax = "";
	var min = ""
	var max = ""
	var vmensaje = ""
	if (este.MENSAJE!="") vmensaje='\n\n'+este.MENSAJE;
	var v_error=0;
	if (este.MINIMO!="") if ("0123456789".indexOf(este.MINIMO.substring(0,1))==-1) min=eval("document.all."+este.MINIMO+".value"); else min=este.MINIMO;
	if (este.MAXIMO!="") if ("0123456789".indexOf(este.MAXIMO.substring(0,1))==-1) max=eval("document.all."+este.MAXIMO+".value"); else max=este.MAXIMO;
	if (min!="") fmin=Number(min.substr(6,4))*10000+Number(min.substr(3,2))*100+(min.substr(0,2)*1);
	if (max!="") fmax=Number(max.substr(6,4))*10000+Number(max.substr(3,2))*100+(max.substr(0,2)*1);
	if (min!="") if ((feste<fmin)) v_error=2
	if (max!="") if ((feste>fmax)) v_error=3
	if (v_error==2)	alert('La Fecha tiene que ser posterior o igual a '+min+vmensaje)
	if (v_error==3)	alert('La Fecha tiene que ser anterior o igual a '+max+vmensaje)
	if (v_error!=0) {este.value="";este.focus();}
	return
}
function sincroniza(este){
	    var inicio = 1;
		while (eval('typeof(zinc' + inicio + ') != "undefined"' )){
			arreglo = eval("zinc" + inicio );
			if (arreglo[0].toUpperCase() == este.name.toUpperCase()){
				destino = eval('document.frm.' +arreglo[1]);
				destino.length = 0;
				// Agregamos el elemento -- Seleccione --
					var newOpt = document.createElement('option');
					newOpt.value = '';
					newOpt.text = '-- Seleccione '+destino.nombre+' --';
					//destino.add(newOpt);
				// Regeneramos los demás elementos
				for (i=2; i < arreglo.length ; i+=3){
					if (este.selectedIndex>=0){
						if (arreglo[i] == este.options[este.selectedIndex].value){
							var newOpt = document.createElement('option');
							newOpt.value = arreglo[i+1];
							newOpt.text = arreglo[i+2];
							destino.add(newOpt);
						}
					}
				}
				if (destino.defecto != null) destino.value = destino.defecto;
				if (destino.selectedIndex == -1) destino.selectedIndex = 0;
				if (destino.onchange != null) destino.onchange();
			}
			inicio++;
		}
}
	
// Funciones para Validar Datos
function no_ocultar(){
		for (var i = 0; i < (document.frm.elements).length; i++)
			if ((document.frm.elements[i].type=="text"||document.frm.elements[i].type=="select-one"||document.frm.elements[i].type=="textarea"))
				document.frm.elements[i].disabled=false;
		for (var i = 0; i < (document.images).length; i++) if (document.images[i].name=="img_fecha") document.images[i].style.display="";
}
function no_ocultar_1(){
		for (var i = 0; i < (document.frm.elements).length; i++)
			if ((document.frm.elements[i].type=="text"||document.frm.elements[i].type=="select-one"||document.frm.elements[i].type=="textarea"))
				document.frm.elements[i].disabled=false;
}
function ocultar(){
		for (var i = 0; i < (document.frm.elements).length; i++)
			if ((document.frm.elements[i].type=="text"||document.frm.elements[i].type=="select-one"||document.frm.elements[i].type=="textarea"))
				document.frm.elements[i].disabled=true;
		for (var i = 0; i < (document.images).length; i++) if (document.images[i].name=="img_fecha") document.images[i].style.display="none";
}
function valida(){
		for (var i = 0; i < (document.frm.elements).length; i++)
			if (document.frm.elements[i].type=="text"||document.frm.elements[i].type=="select-one"||document.frm.elements[i].type=="textarea")
				if ((document.frm.elements[i].value).length==0&&document.frm.elements[i].obligatorio=="SI"){alert("El campo "+document.frm.elements[i].descripcion+" es Obligatorio");document.frm.elements[i].focus();return false;}
		return true;
}

function genera_calendario(este){
		var x = window.showModalDialog('includes/calendario.htm', este.value,'dialogHeight:230px;dialogWidth:200px;status:no;help:no;scroll=no');
		este.value = x;
}

function valnumx(objeto){
		if (isNaN(objeto.value)){
				objeto.value=(objeto.value).substr(0,(objeto.value).length -1)
		}
}

function valemail(objeto) {
		if (objeto.value!=""){
		if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(objeto.value)){
			return (true)
		} else {
			alert("La dirección de email es incorrecta.");
			objeto.value="";
			objeto.focus();
			return (false);
		}}
}

function valnum(objeto){
		var texto = objeto.value;
		var decimales = objeto.decimales;
		vDigitosAceptados = '0123456789';
		if ((decimales==null) || (decimales==0)) {
			vDigitosAceptados = '0123456789';
		}
		else{
			vDigitosAceptados = '0123456789.,';
		}
		var salida='';
		var pos=0;
		for (var x=0; x < texto.length; x++){
			pos = vDigitosAceptados.indexOf(texto.substr(x,1));
			if (pos != -1) salida += texto.substr(x,1);
			if (pos == 10) vDigitosAceptados = '0123456789';
		}
		if (decimales!=null) {
			punto = salida.indexOf('.');
			if (punto != -1) punto = salida.indexOf(',');
			if (punto != -1){
				salida = salida.substr(0,punto + Number(decimales)+1);
			}
		}
		if (objeto.value != salida) objeto.value = salida;
		objeto.value=(objeto.value).replace(",",".")
}

function valchar(objeto)
	{
		vDigitosAceptados = ' abcdefghijklnmñopqrstuvwxyzABCDEFGHIJKLNMÑOPQRSTUVWXYZúéíóáÁÉÍÓÚÀÈÌÒÙàèìòù,:.;-_!·$%&/()=?¿çºª0123456789¡#';
		var texto = objeto.value;
		var salida='';
		for (var x=0; x < texto.length; x++){
			pos = vDigitosAceptados.indexOf(texto.substr(x,1));
			if (pos != -1) salida += texto.substr(x,1);
		}
		if (objeto.value != salida) objeto.value = salida;
}
	
	// Otras Funciones
function valrut(objeto)
	{ if (objeto.value=="") return;
	  objeto.value=objeto.value.replace("k","K");
	  var respuesta=true;
	  if (objeto.value.indexOf("-")==-1)
	  {   v_rut = objeto.value;
		  if (Number(objeto.value.substr(0,objeto.value.length-1)) < 3000000)
		  {
			objeto.value=""
		  }
		  else
		  {		
			objeto.value=objeto.value.substr(0,objeto.value.length-1)+'-'+objeto.value.substr(objeto.value.length-1,1)
		  //alert(objeto.value);
			valrut_(objeto);
		  }
		  //alert(objeto.value);
		  if (objeto.value=="")
		  {	objeto.value = v_rut;
			//alert(objeto.value);
			coloca_verificador(objeto);
			valrut_(objeto);
		    //alert(objeto.value);
			}
		  if (objeto.value==""){objeto.value = v_rut;}
		}
	  valrut_(objeto);
	  //go.go();
	  if (objeto.value==""){v_rut = objeto.value;alert("Rut Invalido");objeto.value="";objeto.focus();return false;}
}
	 
function valrut_(objeto)
	{ var rut=objeto.value;suma=0;mul=2;i=0;
	  respuesta=true;
	  if (rut=="") return false;
	  for (i=rut.length-3;i>=0;i--){
	    suma=suma+Number(rut.charAt(i)) * mul;
	    mul= mul==7 ? 2 : mul+1;
	  }
	  var dvr = ''+(11 - suma % 11);
	  if (dvr=='10') dvr = 'K'; else if (dvr=='11') dvr = '0';
	if (rut.charAt(rut.length-2)!="-"||rut.charAt(rut.length-1)!=dvr) objeto.value="";
}
	
function limpiarut_(objeto)
	{  
	   objeto.value=objeto.value.replace("-.","-K");
	   objeto.value=objeto.value.replace("k","K");
		vDigitosAceptados = '0123456789-K';
		if (objeto.value.substr(0,1) == "0")
		{	objeto.value = objeto.value.substr(1,objeto.value.length);
			}
		var texto = objeto.value;
		var salida='';
		for (var x=0; x < texto.length; x++){
			pos = vDigitosAceptados.indexOf(texto.substr(x,1));
			if (pos != -1) salida += texto.substr(x,1);
		}
	if (objeto.value != salida) objeto.value = salida;
}
	
function coloca_verificador(objeto)
	{ var rut=objeto.value;suma=0;mul=2;i=0;
	//alert(rut.length);
	  if ((rut.length>7)&&(Number(rut.substr(0,2))>22)) return false;
	  if (rut=="") return false;
	  for (i=rut.length-1;i>=0;i--){
	    suma=suma+Number(rut.charAt(i)) * mul;
	    mul= mul==7 ? 2 : mul+1;
	  }
	  var dvr = ''+(11 - suma % 11);
	  if (dvr=='10') dvr = 'K'; else if (dvr=='11') dvr = '0';
	  objeto.value=objeto.value+'-'+dvr;
}


