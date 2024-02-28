<!--#INCLUDE virtual="/includes/funciones.asp"-->
<% 
  ' Generador de Elementos de Formularios 
  ' Diseñado por Oscar Pinto S. 
%>
<script language=JavaScript src="script/valfecha.js"></script>
<script language=JavaScript src="script/funciones.js"></script>
<%
  Function combo_saca_parametros(parametros,tipo, defecto)
	combo_saca_parametros = ""
	p= instr(ucase(parametros),tipo)
	if p>0 then
		q= instr(p,parametros,"|")
		if q>0 then
			combo_saca_parametros = "," & mid(parametros,p+len(tipo),q-p-len(tipo)) & ","
		else
			combo_saca_parametros = "," & mid(parametros,p+len(tipo),len(parametros)) & ","
		end if
	end if
	if defecto <> "" and combo_saca_parametros = "" then
		combo_saca_parametros = "," & defecto & ","
	end if
  End Function

  Function combo_saca_elemento(parametros,tipo, defecto)
	aux = combo_saca_parametros(parametros,tipo, defecto)
	if aux <> "" then
		combo_saca_elemento = mid(aux,2,len(aux)-2)
	end if
  End Function
 
  Function creador_de_fila(id, obligatorio, descripcion, texto)
	tmp = ""
	if obligatorio  = "SI" then
		tmp_obliga = "<font color=""red"">*</font>&nbsp;&nbsp;"
	else
		tmp_obliga = "&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	tmp = "<tr id=""" & id & """ ><td valign=top>" & tmp_obliga & "" & descripcion & "</td><td>" & texto & "</td></tr>" &chr(13)
	creador_de_fila = tmp
  End Function
 
  Function initcaps(texto)
    letras = "abcdefghijklmnopqrstuvwxyzñüáéíóú"
    Min = ".el.la.los.las.les.a.e.i.o.ó.u.de.del.y.al.nos."
    Max = ".i.ii.iii.iv.v.vi.vii.viii.ix.x.xi.xi."
    tipo = "simbolo"
    salida = ""
    texto = LCase(texto)
    Inicio = 1
    For a = 1 To Len(texto)
        corta = False
        If InStr(letras, Mid(texto, a, 1)) Then
            'es letra
            If tipo = "simbolo" Then corta = True
            tipo = "letra"
        Else
            'es no letra
            If tipo = "letra" Then corta = True
            tipo = "simbolo"
        End If
        If corta Then
            palabra = Mid(texto, Inicio, a - Inicio)
            'a mayusculas la inicial
            If Len(palabra) > 1 Then palabra = UCase(Left(palabra, 1)) & Mid(palabra, 2)
            'palabras a minusculas
            If InStr(Min, "." & LCase(palabra) & ".") > 0 Then palabra = LCase(palabra)
            'palabras a mayusculas
            If InStr(Max, "." & LCase(palabra) & ".") > 0 Then palabra = UCase(palabra)
            'abreviaciones de una letra a mayuscula
            If Len(palabra) = 1 Then If Mid(Mid(texto, Inicio, 2), 2, 1) = "." Then palabra = UCase(palabra)
            salida = salida & palabra
            Inicio = a
        End If
    Next
    palabra = Mid(texto, Inicio, a - Inicio)
    'a mayusculas la inicial
    If Len(palabra) > 1 Then palabra = UCase(Left(palabra, 1)) & Mid(palabra, 2)
    'palabras a minusculas
    If InStr(Min, "." & LCase(palabra) & ".") > 0 Then palabra = LCase(palabra)
    'palabras a mayusculas
    If InStr(Max, "." & LCase(palabra) & ".") > 0 Then palabra = UCase(palabra)
    'abreviaciones de una letra a mayuscula
    If Len(palabra) = 1 Then If Mid(Mid(texto, Inicio, 2), 2, 1) = "." Then palabra = UCase(palabra)
    salida = salida & palabra
	salida = UCase(Left(salida, 1)) & Mid(salida, 2)
	salida = Replace(salida, "o'H", "O'H")
    initcaps = salida
  End Function

  sub combo_sql(nombre,consulta,parametros) 

	defecto = combo_saca_elemento(parametros,"DEFECTO=", "")
	onchange = combo_saca_elemento(parametros,"ONCHANGE=", "")
	especial2 = " " & combo_saca_elemento(parametros,"ESPECIAL2=", "")
	obligatorio = combo_saca_elemento(parametros,"OBLIGATORIO=","SI") 'defecto es SI
	pdescripcion = combo_saca_elemento(parametros,"DESCRIPCION=", nombre) 'defecto= nombre_formulario
	sincronizador = combo_saca_elemento(parametros,"SINCRONIZADOR=", "NO") 'defecto no
	sincronizado = combo_saca_elemento(parametros,"SINCRONIZADO=", "") ' padres
	creafila = ucase(combo_saca_elemento(parametros,"CREAFILA=", "NO")) 'defecto no
	creacelda = ucase(combo_saca_elemento(parametros,"CREACELDA=", "NO")) 'defecto no
	IDfila = combo_saca_elemento(parametros,"IDFILA=", "") 'defecto no
	
	otros = "obligatorio=" & obligatorio

	salida = ""
	xsinc = ""
	if sincronizador <> "NO" then xsinc =  "sincroniza(this);"
	if sincronizado <> "" then
		sincronizador_cuenta = sincronizador_cuenta +1
		v_nom_sinc = ""
		if txt_codigo = "SI" then
			v_nom_sinc = "_sel"
		end if
		set qry = oConn.Execute(consulta)
		arr =  "<script>"
		arr = arr & "	var zinc" & sincronizador_cuenta & "  = new Array ("
		arr = arr & "'" & sincronizado & "','" &  nombre & v_nom_sinc & "',"
		while not qry.eof
			'arr = arr & "'" & qry(1) & "','" & qry(0) & "',""" & qry(2) & ""","
			arr = arr & "'" & qry(2) & "','" & qry(0) & "',""" & qry(1) & ""","
			qry.movenext
		wend
		arr = mid(arr,1,len(arr)-1)
		arr = arr & " );"
		arr = arr & "</script>"
		salida = salida + arr
	end if
	salida_tmp_1 = ""
	salida_tmp_2 = ""
	if txt_codigo = "SI" then
		set qry__= oConn.Execute(consulta)
		v_max = 0
		while not qry__.eof
			if v_max < len(qry__(0)) then
				v_max = len(qry__(0))
			end if
			qry__.movenext
		wend
		salida_tmp_1 = "<INPUT value=" & defecto & " type=text size=" & (v_max+2) & " maxlength=" & (v_max+2) & " name=" & nombre & " onkeyup=frm." & nombre & "_sel.value=frm." & nombre & ".value; >"
		xsinc = xsinc & "frm." & nombre & ".value=frm." & nombre & "_sel.value;"
		nombre = nombre + "_sel"
	end if
	if onchange <> "" then
		salida = salida + "<select " & otros & " descripcion=" & nombre & " name=" & nombre & " onchange=" & xsinc & onchange & ">"
	else
		salida = salida + "<select " & otros & " descripcion=" & nombre & " name=" & nombre & " onchange=" & xsinc & ">"
	end if
	if especial2 <> "" then
		salida = salida + "<option value="""">" & especial2 & "</option>"
	else
		if pdescripcion <> "" then
			salida = salida + "<option value="""">Seleccione " & pdescripcion & "</option>"
		end if
	end if
	set qry=oConn.Execute(consulta)
	while not qry.eof
		if defecto=cstr(qry(0)) then
			salida = salida + "<option selected value=""" & qry(0) & """>" & qry(1) & "</option>"
		else
			salida = salida + "<option value=""" & qry(0) & """>" & qry(1) & "</option>"
		end if
		qry.movenext
	wend
	salida = salida + "</select>"
	salida = salida_tmp_1 & salida & salida_tmp_2
	if creafila = "SI" then salida = creador_de_fila(IDfila, obligatorio, pdescripcion, salida)
	if creacelda = "SI" then salida = creador_de_celda(IDfila, obligatorio, pdescripcion, salida)
	response.Write(salida)
  end sub

  sub campo_numero(nombre_formulario, parametros)
	especial2 = " " & combo_saca_elemento(parametros,"ESPECIAL2=", "")
	pobligatorio = combo_saca_elemento(parametros,"OBLIGATORIO=","SI") 'defecto es SI
	pdescripcion = combo_saca_elemento(parametros,"DESCRIPCION=", nombre_formulario) 'defecto= nombre_formulario
	defecto = combo_saca_elemento(parametros,"DEFECTO=", "")
	creafila = ucase(combo_saca_elemento(parametros,"CREAFILA=", "NO")) 'defecto no
	creacelda = ucase(combo_saca_elemento(parametros,"CREACELDA=", "NO")) 'defecto no
	IDfila = combo_saca_elemento(parametros,"IDFILA=", "") 'defecto no
	par_nowrite = ucase(limpia(saca_parametros(parametros,"WRITE="),"SI"))

	largo = combo_saca_elemento(parametros,"LARGO=", "10")
	maximo = combo_saca_elemento(parametros,"MAXIMO=", largo)
	decimales = combo_saca_elemento(parametros,"DECIMALES=", "")

	onkeyup = combo_saca_elemento(parametros,"ONKEYUP=", "")
	onblur = combo_saca_elemento(parametros,"ONBLUR=", "")
	onchange = combo_saca_elemento(parametros,"ONCHANGE=", "")
	onfocus = combo_saca_elemento(parametros,"ONFOCUS=", "")

	obligatorio = " obligatorio=" & pobligatorio
	descripcion = " descripcion=""" & pdescripcion & """"
	decimales = " decimales=""" & decimales & """"

	salida = ""
	salida = salida + "<input" & obligatorio & descripcion & decimales & " size=" & largo & " maxlength=" & maximo & " type=text name=""" & nombre_formulario & """"
	salida = salida + " value=""" & defecto & """ "
	salida = salida + " onkeydown=""valnum(this);"""
	salida = salida + " onkeyup=""valnum(this);" & onkeyup & """"
	salida = salida + " onBlur=""" & onblur & """"
	salida = salida + " onChange=""" & onchange & """"
	salida = salida + " onFocus=""" & onfocus & """"
	salida = salida + " onkeypress=""valnum(this);"" "&especial2&" >"

	if creafila = "SI" then salida = creador_de_fila(IDfila, pobligatorio, pdescripcion, salida)
	if creacelda = "SI" then salida = creador_de_celda(IDfila, pobligatorio, pdescripcion, salida)
	if par_nowrite="SI" then response.Write(salida) else formulario_salida = salida
  end sub

  sub campo_texto(nombre_formulario, parametros)
	especial2 = " " & combo_saca_elemento(parametros,"ESPECIAL2=", "")
	pobligatorio = combo_saca_elemento(parametros,"OBLIGATORIO=","SI") 'defecto es SI
	pvalidacion = combo_saca_elemento(parametros,"VALIDACION=","SI") 'defecto es SI
	pdescripcion = combo_saca_elemento(parametros,"DESCRIPCION=", nombre_formulario) 'defecto= nombre_formulario
	defecto = combo_saca_elemento(parametros,"DEFECTO=", "")
	password = ucase(combo_saca_elemento(parametros,"PASSWORD=", "NO")) 'defecto no
	creafila = ucase(combo_saca_elemento(parametros,"CREAFILA=", "NO")) 'defecto no
	creacelda = ucase(combo_saca_elemento(parametros,"CREACELDA=", "NO")) 'defecto no
	IDfila = combo_saca_elemento(parametros,"IDFILA=", "") 'defecto no
	par_nowrite = ucase(limpia(saca_parametros(parametros,"WRITE="),"SI"))

	onblur = combo_saca_elemento(parametros,"ONBLUR=", "")
	onchange = combo_saca_elemento(parametros,"ONCHANGE=", "")
	onfocus = combo_saca_elemento(parametros,"ONFOCUS=", "")

	largo = combo_saca_elemento(parametros,"LARGO=", "10")
	maximo = combo_saca_elemento(parametros,"MAXIMO=", largo)
	alto = combo_saca_elemento(parametros,"ALTO=", "")

	obligatorio = " obligatorio=" & pobligatorio
	descripcion = " descripcion=""" & pdescripcion & """"

	salida = ""
	tipo_input = "text"
	if password="SI" then tipo_input = "password"
	if alto<>"" then
		salida = salida + "<textarea " & obligatorio & descripcion & " cols=" & largo & " rows=""" & alto & """ name=""" & nombre_formulario & """"
		if pvalidacion="SI" then
			salida = salida + " onkeydown=""valchar(this);"""
			salida = salida + " onkeyup=""valchar(this);"""
		end if
		salida = salida + " onBlur=""" & onblur & """"
		salida = salida + " onChange=""" & onchange & """"
		salida = salida + " onFocus=""" & onfocus & """"
		if pvalidacion="SI" then
			salida = salida + " onkeypress=""valchar(this);"" "
		end if
		salida = salida & especial2 & " "
	else
		salida = salida + "<input" & obligatorio & descripcion & " size=" & largo & " maxlength=" & maximo & " type=" & tipo_input & " name=""" & nombre_formulario & """"
		if pvalidacion="SI" then
			salida = salida + " onkeydown=""valchar(this);"""
			salida = salida + " onkeyup=""valchar(this);"""
		end if
		salida = salida + " onBlur=""" & onblur & """"
		salida = salida + " onChange=""" & onchange & """"
		salida = salida + " onFocus=""" & onfocus & """"
		if pvalidacion="SI" then
		salida = salida + " onkeypress=""valchar(this);"" "
		end if
		salida = salida & especial2 & " "
	end if
	if alto<>"" then
		salida = salida + " >" & defecto & "</textarea>"
	else
		salida = salida + " value=""" & defecto & """ >"
	end if

	if creafila = "SI" then salida = creador_de_fila(IDfila, pobligatorio, pdescripcion, salida)
	if creacelda = "SI" then salida = creador_de_celda(IDfila, pobligatorio, pdescripcion, salida)
	if par_nowrite="SI" then response.Write(salida) else formulario_salida = salida
   end sub
   
   private sub boton(fun1, fun2)
ffun=fun1
ffun2=fun2
if instr(fun1,"|")<>0 then ffun=left(fun1,instr(fun1,"|")-1)
name = "" & combo_saca_elemento(fun1,"NAME=", "")

if true then
select case ucase(ffun)
	case "BUSCAR" : boton1="imagenes/botones/buscar.gif":boton2="imagenes/botones/buscar_over.gif"
end select
end if
mostrar=true
	if name="" then name = "BTN_" & ucase(ffun)
	if ffun2="" and ucase(fun1)="CERRAR" then ffun2="window.close();"
	if ucase(fun1)="CERRAR" and pventana_<>"1" then mostrar=false
	if ffun2="" then ffun2 = ffun&"()"
	if ucase(ffun) = "PRINT" then
	%>
		<Object ID='WebBrowser1' Width='0' Height='0' ClassID='CLSID:8856F961-340A-11D0-A96B-00C04FD705A2' VIEWASTEXT></Object>
		<script language='vbscript'>
		set doc=document.all 
		Function imprime_print() 
		WebBrowser1.ExecWB 6, 2 
		End Function 
		</script>
		<SCRIPT LANGUAGE=javascript>
		<!--//n
		function cambia_print(cual) 
		{	var coll = document.all.item("para_print");
		    if (coll!=null) 
			{	for(i=0;i<coll.length; i++) 
					coll.item(i).style.display=cual; 
				} 
			} 
		//--> 
		</SCRIPT>
		<%
		ffun2="cambia_print('');imprime_print();cambia_print('none');"
	end if
	if mostrar then
		Response.Write(chr(13) & "<img name="""&name&""" id="""&name&""" style=""cursor:hand;position=relative;"" onclick="""&ffun2&";"" src="""&boton1&""" onmouseover=""this.src='"&boton2&"'"" onmouseout=""this.src='"&boton1&"'"" WIDTH=""95"" HEIGHT=""22"" class=noimprimible >&nbsp;")
	end if
end sub

sub campo_mail(nombre_formulario, parametros)
	especial2 = " " & combo_saca_elemento(parametros,"ESPECIAL2=", "")
	pobligatorio = combo_saca_elemento(parametros,"OBLIGATORIO=","SI") 'defecto es SI
	pdescripcion = combo_saca_elemento(parametros,"DESCRIPCION=", nombre_formulario) 'defecto= nombre_formulario
	defecto = combo_saca_elemento(parametros,"DEFECTO=", "")
	creafila = ucase(combo_saca_elemento(parametros,"CREAFILA=", "NO")) 'defecto no
	creacelda = ucase(combo_saca_elemento(parametros,"CREACELDA=", "NO")) 'defecto no
	IDfila = combo_saca_elemento(parametros,"IDFILA=", "") 'defecto no
	par_nowrite = ucase(limpia(saca_parametros(parametros,"WRITE="),"SI"))

	onblur = combo_saca_elemento(parametros,"ONBLUR=", "")
	onchange = combo_saca_elemento(parametros,"ONCHANGE=", "")
	onfocus = combo_saca_elemento(parametros,"ONFOCUS=", "")

	largo = combo_saca_elemento(parametros,"LARGO=", "10")
	maximo = combo_saca_elemento(parametros,"MAXIMO=", largo)

	obligatorio = " obligatorio=" & pobligatorio
	descripcion = " descripcion=""" & pdescripcion & """"

	salida = ""
	salida = salida + "<input" & obligatorio & descripcion & " size=" & largo & " maxlength=" & maximo & " type=text name=""" & nombre_formulario & """"
	salida = salida + " value=""" & defecto & """ "
	salida = salida + " onBlur=""valemail(this);"" " & onblur & """"
	salida = salida + " onChange=""" & onchange & """"
	salida = salida + " onFocus=""" & onfocus & """"
	salida = salida + " onkeypress=""" & especial2 & """ >"

	if creafila = "SI" then salida = creador_de_fila(idfila, pobligatorio, pdescripcion, salida)
	if creacelda = "SI" then salida = creador_de_celda(IDfila, pobligatorio, pdescripcion, salida)
	if par_nowrite="SI" then response.Write(salida) else formulario_salida = salida
end sub

sub campo_fecha(nombre_formulario, parametros)
	especial2 = " " & combo_saca_elemento(parametros,"ESPECIAL2=", "")
	pobligatorio = combo_saca_elemento(parametros,"OBLIGATORIO=","SI") 'defecto es SI
	pdescripcion = combo_saca_elemento(parametros,"DESCRIPCION=", nombre_formulario) 'defecto= nombre_formulario
	defecto = combo_saca_elemento(parametros,"DEFECTO=", "")
	creafila = ucase(combo_saca_elemento(parametros,"CREAFILA=", "NO")) 'defecto no
	creacelda = ucase(combo_saca_elemento(parametros,"CREACELDA=", "NO")) 'defecto no
	IDfila = combo_saca_elemento(parametros,"IDFILA=", "") 'defecto no
	par_nowrite = ucase(limpia(saca_parametros(parametros,"WRITE="),"SI"))

	onblur = combo_saca_elemento(parametros,"ONBLUR=", "")
	onchange = combo_saca_elemento(parametros,"ONCHANGE=", "")
	onfocus = combo_saca_elemento(parametros,"ONFOCUS=", "")
	onclick = combo_saca_elemento(parametros,"ONCLICK=", "")
	minimo = combo_saca_elemento(parametros,"MINIMO=", "")
	maximo = combo_saca_elemento(parametros,"MAXIMO=", "")
	mensaje = combo_saca_elemento(parametros,"MENSAJE=", "")
	cambio = combo_saca_elemento(parametros,"CAMBIO=", "")

	obligatorio = " obligatorio=" & pobligatorio
	descripcion = " descripcion=""" & pdescripcion & """"

	salida = ""

	salida = salida + "<table cellpadding=0 cellspacing=0><tr><td valign=top>"

	salida = salida + "<input" & obligatorio & descripcion & " size=13 maxlength=10 type=text name=""" & nombre_formulario & """"
	salida = salida +  " value=""" & defecto & """ "
	salida = salida + " MINIMO=""" & minimo & """ "
	salida = salida + " MAXIMO=""" & maximo & """ "
	salida = salida + " MENSAJE=""" & mensaje & """ "

	if instr(ucase(especial2),"READONLY")=0 then
		salida = salida + " onFocus=""javascript:vDateType='3';" & onfocus & """ "
		salida = salida + " onKeyUp=""DateFormat(this,this.value,event,false,'3')"" "
		salida = salida + " onBlur=""DateFormat(this,this.value,event,true,'3');compara_fechas(this);" & onblur & """"
		salida = salida + " ondblclick=""genera_calendario(this);" & cambio & """"
	end if
	salida = salida + " onChange=""" & onchange & """"
	salida = salida + especial2 &" >"
	if instr(ucase(especial2),"READONLY")=0 then
		salida = salida + "<img style=""cursor:hand"" name='img_fecha' id='img_fecha' height=18 src=""includes/calendario.gif"" value=""" & defecto & """ onclick=""" & onclick & " genera_calendario(document.all('" & nombre_formulario & "'));compara_fechas(document.all('" & nombre_formulario & "'));" & cambio & """>"
	end if
	salida = salida + "</td></tr></table>"

	if creafila = "SI" then salida = creador_de_fila(IDfila, pobligatorio, pdescripcion, salida)
	if creacelda = "SI" then salida = creador_de_celda(IDfila, pobligatorio, pdescripcion, salida)
	if par_nowrite="SI" then response.Write(salida) else formulario_salida = salida
end sub

sub campo_rut(nombre_formulario, parametros)
	especial2 = " " & combo_saca_elemento(parametros,"ESPECIAL2=", "")
	pobligatorio = combo_saca_elemento(parametros,"OBLIGATORIO=","SI") 'defecto es SI
	pdescripcion = combo_saca_elemento(parametros,"DESCRIPCION=", nombre_formulario) 'defecto= nombre_formulario
	defecto = combo_saca_elemento(parametros,"DEFECTO=", "")
	creafila = ucase(combo_saca_elemento(parametros,"CREAFILA=", "NO")) 'defecto no
	creacelda = ucase(combo_saca_elemento(parametros,"CREACELDA=", "NO")) 'defecto no
	IDfila = combo_saca_elemento(parametros,"IDFILA=", "") 'defecto no
	par_nowrite = ucase(limpia(saca_parametros(parametros,"WRITE="),"SI"))

	onblur = combo_saca_elemento(parametros,"ONBLUR=", "")
	onchange = combo_saca_elemento(parametros,"ONCHANGE=", "")
	onfocus = combo_saca_elemento(parametros,"ONFOCUS=", "")

	obligatorio = " obligatorio=" & pobligatorio
	descripcion = " descripcion=""" & pdescripcion & """"

	salida = ""
	salida = salida + "<input" & obligatorio & descripcion & " size=12 maxlength=12 type=text name=""" & nombre_formulario & """"
	salida = salida + " value=""" & defecto & """ "
	salida = salida + " onBlur=""valrut(this);""" '& onblur & ""
	salida = salida + " onkeydown=""limpiarut_(this);"""
	salida = salida + " onkeyup=""limpiarut_(this);"""
	salida = salida + " onChange=""" & onchange & """"
	salida = salida + " onFocus=""" & onfocus & """"
	salida = salida + " " & especial2 & " >"

	if creafila = "SI" then salida = creador_de_fila(IDfila, pobligatorio, pdescripcion, salida)
	if creacelda = "SI" then salida = creador_de_celda(IDfila, pobligatorio, pdescripcion, salida)
	if par_nowrite="SI" then response.Write(salida) else formulario_salida = salida
end sub
%>