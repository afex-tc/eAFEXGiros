<%
  function saca_parametros(parametros,tipo)
	saca_parametros = ""
	p= instr(ucase(parametros),tipo)
	if p>0 then
		q= instr(p,parametros,"|")
		if q>0 then
			saca_parametros = "," & mid(parametros,p+len(tipo),q-p-len(tipo)) & ","
		else
			saca_parametros = "," & mid(parametros,p+len(tipo),len(parametros)) & ","
		end if
	end if
  end function

  function limpia(valor, defecto)
	if valor = "" then
		limpia = defecto
		exit function
	end if
	limpia=mid(valor,2, len(valor)-2)
  end function
%>