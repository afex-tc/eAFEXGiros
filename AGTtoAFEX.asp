<%@ Language=VBScript %>
<%
	response.Buffer = True
	response.Clear 

	Dim sURL, prmOut
	
	prmOut = Request("prmOut")
	
	If prmOut = "" Then
		'sURL = "//jmiranda/afexmoneyweb_local/afextoagt.asp?sslu=ibero&sslp=123456789&acc=20&prmURL=//jmiranda/afexmoneyweb_local/agttoafex.asp"
		'sURL = "afextoagt.asp"				
		's = Server.execute(sURL)
		'response.Write s
		
		sURL = "http://192.168.111.13/afexmoneyweb/afextoagt.asp?sslu=lax&sslp=41226923&acc=20"
		Set xml = Server.CreateObject("Microsoft.XMLHTTP")		
	    xml.open "Get", sURL, false
		xml.Send
		prmOut = xml.responseText
		response.Write "oki<br>" & prmOut

		'sURL = "http://192.168.111.13/afexmoneyweb/afextoagt.asp?sslu=ibero&sslp=123456789&acc=20&prmURL=http://192.168.111.13/afexmoneyweb/agttoafex.asp"
		'response.Redirect(sUrl)
		
		response.End 
	End If
	
	response.Write "Esta es la lista:<br>" & prmOut
	response.End 
	
%>
