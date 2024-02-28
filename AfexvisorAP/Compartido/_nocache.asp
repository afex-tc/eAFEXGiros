<!-- _nocache.asp -->
<%
	Response.Expires = 0
	Response.buffer = true	
	Response.expires = 0
	Response.expiresabsolute = Now() - 1
	Response.addHeader "pragma", "no-cache"
	Response.addHeader "cache-control", "private"
	Response.CacheControl = "no-cache"	
%>