<!-- MenuDetalleCliente.asp -->
<!--#INCLUDE virtual="/Compartido/Rutinas.htm" -->

<table class="BordeSombra" cellpadding="0" cellspacing="0" border="0" width="175" bgcolor="#f1f1f1" style="background-color: white; color: ">
<tr height="22" bgcolor="#ccddee" style="font-size: 10pt; font-weight: bold">
	<td colspan="5">&nbsp;&nbsp;Configuración Lista</td>
</tr>
<tr height="1"><td colspan="5" bgcolor="#D6D3D6"></td></tr>
<tr>
	<td width="10"></td>
	<td>
		<table cellpadding="0" cellspacing="4" border="0" style="color: gray">
			<tr height="5px"><td></td></tr>					
			<tr>
				<td>Sucursal<br>
					<select id="cbxSucursal" style="width: 300px" <%=sHabilitado%>>
						<% CargarSucursal sSucursal %>
					</select>
				</td>
			</tr>
			<tr>
				<td>Cliente<br>
					<input name="txtCliente" id="txtCliente" style="width: 300px" OnKeyPress="IngresarTexto(3)" maxlength="60" value="<%=sCliente%>">
					<br>
				</td>				
			</tr>
			<tr>
				<td>
					<form id="frmcli" method="post">
						<Input type="radio" name="optRut" style="width: 90px" >Rut
						<br>
						<Input type="radio" name="optApel" style="width: 90px" checked>Nombre o Apellido
						<!-- Jonathan Miranda G. 20-03-2007-->
						<br>
						<Input type="radio" name="optPO" style="width: 90px">Perfil Operacional
						<!--              Fin              -->
						<table id="tblPerfil" style="display: none; font-size: 10px">
							<tr>
								<td></td>
							</tr>
							<tr>
								<td>Nivel de Riesgo<br>
									<select name="cbxRiesgo" style="font-size: 10px">	
									<%	CargarEstado "NIVELRIESGO", iRiesgo %>
										<option value="0"></option>
									</select>
								</td>
								<td>PEP<br>
									<select name="cbxPerfilPEP" style="font-size: 10px">
									<%	CargarPerfil "1", iPerfilPEP %>
									</select>
								</td>
							</tr>
                            <tr>
								<td>Perfil del Cliente<br>
									<select name="cbxPerfilCliente" style="font-size: 10px">
									<%	CargarPerfil "6", iPerfilZona %>
									</select>
								</td>
							</tr>
							<tr>
								<td>Zona<br>
									<select name="cbxPerfilZona" style="font-size: 10px">
									<%	CargarPerfil "2", iPerfilZona %>
									</select>
								</td>
							</tr>
							<tr>
								<td>Residencia<br>
									<select name="cbxPerfilRS" style="font-size: 10px">
									<%	CargarPerfil "3", iPerfilRS %>
									</select>
								</td>
								<td style="display: none">Actividad<br>
									<select name="cbxPerfilACT" style="font-size: 10px">
									<%	CargarPerfil "4", iPerfilACT %>
									</select>
								</td>
							</tr>
							<tr>
								<td>Industria MBS<br>
									<select name="cbxPerfilIndustria" style="font-size: 10px">
									<%	CargarPerfil "5", iPerfilIndustria %>
									</select>
								</td>
							</tr>
						</table>
						<br>
						<Input type="radio" name="optDH" style="width: 90px">Deshabilitados
					</form>
				</td><br>
			</tr>
			<tr><td></td></tr>
			<tr>
				<td align="center"><img src="http:../images/BotonAceptar.jpg" id="cmdAceptar" style="cursor: hand"></td>
			</tr>
			<tr height="10px"><td></td></tr>
		</table>
	</td>	
	<td width="20"></td>
</tr>
</table>

