'Nicolás C. Temporizar Apagado
Dim eleccion, modo, hora, fecha, momento
eleccion = InputBox("Para que el sistema se apague hoy introduzca H" + vbNewLine + vbNewLine + "Si quiere indicar una fecha posterior introduzca F" + vbNewLine + vbNewLine + "Si quiere cancelar el apagado programado pulse Cancelar")
Set oshell = CreateObject("Wscript.Shell")
If eleccion = "H" Then
	'Se va a indicar una hora en el mismo dia
	modo = 1		
	hora = InputBox("Introduce una hora para apagar el sistema. El formato es el siguiente:" + vbNewLine + vbNewLine + "1:2" + vbNewLine + vbNewLine + "El ejemplo indica las una y dos minutos de la mañana. Para las una y dos minutos de la tarde, sería: 13:2")	
	validar(hora)
Else
	If eleccion = "F" Then	
		'Se va a indicar una fecha con hora
		modo = 2		
		fecha = InputBox("Introduce la fecha de apagado. El formato es el siguiente: " + vbNewLine + vbNewLine + "10/1/2013" + vbNewLine + vbNewLine + "El ejemplo indica el día 10 de enero de 2013")
		validar(fecha)
		momento = InputBox("Introduce una hora para apagar el sistema. El formato es el siguiente:" + vbNewLine + vbNewLine + "1:2" + vbNewLine + vbNewLine + "El ejemplo indica las una y dos minutos de la mañana. Para las una y dos minutos de la tarde, sería: 13:2")
		validar(momento)
	Else
		'Si no se ha introducido ni F ni H, el script se desactiva
		MsgBox "Apagado Automático cancelado", vbInformation
		WScript.Quit
	End If	
End If
'Mensaje de confirmacion que se auto-cierra tras 4 segundos
oshell.popup "Ok", 4, "Ok"
'Bucle de comprobacion
While true	
	If modo = 1 Then
		If (Hour(Now()) & ":" & Minute(Now)) = hora Then
			oshell.Exec("shutdown -s -t 00")							
		End If
	Else
		If ((Hour(Now()) & ":" & Minute(Now)) = momento) And ((Day(Date()) & "/" & Month(Date()) & "/" & Year(Date())) = fecha) Then				
			oshell.Exec("shutdown -s -t 00")				
		End If			
	End If
	'Verificar cada 6 segundos
	Wscript.Sleep(6000)		
Wend

'Verifica si se ha introducido un valor nulo (Cancelar). Se puede extender a comprobar que las horas y fechas son posibles...
Sub validar(variablePedida)
	If variablePedida = "" Then
		MsgBox "Apagado Automático cancelado", vbInformation
		WScript.Quit
	End If
End Sub
