Sub enviarMailmasivo()


'Declaro variables que voy a usar para hacer los correos


Dim lista As Worksheet

Set lista = Worksheets("Lista")
Dim appexterna As Object
Dim correo As Object
Dim cuentacorreo As String
Dim msg As String


'que no muestre ventanas auxiliares ni actualice la pantalla hasta que termine

With Application
.EnableEvents = False
.ScreenUpdating = False
End With

'ver si outlook esta abierto y si no abrirlo

On Error Resume Next

Set appexterna = GetObject("", Outlook.Application)

Err.Clear

If appexterna Is Nothing Then
Set appexterna = CreateObject("Outlook.Application")
End If


appexterna.Visible = True


'Elegir la cuenta de correo para enviarlo
For i = 1 To appexterna.Session.Accounts.Count
 MsgBox "La cuenta " & appexterna.Session.Accounts.Item(i) & " es la número " & i
Next i

numerocuenta = InputBox("seleccione el numero de la cuenta con la que quiere enviar el correo", "Elija numero de cuenta")

fila_origen = Cells(2, 1).Row

'añadir direccion, asunto , copia a y cuerpo del mensaje

    hora = Hour(Now)
    
    
    Cells(2, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
        
    
    Dim enviacorreo As Range
    
    For Each enviacorreo In Selection
    
    Set correo = appexterna.CreateItem(0)
    correo.To = lista.Cells(fila_origen, 1).Value
    correo.Subject = lista.Cells(fila_origen, 2).Value
    If hora >= 0 Then
    If hora <= 6 Then
    msg = "Buenas Noches, " & vbCr & vbCr
    End If
    End If
    If hora > 6 Then
    If hora <= 13 Then
    msg = "Buenos Días, " & vbCr & vbCr
    End If
    End If
    If hora > 13 Then
    If hora <= 20 Then
    msg = "Buenas Tardes, " & vbCr & vbCr
    End If
    End If
    If hora > 21 Then
    msg = "Buenas Noches, " & vbCr & vbCr
    End If
    
    msg = msg & "Isidro, gracias por ponernos en contacto." & vbCr & vbCr
    msg = msg & lista.Cells(fila_origen, 4) & ", encantado de saludarle" & vbCr & vbCr
    msg = msg & "KVAR CONSULTORES es la gestora de subvenciones a fondo perdido que ha escogido Minsait y su banco (en el que rellenó hace unos días la solicitud de información de ayudas y subvenciones) para ayudarle a encontrar las mejores ayudas disponibles para su empresa, negocio y proyecto." & vbCr & vbCr
    msg = msg & "Una vez que ya tenemos los datos de contacto en nuestra base de datos a lo largo de los próximos días uno de nuestros consultores especialistas en su Comunidad Autónoma le llamará para conocer más en detalle tu negocio y ver que opciones de ayudas hay abiertas. Si hubiera alguna, procederíamos a realizar la solicitud de inmediato. Si no hay ninguna ayuda que se adaptara a su negocio y necesidades mantendremos sus datos en nuestra base de forma que tan pronto como se abra una nueva ayuda, podamos llamarle y ver si está interesado en que se la tramitemos." & vbCr & vbCr
    msg = msg & "Nuestro consultor le explicará nuestra fórmula de trabajo a éxito." & vbCr & vbCr
    msg = msg & "Un cordial saludo,"
    correo.Body = msg
    Set correo.SendUsingAccount = appexterna.Session.Accounts.Item(numerocuenta)
    correo.Display
    'correo.Send 'esto envia el correo, cuidado
    
    fila_origen = fila_origen + 1
    Next enviacorreo

With Application
.EnableEvents = True
.ScreenUpdating = True
End With



End Sub
