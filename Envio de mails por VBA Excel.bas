Attribute VB_Name = "Module5"
'CoDe bY Skull&BoNes 2023
Sub EnviaMailV3()

'Crea OBJ dicccionario para almacenar los datos con vencimiento en 30 dias o menos recolectados de los Mails
Dim datos As Object
Set datos = CreateObject("Scripting.Dictionary")
'Crea OBJ dicccionario para almacenar los datos con fecha ya vencida recolectados de los Mails
Dim datos2 As Object
Set datos2 = CreateObject("Scripting.Dictionary")

'lista de CC envio de mails
Dim mailsCC As String
mailsCC = Range("T4").value


'Datos a guardar en el diccionario
Dim correo As String
Dim nombre As String
Dim dni As String
Dim serie As String
Dim fecha As String

'Var para validar que no se mande varias veces el mismo mail
Dim fechaActual As Date
fechaActual = Range("T3").value

'Declara vars para las fechas
Dim fechaVencimiento As Date
Dim diasRestantes As Integer

'Valida que no se haya mandando el mismo dia el mismo mail

If fechaActual <> Date Then
    Range("T3").value = Date

'Loop principal de recorrida de celdas
Dim i As Integer
i = 2
Do While Not IsEmpty(Range("Q" & i))
    'Recorre las fechas
    fechaVencimiento = Range("Q" & i).value
    diasRestantes = DateDiff("d", Date, fechaVencimiento) 'Calcula la diferencia de fechas con el dia de hoy
        
        'Filtra y recolecta en diccionario los dias proximos a vencer en este caso cuando faltan 30 dias(en Nro. entero determina el vencimiento)
        If diasRestantes <= 30 And diasRestantes > 0 Then
            correo = Range("G" & i).value
            dni = Range("B" & i).value
            serie = Range("D" & i).value
            nombre = Range("E" & i).value
            fecha = Range("Q" & i).value
         
               'ingresa nuevo correo en el caso que no exista en el diccionario
                If Not datos.Exists(correo) Then
                        datos.Add correo, Array(Array(dni, serie, nombre, fecha))
                
                    'ingresa un array nuevo dentro del array principal en el caso que exista el mail y agrupa los datos en una sola clave
                    Else
                          Dim value As Variant
                          value = datos(correo)
                          ReDim Preserve value(UBound(value) + 1)
                          value(UBound(value)) = Array(dni, serie, nombre, fecha)
                          datos(correo) = value
                 End If
        
        'Filtra y recolecta en dicionario los datos cuando ya vencio el registro
        ElseIf diasRestantes <= 0 Then
                 correo = Range("G" & i).value
                 dni = Range("B" & i).value
                 serie = Range("D" & i).value
                 nombre = Range("E" & i).value
                 fecha = Range("Q" & i).value
                 
                 If Not datos2.Exists(correo) Then
                          datos2.Add correo, Array(Array(dni, serie, nombre, fecha))
                
                    'ingresa un array nuevo dentro del array principal en el caso que exista el mail y agrupa los datos en una sola clave
                    Else
                          Dim value2 As Variant
                          value2 = datos2(correo)
                          ReDim Preserve value2(UBound(value2) + 1)
                          value2(UBound(value2)) = Array(dni, serie, nombre, fecha)
                          datos2(correo) = value2
                    End If
        End If
i = i + 1
Loop

'Valida si se copiaron datos al diccionario en el caso que falten 30 dias o menos para el vencimiento
If datos.Count = 0 Then

    MsgBox "No hay registros que venzan en los próximos 30 días", vbExclamation, "Sin datos"

Else
    
    'Variable donde se van a guardar los datos a mandar por Email
    Dim resultado As String

    'recorre el diccionario
    For Each Key In datos.Keys()
       
        'junta los datos para en body de mail
        For Each Item In datos(Key)
                   resultado = resultado & "<br>" & _
                   "DNI: " & Item(0) & ", Serie: " & Item(1) & ", Nombre: " & Item(2) & ", Fecha de vencimiento registro: " & Item(3)
                   Next Item
        
                'Inicializa el OBJ Outlook
                Set OutApp = CreateObject("outlook.Application").CreateItem(0)
            
                With OutApp
                                    .To = Key
                                    .Subject = "Recordatorio de Vencimiento registros PMHV"
                                    .CC = mailsCC
                                     End With
                                                         
                                    OutApp.HTMLBody = "<h2 style =color:yellow;>Registros PMHV a vencer en los próximos 30 días</h2> " & "<h3> <h3> " & resultado
                                    
                                    '*** Comentar siguiente linea para no enviar y hacer test ***
                                    'OutApp.Send
                                    
                                    '*** Descomentar siguiente linea pra hacer test de lo que se envia ***
                                    OutApp.Display
        
                '*** Descomentar siguiente linea para capturar datos para test ***
                'MsgBox resultado
                
                'resetea el body del mail para enviar otro nuevo
                resultado = ""
                'Espera 1 segundo antes de enviar el siguiente Mail
                Application.Wait (Now + TimeValue("0:00:01"))
        Next Key
        
        
End If
    

'**************************** CASO DONDE LA FECHA YA ESTA VENCIDA *******************************
'Valida si se copiaron datos al diccionario en el caso de registros vencidos
If datos2.Count = 0 Then

    MsgBox "No hay registros vencidos", vbExclamation, "Sin datos"

Else
    
    'Variable donde se van a guardar los datos a mandar por Email
    Dim resultado2 As String

    'recorre el diccionario
    For Each Key In datos2.Keys()
       
        'junta los datos para en body de mail
        For Each Item In datos2(Key)
                   resultado2 = resultado2 & "<br>" & _
                   "DNI: " & Item(0) & ", Serie: " & Item(1) & ", Nombre: " & Item(2) & ", Fecha de vencimiento registro: " & Item(3)
                   Next Item
        
                'Inicializa el OBJ Outlook
                Set OutApp2 = CreateObject("outlook.Application").CreateItem(0)
            
                With OutApp2
                                    .To = Key
                                    .Subject = "Recordatorio de registros PMHV vencidos"
                                    .CC = mailsCC
                                     End With
                                                         
                                    OutApp2.HTMLBody = "<h2 style = color:red;>Registros PMHV vencidos</h2> " & "<h3> <h3> " & resultado2
                                    
                                    '*** Comentar siguiente linea para no enviar y hacer test ***
                                    'OutApp.Send
                                    
                                    '*** Descomentar siguiente linea pra hacer test de lo que se envia ***
                                    OutApp2.Display
        
                '*** Descomentar siguiente linea para capturar datos para test ***
                'MsgBox resultado
                
                'resetea el body del mail para enviar otro nuevo
                resultado2 = ""
                'Espera 1 segundo antes de enviar el siguiente Mail
                Application.Wait (Now + TimeValue("0:00:01"))
        Next Key
End If
        
'Confirma que todo salio bien
MsgBox "Envio exitoso de Mails!", vbInformation, "Éxito"
Else
    MsgBox "Mails ya enviados hoy, NO se pueden enviar de nuevo!", vbExclamation, "No se puede enviar!"
End If

End Sub



