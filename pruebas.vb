Imports System.IO
Imports System.Text.RegularExpressions
Imports Outlook = Microsoft.Office.Interop.Outlook


Module pruebas


    Public Function ModifyCalendar(invitados As List(Of String)) As Boolean

        Dim oApp As Outlook.Application = New Outlook.Application()
        ' Get the NameSpace and Logon information.
        ' Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");

        Dim oNS As Outlook.NameSpace = oApp.GetNamespace("mapi")

        'Log on by using a dialog box to choose the profile
        'oNS.Logon(Reflection.Missing.Value, Reflection.Missing.Value, True, True)

        'Alternate logon method that uses a specific profile.        
        'oNS.Logon("jhonatan.bazalduao@uanl.mx", "contraseña", False, True)

        'Get the Calendar folder.
        Dim oCalendar As Outlook.MAPIFolder = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)

        'Get the Items (Appointments) collection from the Calendar folder.
        Dim oItems As Outlook.Items = oCalendar.Items

        'Get the first item, last or specific.  whit the function Find, you can find a specific item         
        Dim oAppt As Outlook.AppointmentItem = oItems.Find("[Subject]='Revision Directiva Prueba 2'")

        '"[Start] >= '15/03/2014 01:00 p. m.' AND [End] <= '15/03/2014 01:30 p. m.'"
        '"[Subject]='Revision Directiva Prueba 2'"                                       this is another example of serching data

        'Dim filtro As String = "[Start] >= '" + start.ToString("g") + "' AND [End] <= '" + endT.ToString("g") + "'"        

        'Show some common properties to demonstrate that it's the correct item
        Console.WriteLine("Subject: " + oAppt.Subject)
        Console.WriteLine("Organizer: " + oAppt.Organizer)
        Console.WriteLine("Start: " + oAppt.Start.ToString())
        Console.WriteLine("End: " + oAppt.End.ToString())
        Console.WriteLine("Location: " + oAppt.Location)
        Console.WriteLine("Recurring: " + oAppt.IsRecurring.ToString())
        Console.WriteLine("Body: " + oAppt.Body)
        Console.WriteLine("Conversatiom ID: " + oAppt.ConversationID)
        Console.WriteLine("Duration: " + oAppt.Duration.ToString())
        Console.WriteLine("Companies: ")

        'this is an exemple of how modify some common properties.
        'oAppt.Subject = "cumpleaños"
        'oAppt.Start = DateTime.Now.AddDays(5)
        'oAppt.End = DateTime.Now.AddDays(5)

        'oAppt.Body = "this is an example of how to change a calendar event."        

        'oAppt.Send() ' to send the mail 
        'oAppt.Save() ' to save the changes
        'oAppt.Display(True) 'Show the item 

        'with this you can see all the recipients and if they accept the appoitment or not
        'If Not IsNothing(oAppt) Then
        'Dim recipients As Outlook.Recipients = oAppt.Recipients
        'Dim recipient As Outlook.Recipient = Nothing
        'Dim i As Integer = 0

        'Console.WriteLine(recipients.Count().ToString())

        'For Each recipient In recipients
        'Console.WriteLine(recipient.Name + " :: Meeting Status = " + recipient.MeetingResponseStatus.ToString())
        'Next
        'End If


        ' Add attendees.
        Dim oRecipts As Outlook.Recipients = oAppt.Recipients
        ' Add required attendee.
        Dim oRecipt As Outlook.Recipient

        For Each invitado In invitados
            Console.WriteLine(invitado)
            oRecipt = oRecipts.Add(invitado)
            oRecipt.Type = Outlook.OlMeetingRecipientType.olRequired
            oRecipts.ResolveAll()
        Next
        oAppt.Send()

        'Done. Log off.
        oNS.Logoff()

        'Clean up...... this is not necessary but if is in the web server, maybe for the resources of server
        'oAppt = Nothing
        'oItems = Nothing
        'oCalendar = Nothing
        'oNS = Nothing
        'oApp = Nothing

        Console.Read()
        Return 0
    End Function

    Public Function CreateNewAppoitment()

        'I saw that this is the same code than kaizen 
        Try

            Dim olApp As New Outlook.Application()
            Dim mapiNS As Outlook.NameSpace = olApp.GetNamespace("MAPI")

            Dim profile As String = ""
            mapiNS.Logon(profile, Nothing, Nothing, Nothing)

            Dim apt As Outlook._AppointmentItem = DirectCast(olApp.CreateItem(Outlook.OlItemType.olAppointmentItem), Outlook._AppointmentItem)

            'set some properties
            apt.Subject = "My dog Birthday"
            apt.Body = "it's a special day because my dog came with us a sunny day like this one"
            apt.Start = New DateTime(2014, 3, 8, 13, 30, 0)
            apt.End = New DateTime(2014, 3, 8, 14, 31, 0)
            apt.Importance = Outlook.OlImportance.olImportanceHigh
            apt.ReminderMinutesBeforeStart = 15 ' Number of minutes before the event for the remider
            apt.BusyStatus = Outlook.OlBusyStatus.olBusy '
            apt.AllDayEvent = False
            apt.Location = "My house"

            apt.Send()
            apt.Save()

        Catch ex As Exception
            Console.WriteLine(ex)
            Return 0
        End Try
        Return 1
    End Function


    Public Function getNumeroPaginas() As Integer
        Dim sr As New StreamReader("C:\\Users\\jhonatan.bazalduao\\Documents\\Visual Studio 2010\\Projects\\PdfPages\\PdfPages\\PDF\\Reporte_Enero.pdf")
        Dim pattern As String = "/Type\s*/Page[^s]"
        Dim matches As MatchCollection = Regex.Matches(sr.ReadToEnd, pattern, RegexOptions.IgnorePatternWhitespace)
        Console.WriteLine("Son {0} paginas. ", matches.Count)
        sr.Close()
        Return matches.Count

    End Function


    Public Function SendMail2()
        Try

            '        ' Create an Outlook application.
            Dim oApp As Outlook.Application = New Outlook.Application()


            ' Get Mapi NameSpace and Logon.
            Dim oNS As Outlook.NameSpace = oApp.GetNamespace("mapi")
            'oNS.Logon("amabely.rosalese@uanl.mx", Reflection.Missing.Value, False, True) ' TODO:
            'oNS.Logon("jhonatan.bazalduao@uanl.mx", "contraseña", False, True)

            ' Create an AppointmentItem.
            Dim oAppt As Outlook._AppointmentItem = oApp.CreateItem(Outlook.OlItemType.olAppointmentItem)
            '        'oAppt.Display(true)  'Modal
            '        ' Change AppointmentItem to a Meeting. 
            oAppt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting

            '        ' Set some common properties.
            oAppt.Subject = "Reunion Directiva example 2.5"
            oAppt.Body = "it's a little example of how to send a new appointment :D"
            oAppt.Location = "dgi"

            oAppt.Start = Convert.ToDateTime("20/03/2014 07:00 p. m.")
            oAppt.End = Convert.ToDateTime("20/03/2014 08:00 p. m.")

            oAppt.ReminderSet = True
            oAppt.ReminderMinutesBeforeStart = 25
            oAppt.BusyStatus = Outlook.OlBusyStatus.olBusy  '  olBusy
            oAppt.IsOnlineMeeting = False
            oAppt.AllDayEvent = False

            ' Add attendees.
            Dim oRecipts As Outlook.Recipients = oAppt.Recipients

            ' Add required attendee.
            Dim oRecipt As Outlook.Recipient
            oRecipt = oRecipts.Add("jhona.baz@outlook.com") ' TODO:   
            oRecipt = oRecipts.Add("jhona.22.baz@gmail.com") ' TODO:   
            oRecipt.Type = Outlook.OlMeetingRecipientType.olRequired
            oRecipts.ResolveAll()
            'oAppt.Display(true)
            '        ' Send out request.
            oAppt.Send()
            '        ' Logoff.
            oNS.Logoff()

        Catch ex As Exception
            Console.WriteLine("error {0}", ex)
            Return 0
        End Try
        '    Catch ex As Exception
        Return 1
    End Function

    Public Function CrearNuevaRevision(ByVal asunto As String, ByVal fechaInicio As String, ByVal fechaLimite As String, ByVal cuerpoTarea As String, ByVal ubicacion As String, invitados As List(Of String)) As String
        Try

            ' Create an Outlook application.
            Dim oApp As Outlook.Application = New Outlook.Application()
            ' Get Mapi NameSpace and Logon.
            Dim oNS As Outlook.NameSpace = oApp.GetNamespace("mapi")
            ' Create an AppointmentItem.
            Dim oAppt As Outlook._AppointmentItem = oApp.CreateItem(Outlook.OlItemType.olAppointmentItem)
            'oAppt.Display(true)  'Modal
            ' Change AppointmentItem to a Meeting. 
            oAppt.MeetingStatus = Outlook.OlMeetingStatus.olMeeting
            ' Set the properties
            oAppt.Subject = asunto
            oAppt.Body = cuerpoTarea
            oAppt.Location = ubicacion
            oAppt.Start = Convert.ToDateTime(fechaInicio)
            oAppt.End = Convert.ToDateTime(fechaLimite)
            oAppt.ReminderSet = True
            oAppt.ReminderMinutesBeforeStart = 60
            oAppt.BusyStatus = Outlook.OlBusyStatus.olBusy  '  olBusy
            oAppt.IsOnlineMeeting = False
            oAppt.AllDayEvent = False

            ' Add attendees.
            Dim oRecipts As Outlook.Recipients = oAppt.Recipients
            ' Add required attendee.
            Dim oRecipt As Outlook.Recipient

            For Each invitado In invitados
                Console.WriteLine(invitado)
                oRecipt = oRecipts.Add(invitado)
                oRecipt.Type = Outlook.OlMeetingRecipientType.olRequired
                oRecipts.ResolveAll()
            Next
            oAppt.Send()
            ' Logoff.
            oNS.Logoff()

        Catch ex As Exception
            Console.WriteLine("error {0}", ex)
            Return 0
        End Try
        '    Catch ex As Exception
        Return 1
    End Function

    Sub Main()
        'Console.WriteLine("{0} paginas ", getNumeroPaginas())
        Dim mails As List(Of String) = New List(Of String)
        mails.Add("jhonatan.bazalduao@uanl.mx")
        mails.Add("jhona.22.baz@gmail.com")
        mails.Add("jhona.baz@outlook.com")

        Dim FechaInicio As String = "20/03/2014 07:30 p. m."
        Dim FechaFin As String = "20/03/2014 10:30 p. m."
        Dim Cuerpo As String = "Otro cuerpo de una revision directiva prueba"
        Dim Asunto As String = "revision directiva con parametros."
        Dim ubicacion As String = "UANL, La flama"


        'If CrearNuevaRevision(Asunto, FechaInicio, FechaFin, Cuerpo, ubicacion, mails) Then
        'Console.WriteLine("modificacion correcta ")
        'End If

        If ModifyCalendar(mails) Then
            Console.WriteLine(" OK    ")
        End If

        'If SendMail2() Then
        'Console.WriteLine("modificacion correcta ")
        'End If

        Console.ReadLine()
    End Sub

End Module