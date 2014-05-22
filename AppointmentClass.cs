using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;


namespace OutlookClassLibrary
{
    public class AppointmentClass
    {                  
        public string CreateNewAppointmentE(string subject_, string StartDate_,string finisDate_,string body_,string location_, List<string> attendees_)
        {
            string errorAppointment = "";
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.UseDefaultCredentials = true; //sistema.calidad CG29dgi
                //service.Credentials = new WebCredentials("sistema.calidad", "CG29dgi", "uanl");
                service.Credentials = new WebCredentials("jhonatan.bazalduao", "ZB18dgi", "uanl");//@uanl.mx                
                service.AutodiscoverUrl("jhonatan.bazalduao@uanl.mx");

                Appointment appointment = new Appointment(service);
                appointment.Subject = subject_;
                appointment.Body = body_;
                appointment.Start = Convert.ToDateTime(StartDate_);
                appointment.End = Convert.ToDateTime(finisDate_);
                appointment.Location = location_;

                foreach (string attendee in attendees_)                 
                appointment.RequiredAttendees.Add(attendee);
                
                appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
                
                return errorAppointment;
            }
            catch
            {   
                return errorAppointment = "an error trying to send the appoitment happend" ;
            }
        }        
        public string actualizarRevisionDirectivaE(string asunto, string fechaInicio, string fechaLimite, string cuerpoTarea, string ubicacion) 
        {
            string error = "";
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.UseDefaultCredentials = true; //sistema.calidad CG29dgi
                //service.Credentials = new WebCredentials("sistema.calidad", "CG29dgi", "uanl");
                service.Credentials = new WebCredentials("jhonatan.bazalduao", "ZB18dgi", "uanl");//@uanl.mx                
                service.AutodiscoverUrl("jhonatan.bazalduao@uanl.mx");

                string querystring = "Subject:'"+asunto +"'";
                ItemView view = new ItemView(1);

                FindItemsResults<Item> results = service.FindItems(WellKnownFolderName.Calendar, querystring, view);// <<--- Esta parte no la se pasar a VB
                if (results.TotalCount > 0)
                {
                    if (results.Items[0] is Appointment) //if is an appointment, could be other different than appointment 
                    {
                        Appointment appointment = results.Items[0] as Appointment; //<<--- Esta parte no la se pasar a VB

                        if (appointment.MeetingRequestWasSent)//if was send I will update the meeting  
                        {

                            appointment.Start = Convert.ToDateTime(fechaInicio);
                            appointment.End = Convert.ToDateTime(fechaLimite);
                            appointment.Body = cuerpoTarea;
                            appointment.Location = ubicacion;
                            appointment.Update(ConflictResolutionMode.AutoResolve);                            
                        }
                        else//if not, i will modify and sent it
                        {
                            appointment.Start = Convert.ToDateTime(fechaInicio);
                            appointment.End = Convert.ToDateTime(fechaLimite);
                            appointment.Body = cuerpoTarea;
                            appointment.Location = ubicacion;
                            appointment.Save(SendInvitationsMode.SendOnlyToAll);
                        }
                    }
                }
                else
                {                    
                    error = "Wasn't found it's appointment";
                    return error;
                }
                return error;
            }
            catch
            {                
                return error = "an error happend";
            }                       
        }           
        public string crearTareaE(string asunto, string fechaInicio, string fechaLimite, string cuerpoTarea, string correoDestino) 
        {
            string errorTask = "";
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.UseDefaultCredentials = true; //sistema.calidad SC29dgi                
                service.Credentials = new WebCredentials("jhonatan.bazalduao", "ZB18dgi", "uanl");
                service.AutodiscoverUrl("jhonatan.bazalduao@uanl.mx");
                
                Task taskItem = new Task(service);
                taskItem.Subject = asunto;
                taskItem.Body = new MessageBody(cuerpoTarea);
                taskItem.StartDate = Convert.ToDateTime(fechaInicio);
                taskItem.DueDate = Convert.ToDateTime(fechaLimite);
                //Aqui no se como agregar el correo destino, aun no encuentro el metodo para hacer eso.......                                               
                //taskItem.Save();
                taskItem.Save();

                //service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "jhonatan.bazalduao@uanl.mx");                

                /*Folder newFolder = new Folder(service);
                newFolder.DisplayName = "TestFolder1";
                newFolder.Save(WellKnownFolderName.Inbox);
                */
                return errorTask;
            }
            catch(Exception e)
            {
                errorTask = e.Message;                
                return errorTask;
            }
        }
        public string ActualizarRevisionDirectivaCorreosE(string asunto, List<string> invitados )
        {
            string error = "";
            try
            {
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);
                service.UseDefaultCredentials = true; //sistema.calidad CG29dgi
                //service.Credentials = new WebCredentials("sistema.calidad", "CG29dgi", "uanl");
                service.Credentials = new WebCredentials("jhonatan.bazalduao", "ZB18dgi", "uanl");//@uanl.mx                
                service.AutodiscoverUrl("jhonatan.bazalduao@uanl.mx");

                string querystring = "Subject:'" + asunto + "'";
                ItemView view = new ItemView(1);

                FindItemsResults<Item> results = service.FindItems(WellKnownFolderName.Calendar, querystring, view);// <<--- Esta parte no la se pasar a VB
                if (results.TotalCount > 0)
                {
                    if (results.Items[0] is Appointment) //if is an appointment, could be other different than appointment 
                    {
                        Appointment appointment = results.Items[0] as Appointment; //<<--- Esta parte no la se pasar a VB

                        if (appointment.MeetingRequestWasSent)//if was send I will update the meeting  
                        {
                           foreach (string invitado in invitados){
                               appointment.RequiredAttendees.Add(invitado);                               
                           }
                            appointment.Update(ConflictResolutionMode.AutoResolve);                            
                        }
                        else//if not, i will modify and sent it
                        {
                            appointment.Save(SendInvitationsMode.SendOnlyToAll);                            
                        }
                    }
                }
                else
                {                    
                    error = "Wasn't found it's appointment";
                    return error;
                }
                return error;
            }
            catch (Exception ex)
            {                
                return error = ex.Message;
            }
        }                  
    }
}
