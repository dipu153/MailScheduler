using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Net.Mail;
using System.Data.SqlClient;
using RestSharp;
using Newtonsoft.Json;

namespace WindowsServiceCS
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Simple Service stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                this.WriteToFile("Simple Service Mode: " + mode + " {0}");

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                this.WriteToFile("Simple Service scheduled to run after: " + schedule + " {0}");

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void SchedularCallback(object e)
        {
            try
            {
                var url = $"http://localhost:49378/api/Employees";
                var client = new RestClient(url);
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddParameter("text/plain", "", ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                var employees = JsonConvert.DeserializeObject<List<Emp>>(response.Content);

                foreach (var employee in employees)
                {
                    string name = employee.FirstName + " " + employee.LastName;
                    string email = employee.Email;
                    WriteToFile("Trying to send email to: " + name + " " + email);
                    MailMessage mm = new MailMessage("<FromMail>", email);
                    mm.IsBodyHtml = true;
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.EnableSsl = true;
                    System.Net.NetworkCredential credentials = new System.Net.NetworkCredential();
                    credentials.UserName = "<FromMail>";
                    credentials.Password = "<password>";
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = credentials;
                    smtp.Port = 587;


                    var updateReminder = $"http://localhost:49378/api/Employees?email={email}";
                    var clientUpdate = new RestClient(updateReminder);
                    clientUpdate.Timeout = -1;
                    var requestUpdate = new RestRequest(Method.PUT);
                    requestUpdate.AddHeader("Content-Type", "application/json");

                    if (employee.Reminder == 0)
                    {
                        if (employee.VisitLink == false && employee.Reminder == 0)
                        {
                            mm.Subject = "Activation Link for " + name + "";
                            mm.Body = $"<a href='http://localhost:49378/api/Employees?email={email}&visit=True'>Click Here</a><b>to activate</b><br /><br />";
                            IRestResponse responseUpdate = clientUpdate.Execute(requestUpdate);
                            smtp.Send(mm);
                        }
                    }

                    if (employee.VisitLink == false)
                    {
                        if (employee.Reminder > 0 && employee.Reminder <= 3)
                        {
                            mm.Subject = "Reminder mail -- " + name + " (Reminders sent - " + employee.Reminder + ")";
                            mm.Body = $"<a href='http://localhost:49378/api/Employees?email={email}&visit=True'>Click Here</a><b>to activate</b><br /><br />";
                            IRestResponse responseUpdate = clientUpdate.Execute(requestUpdate);
                            WriteToFile("Reminder send to " + email);
                            smtp.Send(mm);
                        }
                    }
                    WriteToFile("Email sent successfully to: " + name + " " + email);
                }
                this.ScheduleService();
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void WriteToFile(string text)
        {
            string path = "C:\\ServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }
    }
    public class Emp
    {
        public int ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public Nullable<bool> VisitLink { get; set; }
        public Nullable<int> Reminder { get; set; }
    }
}
