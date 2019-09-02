using System;
using System.Collections.Generic;
using System.Data;

using System.Text;
using System.Timers;
using System.IO;
using System.Xml;
using MySql.Data.MySqlClient;
using Timer = System.Timers.Timer;
using System.Linq;
using System.Security;
using csom = Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;
using Microsoft.ProjectServer.Client;
using System.Net.Mail;

namespace ConsoleAppProjectServerSSEC
{
 



    class Program
    {

        string Resultados = "";
        string rutas = "";
        string ps = "";
        string hora1 = "";
        string hora2 = "";
        string log = "";
        string ip = "";
        string user = "";
        string passw = "";
        string db = "";
        String ID = "";
        string csv = "";
        string intervalos = "";
        string smtpserver = "";
        string de = "";
        string credencial = "";
        string envioemail = "";
        string para = "";
        string pwaPath = "";
        string userName = "";
        string passWord = "";
        static csom.ProjectContext ProjectCont1;
        MySqlConnection connect = new MySqlConnection();
        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
        Timer tmrExecutor = new Timer();

        /*
        listado dep Proyectos nuevos del lado de Project Online de los 
        @project_New para insertar proyectos a MySql
        ultimos 5 dias actuales
        */
        List<string> project_New;
        /* Listado indices GUI de proyectos del lado de MYSQL
         */
        List<string> ssec_gui;
        /*Listado de indices GUI de proyectos del lado de Project Online
         */
        List<string> project_gui;


        static void Main(string[] args)
        {

            Program p = new Program();

            try
            {
                MySqlConnection connect = new MySqlConnection();
             
               // string filepath = @"C:\Windows\System\config.xml";
                string filepath = AppDomain.CurrentDomain.BaseDirectory + @"\config.xml";
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(filepath);
                foreach (XmlNode row in xmlDoc.SelectNodes("//SoftwareToInstallPaths"))
                {
              
                    //intervalo de tiempo en que se ejecutara el Servicio
                    p.intervalos = row.SelectSingleNode("//interval").InnerText;
                    //Ruta de bitacora de eventos ejecutados
                    p.log = row.SelectSingleNode("//Logs").InnerText;
                    //datos de coneccion al servidor MYSql
                    p.ip = row.SelectSingleNode("//ip").InnerText;
                    p.user = row.SelectSingleNode("//user").InnerText;
                    p.passw = row.SelectSingleNode("//passw").InnerText;
                    p.db = row.SelectSingleNode("//db").InnerText;
                    //Datos de connecion a Project Server 
                    p.pwaPath = row.SelectSingleNode("//pwaPath").InnerText;
                    p.userName = row.SelectSingleNode("//userName").InnerText;
                    p.passWord = row.SelectSingleNode("//passWord").InnerText;
                    //Datos de envio de correos electronicos
                    p.smtpserver = row.SelectSingleNode("//smtpserver").InnerText;
                    p.de = row.SelectSingleNode("//de").InnerText;
                    p.credencial = row.SelectSingleNode("//credencial").InnerText;
                    p.envioemail = row.SelectSingleNode("//envioemail").InnerText;
                    p.para = row.SelectSingleNode("//para").InnerText;
                }
                //p.tmrExecutor.Elapsed += new ElapsedEventHandler(p.tmrExecutor_Elapsed);
                //p.tmrExecutor.Interval = Convert.ToInt16(p.intervalos); // cada 24 horas son 86400000; //cada 5 segundo son 5000 ;
                //p.tmrExecutor.Enabled = true;
                //p.tmrExecutor.Start();
                p.conn();
                p.leerbd();
        
            }
            catch (Exception ex)
            {
                string error = ex.ToString();
                string filePath = p.log + @"\error.txt";
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine("Error :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace +
                       "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                    writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                }

        

            }
            finally
            {
                string mensaje = "Proceso de integracion de Project Online";
                string filePath = p.log + @"\bitacora.txt";
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine("Message :" + mensaje + "<br/>" + Environment.NewLine + "ver archivo csv" +
                     "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                    writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                }
            }

        }

        private void tmrExecutor_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
         

            try
            {
                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);

                //(string gui, string fi, string ff, string duracion, int porcent)
                string sql = "SELECT project_id, start_date, end_date,progress,updated_at  from projects";
                sql += " where updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR  created_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR created_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " ORDER BY id";


                if (connect.State != ConnectionState.Open)
                {
                    connect.Open();
                }
                using (MySqlDataAdapter da = new MySqlDataAdapter(sql, connect))
                {

                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt != null)
                    {


                        foreach (DataRow row in dt.Rows)
                        {
                            //Funcion de actualizacion de resitros en el Project
                            UddateTask(row[0].ToString(), row[1].ToString(), row[2].ToString(), Convert.ToInt16(row[3]));
                            //Lista de GUI lado Mysql para Borrar
                            ssec_gui = new List<string>();
                            ssec_gui.Add(row[0].ToString());

                        }


                    }
                    connect.Close();
                }


            }
            catch (Exception ex)
            {
                string error = ex.ToString();
                //string filePath = log + @"\error.txt";
                string filePath = log + @"\error.txt";
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine("Error :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace +
                       "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                    writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                }

            }
            finally
            {

                string mensaje = "Proceso de integracion de Project Online";
                string filePath = log + @"\bitacora.txt";
                using (StreamWriter writer = new StreamWriter(filePath, true))
                {
                    writer.WriteLine("Message :" + mensaje + "<br/>" + Environment.NewLine + "ver archivo csv" +
                       "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                    writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                }


            }
        }

        private void conn()
        {

            ProjectCont1 = new csom.ProjectContext(pwaPath);
            SecureString securePassword = new SecureString();
            foreach (char c in passWord.ToCharArray())
            {
                securePassword.AppendChar(c);
            }
            ProjectCont1.Credentials = new SharePointOnlineCredentials(userName, securePassword);
            Console.WriteLine("Se conecto a Project Server ");
        }

        private void leerbd()
        {

            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);

                //(string gui, string fi, string ff, string duracion, int porcent)
                string sql = "SELECT project_id, start_date, end_date,progress,updated_at  from projects";
                sql += " where updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR  created_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR created_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " OR updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL 6 DAY), '%Y-%m-%d')";
                sql += " ORDER BY id";


                if (connect.State != ConnectionState.Open)
                {
                    connect.Open();
                }
                using (MySqlDataAdapter da = new MySqlDataAdapter(sql, connect))
                {

                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    if (dt != null)
                    {


                        foreach (DataRow row in dt.Rows)
                        {
                            //Funcion de actualizacion de resitros en el Project
                            UddateTask(row[0].ToString(), row[1].ToString(), row[2].ToString(), Convert.ToInt16(row[3]));
                            //Lista de GUI lado Mysql para Borrar
                            ssec_gui = new List<string>();
                            ssec_gui.Add(row[0].ToString());

                            Console.WriteLine("\n{0}. {1}   {2} \t{3} \n lista de datos actualizados", row[0].ToString(), row[1].ToString(), row[2].ToString(), Convert.ToInt16(row[3]));

                        }


                    }
                    connect.Close();
                }

            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }


        public void enviarCorreo()
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(smtpserver);

                mail.From = new MailAddress(de);
                mail.To.Add(new MailAddress(para));
                mail.Subject = "Error de APP Servicio de Directorio Activo OIRH";
                mail.Body = "Se ha producido un error en el APP de Servicios de Directorio activo SARH , favor ver archivo de error adjunto";
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(log + @"\error.txt");
                mail.Attachments.Add(attachment);

                SmtpServer.Port = 25;
                SmtpServer.Credentials = new System.Net.NetworkCredential(de, credencial);
                SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);

            }
            catch (Exception ex)
            {

            }

        }

        private void detener()
        {
            tmrExecutor.Enabled = false;
        }

        private void UddateTask(string gui, string fi, string ff, int porcent)
        {

            using (ProjectCont1)
            {

                Guid ProjectGuid = new Guid(gui);
                var projCollection = ProjectCont1.LoadQuery(
                 ProjectCont1.Projects
                   .Where(p => p.Id == ProjectGuid));
                ProjectCont1.ExecuteQuery();
                csom.PublishedProject proj2Edit = projCollection.First();
                DraftProject draft2Edit = proj2Edit.CheckOut();
                ProjectCont1.Load(draft2Edit);
                ProjectCont1.Load(draft2Edit.Tasks);
                ProjectCont1.ExecuteQuery();
                //
                var tareas = draft2Edit.Tasks;
                foreach (DraftTask tsk in tareas)
                {
                    tsk.Start = Convert.ToDateTime(fi);
                    tsk.Finish = Convert.ToDateTime(ff);
                    //tsk.Duration = duracion;
                    tsk.PercentComplete = porcent;
                }

                draft2Edit.Publish(true);
                csom.QueueJob qJob = ProjectCont1.Projects.Update();
                csom.JobState jobState = ProjectCont1.WaitForQueue(qJob, 200);
                //
                qJob = ProjectCont1.Projects.Update();
                jobState = ProjectCont1.WaitForQueue(qJob, 20);

                if (jobState == JobState.Success)
                {

                }

            }
        }

    }
}
