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
             
                string filepath = @"C:\Windows\System\config.xml";
               // string filepath = AppDomain.CurrentDomain.BaseDirectory + "config.xml";
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

                 p.conn();
                // p.leerbd();
                p.Project();
                p.insertProject();


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
             
            }

        }

        private void escribir_error() {



        }

        /// <sumary>
        /// Funcion que permite escribir en un archivo tipo texto las transacciones realizadas por la aplicacion 
        /// </summary>
        /// <param name="mensaje"> son los datos procesados </param>
        /// <param name="tipo"> tipo de datos procesados insert, delete , update</param>
        /// 
        private void escribir_log( string mensaje , string tipo) {

          
            string filePath = log + @"\logProject.txt";
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine("Datos procesados :" + mensaje + Environment.NewLine + tipo +""+ Environment.NewLine + "Dia y Hora de proceso :" + DateTime.Now.ToString());
                writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
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
                string sql = "SELECT project_id, start_date, end_date,progress,updated_at, name  from projects";
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
                            string mensaje ="Project Name :"+row[5].ToString()+", start date : " + row[1].ToString()+", end_date : "+ row[2].ToString()+", % progress : "+ Convert.ToInt16(row[3])+"";
                            Console.WriteLine("\n{0}. {1}   {2} \t{3} \n lista de datos actualizados", row[0].ToString(), row[1].ToString(), row[2].ToString(), Convert.ToInt16(row[3]));
                            escribir_log(mensaje,"se ctualizaron estos registro en Project Online desde Mysql");

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
                mail.Subject = "Historico de procesos de transacciones desde Project Server Online";
                mail.Body = "Se ha producido un historico de las transacciones realizadas logProject.txt , si se trata de un error de APP buscar en el archivo error.txt en la ruta configurada, Gracias ";
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(log + @"\logProject.txt");
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

        private void insertProject()
        {


            /*
                     project_New.Add(Guid);//Project_id IdDelProyecto
                    project_New.Add(pubProj.Name);//name NombreDeProyecto
                    project_New.Add(pubProj.Description);//description DescripciónDelProyecto
                  
                    //grouper AgrupadordeProyecto
                    //compromise Compromiso      No va el campo              
                    project_New.Add(pubProj.StartDate.ToShortDateString());//start_date ComienzoAnticipadoDelProyecto
                    project_New.Add(pubProj.FinishDate.ToShortDateString());//end_date FechaDeFinalizaciónDelProyecto
                    //institution InstitucióndelEstado
                    //action_line LíneadeAcción
                    //responsable SeguimientodeProyecto
                    project_New.Add(pubProj.CreatedDate.ToShortDateString());
             INSERT INTO `AIGDB_SSEC`.`projects` (`project_id`, `name`, `description`, `grouper`, `compromise`, 
             `start_date`, `end_date`, `institution`, `action_line`, `responsable`
             */

            try
            {

              
                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);
                var gui = project_New[0];
                var nombre = project_New[1];
                var descripcion = project_New[2];
                var finicio = project_New[3];
                var ffin = project_New[4];
                var fcreacion = project_New[5];



                string sql = " INSERT INTO `AIGDB_SSEC`.`projects` (`project_id`, `name`, `description`, `grouper`, `compromise`, `start_date`, `end_date`, `institution`, `action_line`, `responsable`) VALUES ('" + gui + "','" + nombre + "','" + descripcion + "',NULL,NULL,NULL,'"+finicio+"','"+ffin+"',NULL,NULL,NULL)";
                sql += " SELECT gui, nombre, fecha ";
                sql += " WHERE gui <> " + gui;

                if (connect.State != ConnectionState.Open)
                {
                    connect.Open();
                }
                MySqlCommand cmd = new MySqlCommand(sql, connect);
                cmd.ExecuteNonQuery();
                connect.Close();

            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }

        private void Project()
        {
            int j = 1;

            using (ProjectCont1)
            {

                DateTime dia = DateTime.Today.AddDays(0);
                DateTime hoy = DateTime.Today;
                DateTime ayer = hoy.AddDays(0);

                // 1. Retrieve the project, tasks, etc.
                var projCollection = ProjectCont1.LoadQuery(ProjectCont1.Projects
                    .Where(p => p.CreatedDate >= ayer)
                    .Include(
                        p => p.Id,
                        p => p.Name,
                        p => p.Tasks,
                        p => p.Tasks.Include(
                            t => t.Id,
                            t => t.Name,
                            t => t.CustomFields,
                            t => t.CustomFields.IncludeWithDefaultProperties(
                                cf => cf.LookupTable,
                                cf => cf.LookupEntries
                            )
                        )
                    )
                );


                ProjectCont1.ExecuteQuery();


                PublishedProject theProj = projCollection.First();
                Console.WriteLine("Name:\t{0}", theProj.Name);
                Console.WriteLine("Id:\t{0}", theProj.Id);
               //Console.WriteLine("Tasks count: {0}", theProj.Tasks.Count);
                Console.WriteLine("  -----------------------------------------------------------------------------");


                PublishedTaskCollection taskColl = theProj.Tasks;
                PublishedTask theTask = taskColl.First();
                CustomFieldCollection LCFColl = theTask.CustomFields;
                Dictionary<string, object> taskCF_Dict = theTask.FieldValues;

                foreach (CustomField cf in LCFColl)
                {
                    String textValue = taskCF_Dict[cf.InternalName].ToString();
                    Console.WriteLine("", cf.FieldType, cf.Name, textValue);
                    var cia = cf.LookupTable.Entries.Where(e => e.InternalName == "Instituciones del Estado");
                    ProjectCont1.ExecuteQuery();
                    var a = cia.First().FullValue;
                    var b = cia.First().Description;  
                }



                foreach (PublishedProject pubProj in projCollection)
                {

                    //
            

                        //


                        string Guid = pubProj.Id.ToString();
                    //Lista de Proyectos del lado de Project para insertar
                    project_New = new List<string>();
                    project_New.Add(Guid);//Project_id IdDelProyecto
                    project_New.Add(pubProj.Name);//name NombreDeProyecto
                    project_New.Add(pubProj.Description);//description DescripciónDelProyecto
        
                   

                    //grouper AgrupadordeProyecto
                    //compromise Compromiso      No va el campo              
                    project_New.Add(pubProj.StartDate.ToShortDateString());//start_date ComienzoAnticipadoDelProyecto
                    project_New.Add(pubProj.FinishDate.ToShortDateString());//end_date FechaDeFinalizaciónDelProyecto
                    //institution InstitucióndelEstado
                    //action_line LíneadeAcción
                    //responsable SeguimientodeProyecto
                    project_New.Add(pubProj.CreatedDate.ToShortDateString());

                 

                    //lista de GUI del lado de Project para comparar y borrar
                    ssec_gui = new List<string>();
                    ssec_gui.Add(Guid);

                    Console.WriteLine("\n{0}. {1}   {2} \t{3} \n", j++, pubProj.Id, pubProj.Name, pubProj.CreatedDate);
                    //intento de comparar dos matrices multidimencionales //comparar dos listas y buscar las diferencas
                    //var project = new List<string> { Guid + "," + pubProj.Name + "," + pubProj.CreatedDate };
                    //var projectGUI = new List<string> { Guid };
                    //var ssec = new List<string> { "datos del Mysql "};
                    //var projectFaltan = projectGUI.Except(ssec.ToList()); //list3 contains only 1, 2

                }
            }




            insertProject();
        }

        private void DeleteProject()
        {
                       
            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);

                if (connect.State != ConnectionState.Open) { connect.Open(); }

                var projectFaltan = project_gui.Except(ssec_gui.ToList());

                foreach (string gui in projectFaltan)
                {
                    string sql = " Update projects ";
                    sql += " WHERE gui =" + gui;
                    MySqlCommand cmd = new MySqlCommand(sql, connect);
                    cmd.ExecuteNonQuery();
                }

                connect.Close();


            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }
    }
}
