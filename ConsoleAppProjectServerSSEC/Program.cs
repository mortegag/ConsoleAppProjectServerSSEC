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
        const int PROJECT_BLOCK_SIZE = 20;
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
                 p.Project();


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
                            escribir_log(mensaje,"se actualizaron estos registro en Project Online desde Mysql");

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
                    //tsk.Start = Convert.ToDateTime(fi);
                    //tsk.Finish = Convert.ToDateTime(ff);
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

        private void insertProject(string project_id, string name, string description, string grouper , string compromise, DateTime start_date, DateTime end_date, string institution, string action_line , string responsable  )
        {
            try
            {
                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);
                string sql = "INSERT INTO `AIGDB_SSEC`.`projects` (`project_id`, `name`, `description`, `grouper`, `compromise`, `start_date`, `end_date`, `institution`, `action_line`, `responsable`) "; 
                      sql += " VALUES ('" + project_id + "','" + name + "','" + description + "','"+grouper+"','"+compromise+"','"+start_date+"','"+end_date+"','"+institution+"','"+action_line+"','"+responsable+"');";
              //  sql += " SELECT gui from `AIGDB_SSEC`.`projects` ";
              //  sql += " WHERE gui <> " + project_id;
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
                string mensaje = ex.ToString();
                escribir_log("Hubo un error al tratar de Insertar registros : ", mensaje);
              

            }
        }

        private void Project()
        {
            using (ProjectCont1)
            {

                //************************************
                ProjectCont1.Load(ProjectCont1.Projects, qp => qp.Include(qr => qr.Id));
                ProjectCont1.ExecuteQuery();
                var allIds = ProjectCont1.Projects.Select(p => p.Id).ToArray();
                int numBlocks = allIds.Length / PROJECT_BLOCK_SIZE + 1;
                for (int i = 0; i < numBlocks; i++)
                {
                    var idBlock = allIds.Skip(i * PROJECT_BLOCK_SIZE).Take(PROJECT_BLOCK_SIZE);
                    Guid[] block = new Guid[PROJECT_BLOCK_SIZE];
                    Array.Copy(idBlock.ToArray(), block, idBlock.Count());
                    DateTime hoy = DateTime.Today;
                    DateTime ayer = hoy.AddDays(-10);
                    string last = ayer.ToShortDateString();

                    var projBlk = ProjectCont1.LoadQuery(
                         ProjectCont1.Projects
                        .Where(p =>
                            p.Id == block[0] || p.Id == block[1] ||
                            p.Id == block[2] || p.Id == block[3] ||
                            p.Id == block[4] || p.Id == block[5] ||
                            p.Id == block[6] || p.Id == block[7] ||
                            p.Id == block[8] || p.Id == block[9] ||
                            p.Id == block[10] || p.Id == block[11] ||
                            p.Id == block[12] || p.Id == block[13] ||
                            p.Id == block[14] || p.Id == block[15] ||
                            p.Id == block[16] || p.Id == block[17] ||
                            p.Id == block[18] || p.Id == block[19]
                        )
                        .Include(p => p.Id,
                            p => p.Name,
                            p => p.Description,
                            p => p.StartDate,
                            p => p.FinishDate,
                            p => p.CreatedDate,
                            p => p.IncludeCustomFields,
                            p => p.IncludeCustomFields.CustomFields,
                            P => P.IncludeCustomFields.CustomFields.IncludeWithDefaultProperties(
                                lu => lu.LookupTable,
                                lu => lu.LookupEntries
                            )
                        )
                    );

                    ProjectCont1.ExecuteQuery();

                    foreach (PublishedProject pubProj in projBlk)
                    {

                        DateTime fechaP = Convert.ToDateTime(pubProj.CreatedDate.ToShortDateString());
                        DateTime fechaA = Convert.ToDateTime(ayer.ToShortDateString());
                        //
                        string project_id = pubProj.Id.ToString();
                        string name = pubProj.Name;
                        string description = pubProj.Description;
                        DateTime start_date = pubProj.StartDate;
                        DateTime end_date = pubProj.FinishDate;
                        string grouper = "";
                        string compromise = "";
                        string institution = "";
                        string action_line = "";
                        string responsable = "";


                        if (fechaP >= fechaA)
                        {

                            var projECFs = pubProj.IncludeCustomFields.CustomFields;
                            Dictionary<string, object> ECFValues = pubProj.IncludeCustomFields.FieldValues;
                            //valores nativos
               

                            int j = 0;
                            foreach (CustomField cf in projECFs)
                            {
                                j++;

                                String[] entries = (String[])ECFValues[cf.InternalName];

                                foreach (String entry in entries)
                                {                                
                                    var luEntry = ProjectCont1.LoadQuery(cf.LookupTable.Entries
                                        .Where(e => e.InternalName == entry));

                                    ProjectCont1.ExecuteQuery();

                                    if (j == 1) { grouper = luEntry.First().FullValue; Console.WriteLine(" {0} {1, -22}  ", j, luEntry.First().FullValue); }
                                    else if (j == 3) { compromise = luEntry.First().FullValue; Console.WriteLine("{0} {1, -22}", j, luEntry.First().FullValue); }
                                    else if (j == 6) { institution = luEntry.First().FullValue; Console.WriteLine("{0} {1, -22}", j, luEntry.First().FullValue); }
                                    else if (j == 7) { action_line = luEntry.First().FullValue; Console.WriteLine("{0} {1, -22}", j, luEntry.First().FullValue); }
                                    else if (j == 9) { responsable = luEntry.First().FullValue; Console.WriteLine("{0} {1, -22}", j, luEntry.First().FullValue); }
                                    
                                }
                              
                            }
                            string mensaje = " GUI : " + project_id + " NOMBRE : " + name + " DESCRIPCION : " + description + " GRUPO :" + grouper + " COMPROMISO " + compromise + " DIA INICIO :" + start_date + " DIA FIN :" + end_date + " INSTITUCION : " + institution + " LINEA ACCION :" + action_line + " REPONSABLE :" + responsable + "";
                            insertProject(project_id, name, description, grouper, compromise, start_date, end_date, institution, action_line, responsable);
                            escribir_log(mensaje, "se insertaron estos registros desde Project Online a Mysql");
                        }
                    }
                }
            }

            Console.Write("\nPress any key to exit: ");
            Console.ReadKey(false);
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
