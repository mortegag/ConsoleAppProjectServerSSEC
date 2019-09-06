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

    public static class Seguridad
    {
        /// Encripta una cadena
        public static string Encriptar(this string _cadenaAencriptar)
        {
            string result = string.Empty;
            byte[] encryted = System.Text.Encoding.Unicode.GetBytes(_cadenaAencriptar);
            result = Convert.ToBase64String(encryted);
            return result;
        }
        /// Esta función desencripta la cadena que le envíamos en el parámentro de entrada.
        public static string DesEncriptar(this string _cadenaAdesencriptar)
        {
            string result = string.Empty;
            byte[] decryted = Convert.FromBase64String(_cadenaAdesencriptar);
            //result = System.Text.Encoding.Unicode.GetString(decryted, 0, decryted.ToArray().Length);
            result = System.Text.Encoding.Unicode.GetString(decryted);
            return result;
        }
    }



    class Program
    {
        const int PROJECT_BLOCK_SIZE = 20;
        string log = "";
        string ip = "";
        string user = "";
        string passw = "";
        string db = "";
        string intervalos = "";
        string smtpserver = "";
        string de = "";
        string credencial = "";
        string envioemail = "";
        string para = "";
        string pwaPath = "";
        string userName = "";
        string passWord = "";
        string insertar = "";
        string actualizar = "";
        string dias_i = "";
        string dias_a = "";
        string project_id;
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
                    //incluir parametros para actualizar , insertar o borrar 
                    p.insertar = row.SelectSingleNode("//insert").InnerText;
                    p.actualizar = row.SelectSingleNode("//update").InnerText;
                    //parametros para cantidad de dias para transacciones 
                    p.dias_i = row.SelectSingleNode("//rango_insert").InnerText;
                    p.dias_a = row.SelectSingleNode("//rango_update").InnerText;

                }


                if (p.passWord == "") //Contraseña de Project Server
                {
                    Console.Write("Introdusca su clave para ProjectServerOnline : ");
                    Console.ForegroundColor = ConsoleColor.Black;
                    p.passWord = Console.ReadLine();
                    Console.ForegroundColor = ConsoleColor.White;
                    //
                    XmlNode node = xmlDoc.SelectSingleNode("parameters/SoftwareToInstallPaths/passWord");
                    node.InnerText = Seguridad.Encriptar(p.passWord);
                    xmlDoc.Save(filepath);
                    //
 
                }
                else { p.passWord = Seguridad.DesEncriptar(p.passWord);}

                if (p.passw == "") //Contraseña de Base de Datos
                {

                    Console.Write("Introdusca password BD MySql : ");
                    Console.ForegroundColor = ConsoleColor.Black;
                    p.passw = Console.ReadLine();
                    Console.ForegroundColor = ConsoleColor.White;
                    //
                    XmlNode node1 = xmlDoc.SelectSingleNode("parameters/SoftwareToInstallPaths/passw");
                    node1.InnerText = Seguridad.Encriptar(p.passw);
                    xmlDoc.Save(filepath);
                    //
                }
                else { p.passw = Seguridad.DesEncriptar(p.passw); }


                if (p.credencial == "") //Contraseña de SMTP
                {

                    Console.Write("Introdusca password SMTP : ");
                    Console.ForegroundColor = ConsoleColor.Black;
                    p.credencial = Console.ReadLine();
                    Console.ForegroundColor = ConsoleColor.White;
                    //
                    XmlNode node = xmlDoc.SelectSingleNode("parameters/SoftwareToInstallPaths/credencial");
                    node.InnerText = Seguridad.Encriptar(p.credencial);
                    xmlDoc.Save(filepath);
                    //
                }
                else { p.credencial = Seguridad.DesEncriptar(p.credencial); }
                               
                p.conn();

                if (p.insertar == "Si") { p.CamposPersonalizadosProject(); }
                if (p.actualizar == "Si") { p.MysqltoProject(); }
                if (p.envioemail == "Si") { p.enviarCorreo(); }


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


        /// <sumary>
        /// Funcion que permite escribir en un archivo tipo texto las transacciones realizadas por la aplicacion 
        /// </summary>
        /// <param name="mensaje"> son los datos procesados </param>
        /// <param name="tipo"> tipo de datos procesados insert, delete , update</param>
        private void escribir_log(string mensaje, string tipo)
        {


            string filePath = log + @"\logProject.txt";
            using (StreamWriter writer = new StreamWriter(filePath, true))
            {
                writer.WriteLine("Datos procesados :" + mensaje + Environment.NewLine + tipo + "" + Environment.NewLine + "Dia y Hora de proceso :" + DateTime.Now.ToString());
                writer.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
            }


        }
        /// <sumary>
        /// Funcion que permite conectarse a Project Server por medio de parametros de un archivo XML 
        /// </summary>
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
        /// <sumary>
        /// Funcion que permite leer los registros de la base de datos de MYSQL para actualizar
        /// los campos nativos de Project Server 
        /// </summary>
        private void MysqltoProject()
        {

            try
            {

                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);

                //(string gui, string fi, string ff, string duracion, int porcent)
                string sql = "SELECT project_id, start_date, end_date,progress,updated_at, name  from " + db + ".projects";
                sql += " where updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL " + dias_a + " DAY), '%Y-%m-%d')";
                sql += " OR  created_at >= DATE_FORMAT((SYSDATE() - INTERVAL " + dias_a + " DAY), '%Y-%m-%d')";
                sql += " OR created_at >= DATE_FORMAT((SYSDATE() - INTERVAL " + dias_a + " DAY), '%Y-%m-%d')";
                sql += " OR updated_at >= DATE_FORMAT((SYSDATE() - INTERVAL " + dias_a + " DAY), '%Y-%m-%d')";
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
                            string mensaje = "Project Name :" + row[5].ToString() + ", start date : " + row[1].ToString() + ", end_date : " + row[2].ToString() + ", % progress : " + Convert.ToInt16(row[3]) + "";
                            Console.WriteLine("\n{0}. {1}   {2} \t{3} \n lista de datos actualizados", row[0].ToString(), row[1].ToString(), row[2].ToString(), Convert.ToInt16(row[3]));
                            escribir_log(mensaje, "se actualizaron estos registro en Project Online desde Mysql");

                        }

                        if (dt.Rows.Count == 0) { escribir_log("El rango de dias del parametro 'rango_update' = " + dias_a + "   no encontro regitros para las transacciones  ", "modifique el parametro a menos dias para el proceso de actualizacion "); }

                    }

                    connect.Close();
                }

            }
            catch (Exception ex)
            {
                ex.ToString();

            }
        }
        /// <sumary>
        /// Funcion que permite enviar un correo con  el archivo adjunto logProject.txt a uno o mas destinatarios ,
        /// se paramertiza en el archivo XML  , esta funcion llama a la funcion  UddateTask
        /// </summary>
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
        /// <sumary>
        /// Funcion que permite actualizar las tareas en Project Server  
        /// </summary>
        /// <param gui="String">Identificar Grafico Unico de proyecto de Project Server </param>
        /// <param fi="Date"> Fecha Inicial de Tarea</param>
        /// <param ff="Date"> Fecha Final de Tarea</param>
        /// <param porcent="Int"> Valor numerico de porcentaje de progreso de la tarea</param>
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
                    double costo = tsk.Cost;

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
        /// <sumary>
        /// Funcion  que permite insertar un nuevo proyecto a la base de datos de MYSQL ,
        /// verifica que previamente que no se duplique la informacion antes de insertarla
        /// </summary>
        /// <param project_id="String">Identificar Grafico Unico de proyecto del lado de MYSQL </param>
        /// 
        private void insertProject(string project_id, string name, string description, string grouper, string compromise, DateTime start_date, DateTime end_date, string institution, string action_line, string responsable, double monto, string ubicacion)
        {
            try
            {
                string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
                connect = new MySqlConnection(connectionString);
                string sql = "INSERT INTO `" + db + "`.`projects` (`project_id`, `name`, `description`, `grouper`, `compromise`, `start_date`, `end_date`, `institution`, `action_line`, `responsable`, `monto`,`ubicacion`) ";
                sql += " VALUES ('" + project_id + "','" + name + "','" + description + "','" + grouper + "','" + compromise + "','" + start_date + "','" + end_date + "','" + institution + "','" + action_line + "','" + responsable + "'," + monto + ",'" + ubicacion + "');";

                if (Existe(project_id) != true)
                {
                    if (connect.State != ConnectionState.Open) { connect.Open(); }
                    MySqlCommand cmd = new MySqlCommand(sql, connect);
                    cmd.ExecuteNonQuery();
                    connect.Close();
                    string mensaje = " GUI : " + project_id + " NOMBRE : " + name + " DESCRIPCION : " + description + " GRUPO :" + grouper + " COMPROMISO " + compromise + " DIA INICIO :" + start_date + " DIA FIN :" + end_date + " INSTITUCION : " + institution + " LINEA ACCION :" + action_line + " REPONSABLE :" + responsable + "";
                    escribir_log("Inserción  de registros desde Project Online a Mysql, Detalles : ", mensaje);
                }
                else { escribir_log(" No hubo insercion porrque el GUI existe en Mysql ", " No procesada la transaccion de INSERT "); }
            }
            catch (Exception ex)
            {
                string mensaje = ex.ToString();
                escribir_log("Hubo un error al tratar de Insertar registros : ", mensaje);

            }
        }
        /// <sumary>
        /// Funcion que permite la verificacion para NO insertar un nuevo proyecto existente 
        /// </summary>
        /// <param gui="String">Identificar Grafico Unico de proyecto de Project Server </param>
        private bool Existe(string GUI)
        {
            string connectionString = "server=" + ip + ";uid=" + user + ";pwd=" + passw + " ;database=" + db + ";";
            connect = new MySqlConnection(connectionString);
            string sql = " SELECT count(*) FROM aigdb_ssec.projects WHERE project_id='" + GUI + "';";
            MySqlCommand cmd = new MySqlCommand(sql, connect);
            //cmd.Parameters.AddWithValue("Id", GUI);
            cmd.Parameters.Add("Id", GUI);
            if (connect.State != ConnectionState.Open) { connect.Open(); }
            int count = Convert.ToInt32(cmd.ExecuteScalar());
            if (count == 0)
                return false;
            else
                return true;
        }
        /// <sumary>
        /// Funcion que permite listar todos los campos de Project Server tanto los nativos como los personalizados
        /// con el objetivo de insertar todos los proyectos NUEVOS en la base de datos de MYSQL
        /// utiliza la funcion insertProject para insertar el nuevo proyecto.
        /// </summary>
        private void CamposPersonalizadosProject()
        {
            using (ProjectCont1)
            {

                //************************************
                ProjectCont1.Load(ProjectCont1.Projects, qp => qp.Include(qr => qr.Id));
                ProjectCont1.ExecuteQuery();
                var allIds = ProjectCont1.Projects.Select(p => p.Id).ToArray();
                int numBlocks = allIds.Length / PROJECT_BLOCK_SIZE + 1;
                int x = 0;
                for (int i = 0; i < numBlocks; i++)
                {
                    var idBlock = allIds.Skip(i * PROJECT_BLOCK_SIZE).Take(PROJECT_BLOCK_SIZE);
                    Guid[] block = new Guid[PROJECT_BLOCK_SIZE];
                    Array.Copy(idBlock.ToArray(), block, idBlock.Count());
                    DateTime hoy = DateTime.Today;
                    DateTime ayer = hoy.AddDays(Convert.ToInt32(dias_i));
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
                        double monto = 0;
                        ProjectCont1.Load(pubProj.Tasks);
                        ProjectCont1.ExecuteQuery();
                        var tareas =  pubProj.Tasks;
                       foreach (PublishedTask tsk in tareas)
                        {
                            monto = tsk.Cost;

                        }

                        DateTime fechaP = Convert.ToDateTime(pubProj.CreatedDate.ToShortDateString());
                        DateTime fechaA = Convert.ToDateTime(ayer.ToShortDateString());
                        //
                        x = 0;

                        if (fechaP >= fechaA)
                        {
                            i++;
                            project_id = pubProj.Id.ToString();
                            string name = pubProj.Name;
                            string description = pubProj.Description;
                            DateTime start_date = pubProj.StartDate;
                            DateTime end_date = pubProj.FinishDate;
                           
                            string grouper = "";
                            string compromise = "";
                            string institution = "";
                            string action_line = "";
                            string responsable = "";
                            string ubicacion = "";

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

                                    if (j == 1) { grouper = luEntry.First().FullValue; }
                                    else if (j == 3) { compromise = luEntry.First().FullValue; }
                                    else if (j == 6) { institution = luEntry.First().FullValue; }
                                    else if (j == 7) { action_line = luEntry.First().FullValue; }
                                    else if (j == 9) { responsable = luEntry.First().FullValue; }
                                    else if (j == 10) { ubicacion = luEntry.First().FullValue; }

                                }

                            }

                            insertProject(project_id, name, description, grouper, compromise, start_date, end_date, institution, action_line, responsable, monto, ubicacion);
                            //(id ,deleted_at ,created_at ,updated_at ,project_id ,name ,description ,grouper ,compromise ,amount ,progress ,risk ,amount2 ,progress2 ,risk_description ,start_date ,end_date ,institution ,action_line ,responsable ,responsable_phone ,responsable_email ,monto ,ubicacion)
                        }
                    }
                    if (x == 0) { escribir_log("El rango de dias del parametro 'rango_insert' = " + dias_i + "   no encontro regitros para las transacciones de esa fecha ", "modifique el parametro a menos dias para el proceso de insercion "); }
                }
            }

            //  Console.Write("\nPress any key to exit: ");
            //  Console.ReadKey(false);
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
                    string sql = " Update projects "; //no se borra o elinmina sino que se actualiza el campo 
                    sql += " WHERE gui =" + gui; // se debe hace un barrido de todos los proyectos de Project Server
                    MySqlCommand cmd = new MySqlCommand(sql, connect); // para asegurarse de desactivar correctamente
                    cmd.ExecuteNonQuery(); // por el tema de que los procesos son por rango de fecha
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
