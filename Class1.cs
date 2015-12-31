using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Reflection;


namespace Avon_cts
{
    public class zip_unzip
    {
        public zip_unzip()

        {
        }

        public string prueba()
        {
            string _vari;
            _vari = "valor33";
            return _vari;

        }

        public void Compress(string directorySource, string directoryDest, string OutputFileName, string ExtensionsToZip)
        {
            DirectoryInfo Di = new DirectoryInfo(directorySource);

            if (!Directory.Exists(directoryDest))
            {
                Directory.CreateDirectory(directoryDest);
            }
            string zipPath = directoryDest + "\\" + OutputFileName + ".zip";
            FileInfo fi = new FileInfo(zipPath);
            if (fi.Exists)
            {
                fi.Delete();
            }

            foreach (FileInfo fileToCompress in Di.GetFiles(ExtensionsToZip))
            {

                using (FileStream originalFileStream = fileToCompress.OpenRead())
                {
                    if ((File.GetAttributes(fileToCompress.FullName) &
                       FileAttributes.Hidden) != FileAttributes.Hidden & fileToCompress.Extension.ToLower() != ".zip")
                    {
                        string newFile = fileToCompress.FullName;

                        using (ZipArchive archive = ZipFile.Open(zipPath, ZipArchiveMode.Update))
                        {
                            archive.CreateEntryFromFile(newFile, fileToCompress.Name);
                        }
                    }
                }
            }
        }

        public void Decompress(string fileToDecompress, string directoryDest)
        {
            DirectoryInfo fi = new DirectoryInfo(fileToDecompress);
            if ((File.GetAttributes(fi.FullName) &
                      FileAttributes.Hidden) != FileAttributes.Hidden & fi.Extension.ToLower().Equals(".zip"))
            {
                if (!Directory.Exists(directoryDest))
                {
                    Directory.CreateDirectory(directoryDest);
                }
                string currentFileName = fi.FullName;

                using (ZipArchive archive = ZipFile.OpenRead(currentFileName))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        //entry.ExtractToFile(Path.Combine(directoryDest, entry.FullName), true);
                        entry.ExtractToFile(entry.FullName);
                        try
                        {

                        }
                        catch
                        {
                            return;
                        }

                    }
                }

            }
        }
        public void Decompress1(string fileToDecompress, string directoryDest)
        {
            DirectoryInfo fi = new DirectoryInfo(fileToDecompress);
            if ((File.GetAttributes(fi.FullName) &
                      FileAttributes.Hidden) != FileAttributes.Hidden & fi.Extension.ToLower().Equals(".zip"))
            {
                if (!Directory.Exists(directoryDest))
                {
                    Directory.CreateDirectory(directoryDest);
                }
                string currentFileName = fi.FullName;

                using (ZipArchive archive = ZipFile.OpenRead(currentFileName))
                {
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        entry.ExtractToFile(Path.Combine(directoryDest, entry.FullName), true);
                        //entry.ExtractToFile(entry.FullName);
                        try
                        {
                        }
                        catch
                        {
                            return;
                        }
                    }
                }

            }
        }

    }
    public class SMTPClient 
        {
        public string mensajit;

            public SMTPClient()
            {
                // parameterless constructor
            }

            public string pruebaM()
            {
                string _vari;
                _vari = "valor6789";
                return _vari;

            }



        //public string EnviarMail( string from, string to,string cc, string subject, string body, string attach, string host,   int port, string exceptionMessage, int exception, string pass, string usua)
        public string EnviarMail(string from, string to, string cc, string subject, string body, string attach, string host, int port, string exceptionMessage, int exception, string pass, string usua)
        {
             
            MailMessage message = new MailMessage();
          
            string mailRecipient = string.Empty;
                string Attachements = string.Empty;

                try
                {
                    #region TASKConfiguration

                    #endregion

                    #region FROM

                    //string from = string.Empty;
                    try
                    {
                        // from = "SMTPClient@avon.com";//CheckAndGetFromContext(config.From).ToString();
                       
                        message.From = new MailAddress(from);
                        exception = 0;
                    }
                    catch (FormatException ex)
                    {
                        exception = -2;
                        exceptionMessage = ex.Message;
                    }

                    #endregion

                    #region TO
                    //string                     
                    // to = "guillermo.paredes@avon.com,guillermo.paredes@avon.com";//"guillermo.paredes@avon.com,roberto.madoery@avon.com,roberto.yaccarino@avon.com";//string.Empty;
                    //to = CheckAndGetFromContext(config.To).ToString();
                   
                    mailRecipient = to;
                    string[] split = to.Split(new Char[] { ';' });
                    foreach (string s in split)
                    {
                        if (s.Trim() != "")
                        { 
                            mailRecipient = s.Trim();
                      
                            message.To.Add(new MailAddress(mailRecipient));

                        }
                    }
                    #endregion

                    #region CC
                    if (!String.IsNullOrEmpty(cc))
                    {
                        mailRecipient = cc;
                        string[] split1 = cc.Split(new Char[] { ';' });
                        foreach (string s in split1)
                        {
                            if (s.Trim() != "")
                            {
                                mailRecipient = s.Trim();

                                message.To.Add(new MailAddress(mailRecipient));

                            }
                        }
                    }
                    //Logger.WriteActivityTrace("Agregando el/los destinatarios en copia al correo");

                     //Logger.WriteActivityTrace("Algun correo de destinatario en copia es erroneo");
                    #endregion

                    #region PRIORITY             

                    message.Priority = MailPriority.High;//config.Priority;

                    #endregion

                    #region SUBJECT              

                    message.Subject = subject; //"Testing SMTP Mail ABC2.NET";//CheckAndGetFromContext(config.Subject).ToString();

                    #endregion

                    #region BODY

                    message.Body = body;//"SMTPClient Test";//CheckAndGetFromContext(config.Body).ToString();

                    #endregion

                    #region ATTACHS

                    if (!String.IsNullOrEmpty(attach))
                    {
                        Attachements = attach;
                        string[] split2 = attach.Split(new Char[] { ';' });
                        foreach (string s in split2)
                        {
                            if (s.Trim() != "")
                            {
                                Attachements = s.Trim();
                               
                                message.Attachments.Add(new Attachment(Attachements));
                            }
                        }

                    }
                    //Logger.WriteActivityTrace("Se cargaron los Attachments correctamente en el correo");
                    #endregion

                    #region SMTPConfiguration
                    // Logger.WriteActivityTrace("Configurando el Cliente SMTP");

                    // Connecting to the server and configuring it
                    SmtpClient client = new SmtpClient();

                    client.Host = host;//"buantnss";//config.SMTPServerName;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(usua, pass);
                  

                    //int 

                    //  port = 25;
                    try
                    {
                        // port = Convert.ToInt32(config.SMTPServerPort);
                        client.Port = port;
                        exception = 0;
                    }
                    catch (FormatException ex)
                    {
                        //Logger.WriteError(string.Format("El valor para el puerto \"{0}\" no es un valor numerico", port.ToString()));
                        //return WorkingUnitStates.Failed; 
                        exception = -2;
                        exceptionMessage = ex.Message;
                    }

                    client.EnableSsl = true;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;

                    try
                    {
                    //Logger.WriteActivityTrace(string.Format("Enviando mail con SMTPServer {0} Puerto {1}",
                    //   client.Host, client.Port));

                    
                    mensajit = "oka";

  

                        client.Send(message);
                        exception = 0;

                        //Logger.WriteActivityTrace("El mail fue enviado satisfactoriamente");
                    }
                    catch (SmtpException ex)
                    {

                        // Logger.WriteError(string.Format("Se produjo un error al enviar el Mail. Exception: {0}", ex.Message));
                        //  result = WorkingUnitStates.Failed;
                        exception = -2;
                        exceptionMessage = ex.Message;
                        mensajit = ex.Message;
                    return mensajit;
                    }
                return mensajit;
                #endregion
            }
                catch (Exception ex)
                {
                
            
                // Logger.WriteError(string.Format("Se produjo la siguiente falla en la tarea {0}. Exception: {1}",
                //    TaskName, ex.Message));
                //  result = WorkingUnitStates.Failed;
                exception = -2;
                    exceptionMessage = ex.Message;
                return mensajit;
            }
                //return result;
            }
        }
}



