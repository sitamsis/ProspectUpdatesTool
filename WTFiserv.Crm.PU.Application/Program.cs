using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;
using System.Configuration;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Threading.Tasks;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using System.Net;
using System.IO;
using WTFiserv.Crm.PU.DataAccess;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using System.Net.Mail;

namespace WTFiserv.Crm.PU.Application
{
    public class Program
    {

        // Declare the service proxy referring the CRUD
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        static void Main(string[] args)
        {

            string toemails = ConfigurationManager.AppSettings["toemails"].ToString();
            string directory_path = ConfigurationManager.AppSettings["directory_path"].ToString();
            string archive_path = ConfigurationManager.AppSettings["archive_path"].ToString();
            log.Info("Attempting connection to CRM...");
            Connection cc = new Connection();
            OrganizationServiceProxy _service = cc.ConnectToCRM(log);
            

            if (_service != null)
            {
                log.Info("Connection succedded");
                //ClientsRepository cr = new ClientsRepository();
               
                if (!File.Exists(directory_path + ((directory_path.EndsWith("\\")) ? "ProspectUpdatesExecutionON.txt" : "\\ProspectUpdatesExecutionON.txt")))
                {
                    //if the flag file doesn't exists, the execution is finished
                    log.Info("ProspectUpdatesExecutionON.txt file not found in directory path.");
                    log.Info("Finishing execution");
                    Environment.Exit(0);
                }

               

                //gets the list of files to retrieve the guids
                List<String> fileList = ProcessDirectory(directory_path);
                if (fileList.Count() > 0)
                {
                    foreach (String file in fileList)
                    {
                        try
                        {
                            ProcessFile(file, _service);
                           
                        }
                        catch (Exception ex)
                        {
                            log.Error(ex.ToString());
                        }
                        log.Info("Finishing processing file " + file);

                        //need to create a file with _OK at the end
                        String newFile = file.Insert(file.Length - 4, "_OK");
                        if (!File.Exists(newFile))
                        {//creates the new file if it doesn't exist
                            try
                            {
                                File.Create(newFile).Dispose();
                                log.Info(newFile + " created");
                            }
                            catch (Exception ex)
                            {
                                log.Error(ex.ToString());
                            }
                        }
                    }
                }
                else
                {
                    log.Info("No files to process were found in " + directory_path);
                    SendEmailMessage(toemails, "Error while updating CRA data. No files to process were found in " + directory_path);
                }
            }
            else
            {
                log.Error("Couldn't connect to CRM...");
                SendEmailMessage(toemails, "Error while updating CRA data. Couldn't connect to CRM!");
            }
            log.Info("Execution finished.");

            SendEmailMessage(toemails, "CRA Data for this week is updated!");
            MoveFilesToArchive(directory_path, archive_path);
            DeleteArchiveFiles(archive_path);
        }


        public static List<AttributeMetadata> GetMetadata(string EntityName, OrganizationServiceProxy service, List<String> headers)
        {
            try
            {
                RetrieveEntityRequest request = new RetrieveEntityRequest()
                {
                    EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                    LogicalName = EntityName

                };
                RetrieveEntityResponse response = service.Execute(request) as RetrieveEntityResponse;
                var retval = response.EntityMetadata.Attributes
                    .Where(
                m => headers.Contains(
                    m.SchemaName.ToString().ToLower()
                    )
                )
                .ToList();
                return retval;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public static void ProcessFile(String path, OrganizationServiceProxy service)
        {
            try
            {
                List<AttributeMetadata> AccountMetadata;

                log.Info("Reading file...");
                log.Info("Attempting to read file from: " + path);
                StreamReader csvreader = new StreamReader(path);
                char[] delimiters = new char[] { '\t' };
                var line = csvreader.ReadLine().ToLower();
                if (line != null)
                {
                    List<String> values = line.Split(delimiters, StringSplitOptions.None).ToList();
                    List<String> headers = values;

                    AccountMetadata = GetMetadata("account", service, headers);
                    
                    int idindx = headers.IndexOf("accountid");
                    while (!csvreader.EndOfStream)
                    {
                        try
                        {
                            line = csvreader.ReadLine();
                            values = line.Split(delimiters, StringSplitOptions.None).ToList();
                           
                            //store metadata on hashtable
                            Dictionary<string, AttributeMetadata> metadatadic = new Dictionary<string, AttributeMetadata>();
                            foreach (var am in AccountMetadata) {
                                metadatadic.Add(am.LogicalName.ToLower(),am);
                            
                            }

                            //--
                            var columns = headers.Zip(values, (h, v) => new { header = h, value = v });
                           
                            
                            Entity account = new Entity();
                            account.LogicalName = "account";
                           account.Id = new Guid(values[idindx]);
                           log.Info("Processing Client ID: " + values[idindx]);

                           //account = service.Retrieve(account.LogicalName, account.Id, new ColumnSet (new string[] { "name", "ownerid" } ));

                           foreach (var c in columns) {

                              
                                AttributeMetadata am = null;
                                metadatadic.TryGetValue(c.header, out am);
                                if (am != null )
                                {
                                    if (am.IsValidForUpdate == true && am.AttributeType != null && c.value != "" && c.value != null)
                                   {
                                       bool boolValueMapping = false;

                                       if (am.AttributeType.ToString() == "Boolean")
                                       {
                                           switch (c.value.ToUpper())
                                           {
                                               case "1":
                                                   boolValueMapping = true;
                                                   break;
                                               case "YES":
                                                   boolValueMapping = true;
                                                   break;
                                               case "TRUE":
                                                   boolValueMapping = true;
                                                   break;
                                               case "0":
                                                   boolValueMapping = false;
                                                   break;
                                               case "NO":
                                                   boolValueMapping = false;
                                                   break;
                                               case "FALSE":
                                                   boolValueMapping = false;
                                                   break;
                                               default:
                                                   break;

                                           }
                                       }

                                       switch (am.AttributeType.ToString())
                                       {
                                              
                                           case "Money":
                                               account[am.LogicalName.ToLower()] = new Money(Decimal.Parse(c.value));
                                               break;
                                           case "Picklist":
                                               account[am.LogicalName.ToLower()] = new OptionSetValue(int.Parse(c.value));
                                               break;
                                           case "Integer":
                                               account[am.LogicalName.ToLower()] = int.Parse(c.value);
                                               break;
                                           case "Boolean":
                                               account[am.LogicalName.ToLower()] = boolValueMapping;
                                               break;
                                           case "DateTime":
                                               account[am.LogicalName.ToLower()] = DateTime.Parse(c.value);
                                               break;
                                           case "Decimal":
                                               account[am.LogicalName.ToLower()] = Decimal.Parse(c.value);
                                               break;
                                           case "Double":
                                               account[am.LogicalName.ToLower()] = Double.Parse(c.value);
                                               break;
                                           case "BigInt":
                                               account[am.LogicalName.ToLower()] = long.Parse(c.value);
                                               break;
                                           default:
                                               account[am.LogicalName.ToLower()] = c.value;
                                               break;
                                       }
                                   }

                                  


                               }
                           
                           }


                            //update entity with new values.
                           service.Update(account);

                           //execute record reassignment.
                           string ownerEntity = "systemuser";
                           Guid ownerId = new Guid();

                           if (headers.Contains("ownerid")) {

                               if (headers.Contains("ownertype")) {

                                   ownerEntity = values[headers.IndexOf("ownertype")].ToLower();
                                   if (ownerEntity == "user") { ownerEntity = "systemuser"; }

                               }
                               ownerId = Guid.Parse(values[headers.IndexOf("ownerid")].ToString());
                               AssignRequest assign = new AssignRequest
                               {


                                   Assignee = new EntityReference(ownerEntity,
                                       ownerId),
                                   Target = new EntityReference(account.LogicalName,
                                       account.Id)
                               };

                               service.Execute(assign);
                           
                           
                           }
                          


                        }
                        catch (Exception ex)
                        {
                            log.Error("Error:" + ex.ToString());
                           
                        }
                    }
                }
                else
                {
                    log.Info("Couldn't read line from the guids file");
                }

            }
            catch (Exception ex)
            {
                log.Error("Error:" + ex.ToString());
               
            }
           
        }

        // Process all files in the directory passed in, recurse on any directories 
        // that are found, and process the files they contain.
        public static List<String> ProcessDirectory(string targetDirectory)
        {
            log.Info("Searching directory for files...");
            List<String> fileList = new List<String>();
            // Process the list of files found in the directory.
            string[] fileEntries = Directory.GetFiles(targetDirectory);
            foreach (string fileName in fileEntries)
            {
                //if the filename contains the correct name
                //ignoring the flag file
                if (fileName.Contains("CRM_Updates") && !fileName.Contains("_OK") && !fileName.Contains("ProspectUpdatesExecutionON"))
                {
                    if (!File.Exists(fileName.Insert(fileName.Count() - 4, "_OK")))
                    {//if the file hastn't been processed
                        fileList.Add(fileName);
                        log.Info(fileName + " added to the list of files to process");
                    }
                }
            }

            //// Recurse into subdirectories of this directory.
            //string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            //foreach (string subdirectory in subdirectoryEntries)
            //{
            //    //adds the other files in the directories of the path
            //    fileList.InsertRange(0, ProcessDirectory(subdirectory));
            //}
            return fileList;
        }

        public static void MoveFilesToArchive(string srcDir, string destDir)
        {
            string[] files = Directory.GetFiles(srcDir);
            string fileName, destFileName;
            foreach (string file in files)
            {
                if (!file.Contains("ProspectUpdatesExecutionON"))
                {
                    fileName = string.Concat(Path.GetFileNameWithoutExtension(file),DateTime.Now.ToString("yyyyMMddHHmmssfff"),Path.GetExtension(file));
                    destFileName = Path.Combine(destDir, fileName);
                    File.Move(file, destFileName);
                }
            }
        }

        public static void DeleteArchiveFiles(string dirName)
        {
            string[] files = Directory.GetFiles(dirName);

            foreach (string file in files)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.CreationTime < DateTime.Now.AddMonths(-1))
                    fi.Delete();
            }
        }

        public static void SendEmailMessage(string emailAddress, string body)
        {
            try
            {
                using (MailMessage mailMessage = new MailMessage())
                {
                    mailMessage.From = new MailAddress(ConfigurationManager.AppSettings["fromemail"].ToString());
                    mailMessage.Subject = ConfigurationManager.AppSettings["emailsubject"].ToString();
                    mailMessage.Body = body;
                    mailMessage.IsBodyHtml = true;
                    //mailMessage.To.Add(new MailAddress(emailAddress));

                    string[] multiEmails = emailAddress.Split(',');
                    foreach (string toEmail in multiEmails)
                    {
                        mailMessage.To.Add(new MailAddress(toEmail)); //adding multiple email addresses
                    }
                    SmtpClient smtp = new SmtpClient();
                    smtp.Send(mailMessage);
                }
            }
            catch (SmtpException exSmtp)
            {
                log.Info("Exception occured in SendEmailMessage method" + exSmtp);
            }
            catch (Exception ex)
            {
                log.Info("Exception occured in SendEmailMessage method" + ex);
            }
        }
    }
}
