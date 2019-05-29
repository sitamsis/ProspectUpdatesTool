using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;

namespace WTFiserv.Crm.PU.Application
{
    public class CommonObject
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
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
                log.Info("Error occured in the method GetMetadata" + ex.Message);
                var stList = ex.StackTrace.ToString().Split('\\');
                log.Info("Exception occurred at " + stList[stList.Count() - 1]);
                log.Info("Exception details:" + ex.ToString());
                return null;
            }
        }

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
                else if(!fileName.Contains("ProspectUpdatesExecutionON") && !fileName.Contains("CampaignUpdateExecutionOn"))
                {
                    fileList.Add(fileName);
                    log.Info(fileName + " added to the list of files to process");
                }
            }

            return fileList;
        }


        public static void ProcessFile(String path, OrganizationServiceProxy service, string campaignName = "")
        {
            try
            {
                List<AttributeMetadata> AccountMetadata;

                log.Info("Reading file...");
                log.Info("Attempting to read file from: " + path);
                char[] delimiters = new char[] { '\t' };

                using (StreamReader csvreader = new StreamReader(path))
                {
                    var line = csvreader.ReadLine().ToLower();
                    if (line != null)
                    {
                        List<String> values = line.Split(delimiters, StringSplitOptions.None).ToList();
                        List<String> headers = values;

                        AccountMetadata = GetMetadata("account", service, headers);

                        int idindx = headers.IndexOf("accountid");
                        if(idindx < 0)
                        {
                            throw new Exception("Column accountid could not found or wrongly spelled in the file " + path);
                        }
                        while (!csvreader.EndOfStream)
                        {
                            try
                            {
                                line = csvreader.ReadLine();
                                values = line.Split(delimiters, StringSplitOptions.None).ToList();

                                //store metadata on hashtable
                                Dictionary<string, AttributeMetadata> metadatadic = new Dictionary<string, AttributeMetadata>();
                                foreach (var am in AccountMetadata)
                                {
                                    metadatadic.Add(am.LogicalName.ToLower(), am);
                                }
                                var columns = headers.Zip(values, (h, v) => new { header = h, value = v });
                                Entity account = new Entity();
                                account.LogicalName = "account";
                                account.Id = new Guid(values[idindx]);
                                log.Info("Processing Client ID: " + values[idindx]);

                                foreach (var c in columns)
                                {
                                    AttributeMetadata am = null;
                                    metadatadic.TryGetValue(c.header, out am);
                                    if (am != null)
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
                                                    case "IN":
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
                                                    case "OUT":
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
                                                    string fieldStrValue = c.value;
                                                    if (fieldStrValue.Substring(0, 1) == "\"")
                                                        fieldStrValue = fieldStrValue.Substring(1);
                                                    if (fieldStrValue.Substring(fieldStrValue.Length - 1, 1) == "\"")
                                                        fieldStrValue = fieldStrValue.Substring(0, fieldStrValue.Length - 1);
                                                    account[am.LogicalName.ToLower()] = fieldStrValue;
                                                    break;
                                            }
                                        }
                                    }
                                }

                                //**********************************************************************************************
                                //Following line is to update the campaign name into Client/account entity for Campaign update only
                                //**********************************************************************************************
                                if (campaignName != "")
                                {
                                    account["wtfiserv_campaignname"] = campaignName;
                                }
                                //**********************************************************************************************

                                //update entity with new values.
                                service.Update(account);
                                //execute record reassignment.
                                string ownerEntity = "systemuser";
                                Guid ownerId = new Guid();
                                if (headers.Contains("ownerid"))
                                {
                                    if (headers.Contains("ownertype"))
                                    {
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
                                throw ex;
                            }
                        }
                    }
                    else
                    {
                        log.Info("Couldn't read line from the guids file");
                    }
                    csvreader.Close();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error:" + ex.ToString());
                throw ex;
            }
        }

        public class Campaign
        {
            public String CampaignName { get; set; }
            public Guid CampaignId { get; set; }
            public DateTime CreatedOn { get; set; }
            public Guid OwnerId { get; set; }
            public Guid CreatedBy { get; set; }
            public bool IsNew { get; set; }
            public bool hasActualStart { get; set; } 
        }

        public static void MoveFilesToArchive(string srcDir, string destDir,string process="")
        {
            try
            {
                string[] files = Directory.GetFiles(srcDir);
                string fileName, destFileName;
                foreach (string file in files)
                {
                    if (process != "Campaign")
                    {
                        if (!file.Contains("ProspectUpdatesExecutionON"))
                        {
                            fileName = string.Concat(Path.GetFileNameWithoutExtension(file), DateTime.Now.ToString("yyyyMMddHHmmssfff"), Path.GetExtension(file));
                            destFileName = Path.Combine(destDir, fileName);
                            File.Move(file, destFileName);
                        }
                    }
                    else
                    {
                        if (!file.Contains("CampaignUpdateExecutionOn"))
                        {
                            fileName = string.Concat(Path.GetFileNameWithoutExtension(file), Path.GetExtension(file));
                            destFileName = Path.Combine(destDir, fileName);
                            File.Move(file, destFileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Error in MoveFilesToArchive:" + ex.ToString());
                throw ex;
            }
        }

        public static void DeleteArchiveFiles(string dirName)
        {

            try
            {
                string[] files = Directory.GetFiles(dirName);

                foreach (string file in files)
                {
                    FileInfo fi = new FileInfo(file);
                    if (fi.CreationTime < DateTime.Now.AddMonths(-1))
                        fi.Delete();
                }
            }
            catch (Exception ex)
            {
                log.Error("Error in DeleteArchiveFiles:" + ex.ToString());
                throw ex;
            }
        }

        public static void SendEmailMessage(string emailAddress, string body,string emailSubject)
        {
            try
            {
                using (MailMessage mailMessage = new MailMessage())
                {
                    mailMessage.From = new MailAddress(ConfigurationManager.AppSettings["fromemail"].ToString());
                    mailMessage.Subject = emailSubject;
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
