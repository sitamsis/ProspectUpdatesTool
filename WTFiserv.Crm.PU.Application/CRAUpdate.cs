using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WTFiserv.Crm.PU.DataAccess;
using static WTFiserv.Crm.PU.Application.CommonObject;

namespace WTFiserv.Crm.PU.Application
{
    class CRAUpdate
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static string toemails;
        static string directory_path;
        static string archive_path;
        static string emailSubject;
        static OrganizationServiceProxy _service;
        public void CRAUpdatefunc()
        {
            toemails = Convert.ToString(ConfigurationManager.AppSettings["toemails"]);
            directory_path = Convert.ToString(ConfigurationManager.AppSettings["directory_path"]);
            archive_path = Convert.ToString(ConfigurationManager.AppSettings["archive_path"]);
            emailSubject = ConfigurationManager.AppSettings["emailsubject"].ToString();
            log.Info("Attempting connection to CRM...");
            Connection cc = new Connection();
            _service = cc.ConnectToCRM(log);
            try
            {
                if (_service != null)
                {
                    log.Info("Connection succedded");
                    if (!File.Exists(directory_path + ((directory_path.EndsWith("\\")) ? "ProspectUpdatesExecutionON.txt" : "\\ProspectUpdatesExecutionON.txt")))
                    {
                        //if the flag file doesn't exists, the execution is finished
                        log.Info("ProspectUpdatesExecutionON.txt file not found in directory path.");
                        log.Info("Finishing execution");
                        Environment.Exit(0);
                    }
                    else
                    {
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
                            SendEmailMessage(toemails, "Error while updating CRA data. No files to process were found in " + directory_path,emailSubject);
                            return;
                        }
                    }
                }
                else
                {
                    log.Error("Couldn't connect to CRM...");
                    SendEmailMessage(toemails, "Error while updating CRA data. Couldn't connect to CRM!",emailSubject);
                }
                log.Info("Execution finished.");

                SendEmailMessage(toemails, "CRA Data for this week is updated!",emailSubject);
                MoveFilesToArchive(directory_path, archive_path);
                DeleteArchiveFiles(archive_path);

            }
            catch (Exception ex)
            {
                log.Info("Error occured" + ex.Message);
                var stList = ex.StackTrace.ToString().Split('\\');
                log.Info("Exception occurred at " + stList[stList.Count() - 1]);
                log.Info("Exception details:" + ex.ToString());
            }
        }
    

    }
}
