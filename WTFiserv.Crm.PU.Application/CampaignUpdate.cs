using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using WTFiserv.Crm.PU.DataAccess;
using static WTFiserv.Crm.PU.Application.CommonObject;

namespace WTFiserv.Crm.PU.Application
{
    class CampaignUpdate
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static OrganizationServiceProxy _service;
        static string toemails;
        static string directory_path;
        static string archive_path;
        static string emailSubject;
        List<Campaign> CampaignList { get; set; }
        public void CampaignUpdatefunc()
        {
            directory_path = Convert.ToString(ConfigurationManager.AppSettings["directory_path_campaigns"]);
            archive_path = Convert.ToString(ConfigurationManager.AppSettings["archive_path_campaigns"]);
            toemails = Convert.ToString(ConfigurationManager.AppSettings["toemails_campaigns"]);
            emailSubject = ConfigurationManager.AppSettings["emailsubject_campaigns"].ToString();
            Connection cc = new Connection();
            _service = cc.ConnectToCRM(log);
            string Campaigns = string.Empty;
            try
            {
                if (_service != null)
                {
                    log.Info("Connection succeeded!");

                    List<String> fileList = ProcessDirectory(directory_path);
                    if (fileList.Count() > 0)
                    {
                        foreach (String fileName in fileList)
                        {
                            try
                            {
                                log.Info("Loading Clients from Campaign file" + fileName);
                                String campaignName = Path.GetFileNameWithoutExtension(fileName);
                                Campaigns += campaignName + " ";
                                if (campaignName.IndexOf("_")>0)
                                {
                                    campaignName = campaignName.Substring(0, campaignName.IndexOf("_"));
                                }
                                String line;
                                char[] delimiters = new char[] { '\t' };
                                StreamReader file = new StreamReader(fileName);
                                Campaign campaign = GetExistingCampaign(campaignName);
                                line = file.ReadLine();
                                String[] headers = line.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

                                ProcessCampaign(campaign);
                                ProcessFile(fileName, _service, campaignName);
                            }
                            catch (Exception ex)
                            {
                                log.Error(ex.ToString());
                            }
                            log.Info("Finishing processing file " + fileName);
                        }
                    }
                    else
                    {
                        log.Info("No files to process were found in " + directory_path);
                        SendEmailMessage(toemails, "Error while updating Campaigns. No files to process were found in " + directory_path,emailSubject);
                        return;
                    }
                }
                else
                {
                    log.Error("Couldn't connect to CRM...");
                    SendEmailMessage(toemails, "Error while updating Campaigns. Couldn't connect to CRM!",emailSubject);
                }
                log.Info("Execution finished.");

                //SendEmailMessage(toemails, "Campaigns for the files " + Campaigns + " is updated!",emailSubject);
                MoveFilesToArchive(directory_path, archive_path,"Campaign");
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

        private void ProcessCampaign(Campaign campaign)
        {
            log.Info("Processing Campaign" + campaign.CampaignName);
            if (campaign.IsNew)
            {
                Entity entity = new Entity("campaign");
                entity["campaignid"] = campaign.CampaignId;
                entity["name"] = campaign.CampaignName;
                entity["actualstart"] = DateTime.Today; //Solution 25
                entity["createdon"] = campaign.CreatedOn;
                //entity["ownerid"] = new EntityReference("systemuser", campaign.OwnerId);
                //entity["createdby"] = new EntityReference("systemuser", campaign.CreatedBy);
                entity["statecode"] = new OptionSetValue(0);
                entity["statuscode"] = new OptionSetValue(0);
                _service.Create(entity);
                campaign.IsNew = false;
            }
            else
            {//Solution 25
                if (!campaign.hasActualStart)
                {
                    Entity entity = new Entity("campaign");
                    entity["campaignid"] = campaign.CampaignId;
                    entity["actualstart"] = DateTime.Today; //Solution 25
                    _service.Update(entity);
                }
            }
        }

        private Campaign GetExistingCampaign(String campaignName)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = "campaign";
            //query.ColumnSet = new ColumnSet(new string[] { "name","actualstart" });
            query.ColumnSet = new ColumnSet("name", "actualstart");
            query.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "name",
                Operator = ConditionOperator.Equal,
                Values = { campaignName }
            });
            EntityCollection entities = _service.RetrieveMultiple(query);
            if (entities.Entities.Count > 0)
                return new Campaign
                {
                    CampaignId = entities.Entities[0].Id,
                    CampaignName = campaignName,
                    IsNew = false,
                    hasActualStart = entities.Entities[0].Attributes.Contains("actualstart"),
                };
            else
                return new Campaign
                {
                    CampaignId = Guid.NewGuid(),
                    CampaignName = campaignName,
                    //CreatedBy = UserId,
                    //OwnerId = UserId,
                    CreatedOn = DateTime.Now,
                    IsNew = true,
                };
        }
    }
}