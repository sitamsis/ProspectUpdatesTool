using System;
using System.Linq;
using System.Configuration;

namespace WTFiserv.Crm.PU.Application
{
    public class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static string process;
        static void Main(string[] args)
        {
            try
            {
                process = Convert.ToString(ConfigurationManager.AppSettings["process"]);
                if (process == "Campaigns")
                {
                    CampaignUpdate campaignsUpdate = new CampaignUpdate();
                    campaignsUpdate.CampaignUpdatefunc();
                }
                else
                {
                    CRAUpdate craUpdate = new CRAUpdate();
                    craUpdate.CRAUpdatefunc();
                }
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
