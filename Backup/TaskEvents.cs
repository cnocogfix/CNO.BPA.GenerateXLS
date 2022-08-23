using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Emc.InputAccel.QuickModule.ClientScriptingInterface;
using Emc.InputAccel.ScriptEngine.Scripting;
using log4net;
using System.IO;
using System.Reflection;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace CNO.BPA.GenerateXLS
{    
    public class TaskEvents : ITaskEvents
    {        
        private log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType); 
             
        /// <summary>
        /// 1. Use the Standard_MDF to pull out required values from MDF (DB Credentials and DefaultXLSPath).
        /// 2. DbScript DB credentials in DataAccess.cs
        /// 3. Connect to database and get values from IA_XLS_DEFINITION table.
        /// 4. Build a dynamic query based on the ref cursor fetched above.
        /// 5. Query values from BATCH_ITEM table.
        /// 6. Generate XLS and place in a location.
        /// </summary>
        /// <param name="taskInfo"></param>
        public void ExecuteTask(ITaskInformation taskInfo)
        {            
            try
            {
                //Initialize the logger
                FileInfo fi = new FileInfo(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "CNO.BPA.GenerateXLS.dll.config"));
                                
                log4net.Config.XmlConfigurator.Configure(fi);       
            
                log.Info("Beginning the ExecuteTask method");
                
                //Get database credentials
                getDBCredentials(taskInfo);

                log.Info("DB Credentials: User, Pwd: " + BatchDetail.DB_User + ", " + BatchDetail.DB_Pass);

                //DataSet to hold Batch Item details
                DataSet dsBatchItemDetails = new DataSet();

                DataHandler.DataAccess da = new DataHandler.DataAccess();

                //Get Batch Item details from database
                dsBatchItemDetails = da.getBatchItemDetails();

                //Generate XLS file
                generateXLSFile(dsBatchItemDetails);

                log.Debug("Finished generating XLS file and is placed at the default XLS path in the server");
                log.Info("Completed the ExecuteTask method");
            }
            catch (Exception ex)
            {
                log.Error("Error within the ExecuteTask method: " + ex.Message, ex);
                throw ex;
            }
        }
        private void getDBCredentials(ITaskInformation taskInfo)
        {
            try
            {
                log.Debug("Preparing to loop through each envelope within the batch to get the database credentials");

                string DSN = string.Empty;
                string DB_User = string.Empty;
                string DB_Pass = string.Empty;

                foreach (IWorkflowStep wfStep in taskInfo.Task.Batch.WorkflowSteps)
                {
                    if (wfStep.Name.ToUpper() == "STANDARD_MDF")
                    {
                        Framework.Cryptography crypto = new Framework.Cryptography();

                        BatchDetail.DSN = taskInfo.Task.Batch.Tree.Values(wfStep).GetString("DSN", "");

                        BatchDetail.DB_User = taskInfo.Task.Batch.Tree.Values(wfStep).GetString("DB_USER", "");
                        BatchDetail.DB_Pass = taskInfo.Task.Batch.Tree.Values(wfStep).GetString("DB_PASS", "");

                        BatchDetail.BatchNo = taskInfo.Task.Batch.Tree.Values(wfStep).GetString("BATCH_NO", "");
                        BatchDetail.Department = taskInfo.Task.Batch.Tree.Values(wfStep).GetString("BATCH_DEPARTMENT", "");
                        BatchDetail.DefaultXLSPath = taskInfo.Task.Batch.Tree.Values(wfStep).GetString("DEFAULT_XLS_PATH", "");
                        
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                log.Error("Error within the getDBCredentials method: " + ex.Message, ex);
                throw ex;
            }
        }

        private void generateXLSFile(DataSet dsBatchItemDetails)
        {
            try
            {
                XLSBuilder cb = new XLSBuilder();

                cb.createXLS(dsBatchItemDetails);
            }
            catch(Exception ex)
            {
                log.Error("Error within the generateXLSFile method: " + ex.Message, ex);
                throw ex;
            }
        }
    }
}