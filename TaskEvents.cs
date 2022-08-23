using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Emc.InputAccel.CaptureClient;
using log4net;
using System.IO;
using System.Reflection;

[assembly: log4net.Config.XmlConfigurator(Watch = true)]
namespace CNO.BPA.GenerateXLS
{    
    public class CodeModule : CustomCodeModule
    {        
        private log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public CodeModule()
        {

        }

        /// <summary>
        /// 1. Use the Standard_MDF to pull out required values from MDF (DB Credentials and DefaultXLSPath).
        /// 2. DbScript DB credentials in DataAccess.cs
        /// 3. Connect to database and get values from IA_XLS_DEFINITION table.
        /// 4. Build a dynamic query based on the ref cursor fetched above.
        /// 5. Query values from BATCH_ITEM table.
        /// 6. Generate XLS and place in a location.
        /// </summary>
        /// <param name="taskInfo"></param>
        /// 
        public override void StartModule(ICodeModuleStartInfo startInfo)
        {
            
        }

        public override void ExecuteTask(IClientTask task, IBatchContext batchContext)
        {            
            try
            {
                //Initialize the logger
                FileInfo fi = new FileInfo(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "CNO.BPA.GenerateXLS.dll.config"));
                                
                log4net.Config.XmlConfigurator.Configure(fi);       
            
                log.Info("Beginning the ExecuteTask method");
                
                //Get database credentials
                getDBCredentials(task, batchContext);

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
                task.CompleteTask();
            }
            catch (Exception ex)
            {
                log.Error("Error within the ExecuteTask method: " + ex.Message, ex);
                task.FailTask(FailTaskReasonCode.GenericRecoverableError, ex);
                throw ex;
            }
        }
        private void getDBCredentials(IClientTask task, IBatchContext batchContext)
        {
            try
            {
                log.Debug("Preparing to loop through each envelope within the batch to get the database credentials");

                string DSN = string.Empty;
                string DB_User = string.Empty;
                string DB_Pass = string.Empty;

                IBatchNode batch = task.BatchNode;

                //foreach (IBatchNode envelope in envelopes)
                //{
                    Framework.Cryptography crypto = new Framework.Cryptography();

                    IBatchNode stnd_MDF = batchContext.GetStepNode(batch, "STANDARD_MDF");

                    BatchDetail.DSN = stnd_MDF.NodeData.ValueSet.ReadString("DSN", "");

                    BatchDetail.DB_User = stnd_MDF.NodeData.ValueSet.ReadString("DB_USER", "");
                    BatchDetail.DB_Pass = stnd_MDF.NodeData.ValueSet.ReadString("DB_PASS", "");

                    BatchDetail.BatchNo = stnd_MDF.NodeData.ValueSet.ReadString("BATCH_NO", "");
                    BatchDetail.Department = stnd_MDF.NodeData.ValueSet.ReadString("BATCH_DEPARTMENT", "");
                    BatchDetail.DefaultXLSPath = stnd_MDF.NodeData.ValueSet.ReadString("DEFAULT_XLS_PATH", "");
                    
                    //break;
                //}
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