using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OracleClient;
using System.Linq;
using System.Text;

namespace CNO.BPA.GenerateXLS.DataHandler
{
    public class DataAccess : IDisposable
    {
        #region Variables

        //private CustomParameters _parmsCustom = null;
        private Framework.Cryptography crypto = new Framework.Cryptography();
        private OracleConnection _connection = null;
        private string _connectionString = null;
        private OracleTransaction _transaction = null;
        private string _DSN = string.Empty;
        private string _DBUser = string.Empty;
        private string _DBPass = string.Empty;

        #endregion

        #region Constructors
        public DataAccess()
        {
            //check to see that we have values for the db info
            if (BatchDetail.DSN.Length != 0 && BatchDetail.DB_User.Length != 0 &&
                BatchDetail.DB_Pass.Length != 0)
            {
                _DSN = BatchDetail.DSN;
                _DBUser = crypto.Decrypt(BatchDetail.DB_User);
                _DBPass = crypto.Decrypt(BatchDetail.DB_Pass);
                //build the connection string
                _connectionString = "Data Source=" + _DSN + ";Persist Security Info=True;User ID="
                   + _DBUser + ";Password=" + _DBPass + "";
            }
            else
            {
                throw new ArgumentNullException("-266007825; Database connection information was "
                   + "not found.");
            }
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Connects and logs in to the database, and begins a transaction.
        /// </summary>
        public void Connect()
        {
            _connection = new OracleConnection();
            _connection.ConnectionString = _connectionString;
            try
            {
                _connection.Open();
                _transaction = _connection.BeginTransaction();
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred while connecting to the database.", ex);
            }
        }
        /// <summary>
        /// Commits the current transaction and disconnects from the database.
        /// </summary>
        public void Disconnect()
        {
            try
            {
                if (null != _connection)
                {
                    _transaction.Commit();
                    _connection.Close();
                    _connection = null;
                    _transaction = null;
                }
            }
            catch { } // ignore an error here
        }
        /// <summary>
        /// Commits all of the data changes to the database.
        /// </summary>
        internal void Commit()
        {
            _transaction.Commit();
        }
        /// <summary>
        /// Cancels the transaction and voids any changes to the database.
        /// </summary>
        public void Cancel()
        {
            _transaction.Rollback();
            _connection.Close();
            _connection = null;
            _transaction = null;
        }
        /// <summary>
        /// Generates the command object and associates it with the current transaction object
        /// </summary>
        /// <param name="commandText"></param>
        /// <param name="commandType"></param>
        /// <returns></returns>
        internal OracleCommand GenerateCommand(string commandText, System.Data.CommandType commandType)
        {
            OracleCommand cmd = new OracleCommand(commandText, _connection);
            cmd.Transaction = _transaction;
            cmd.CommandType = commandType;
            return cmd;
        }
        #endregion

        #region Public Methods
        public string getBatchNo()
        {
            try
            {
                string batchNo = string.Empty;
                Connect();
                OracleCommand cmd = GenerateCommand("bpa_apps.pkg_batch.get_batch_no", CommandType.StoredProcedure);
                DBUtilities.CreateAndAddParameter("p_in_batch_source_code",
                   BatchDetail.ScannerID, OracleType.VarChar, ParameterDirection.Input, cmd);
                DBUtilities.CreateAndAddParameter("p_out_batch_no",
                  OracleType.VarChar, ParameterDirection.Output, 15, cmd);
                DBUtilities.CreateAndAddParameter("p_out_result",
                   OracleType.VarChar, ParameterDirection.Output, 255, cmd);
                DBUtilities.CreateAndAddParameter("p_out_error_message",
                   OracleType.VarChar, ParameterDirection.Output, 4000, cmd);

                cmd.ExecuteNonQuery();

                if (cmd.Parameters["p_out_result"].Value.ToString().ToUpper() == "SUCCESSFUL")
                {
                    //grab the values
                    batchNo = cmd.Parameters["p_out_batch_no"]
                       .Value.ToString();
                }
                else
                {
                    throw new Exception("-266088529; Procedure Error: " +
                       cmd.Parameters["p_out_result"].Value.ToString() + "; Oracle Error: " +
                       cmd.Parameters["p_out_error_message"].Value.ToString());
                }
                Disconnect();
                return batchNo;
            }
            catch (Exception ex)
            {
                throw new Exception("DataHandler.DataAccess.getBatchNo: " + ex.Message);
            }
        }
      
        public DataSet getDepartmentDetails()
        {
            try
            {
                DataSet DataSetResults = new DataSet();
                Connect();
                OracleCommand cmd = GenerateCommand("bpa_apps.pkg_ia.select_department", CommandType.StoredProcedure);
                DBUtilities.CreateAndAddParameter("p_in_department_name",
                  BatchDetail.Department, OracleType.VarChar, ParameterDirection.Input, cmd);
                DBUtilities.CreateAndAddParameter("p_out_ref_cursor",
                   DBNull.Value, OracleType.Cursor, ParameterDirection.Output,
                   cmd);
                DBUtilities.CreateAndAddParameter("p_out_result",
                   OracleType.VarChar, ParameterDirection.Output, 255, cmd);
                DBUtilities.CreateAndAddParameter("p_out_error_message",
                   OracleType.VarChar, ParameterDirection.Output, 4000, cmd);

                using (OracleDataReader dataReader = cmd.ExecuteReader())
                {
                    if (cmd.Parameters["p_out_result"].Value.ToString()
                       .ToUpper() != "SUCCESSFUL")
                    {
                        throw new Exception("-266088529; Procedure Error: " +
                           cmd.Parameters["p_out_result"].Value.ToString() +
                           "; Oracle Error: " + cmd.Parameters[
                           "p_out_error_message"].Value.ToString());
                    }
                    else
                    {
                        if (dataReader.HasRows)
                        {
                            DataTable dt = new DataTable("Results");
                            DataSetResults.Tables.Add(dt);
                            DataSetResults.Load(dataReader, LoadOption.PreserveChanges, DataSetResults.Tables[0]);
                            Disconnect();
                            return DataSetResults;
                        }
                        else
                        {
                            Disconnect();
                            return null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("DataHandler.DataAccess.getDepartmentDetails: " + ex.Message);
            }
        }

        /// <summary>
        /// 1. Get values from IA_XLS_DEFINITION table.
        /// 2. Build a dynamic input string based on the ref cursor fetched above.
        /// 3. Query values from BATCH_ITEM table.
        /// </summary>
        /// <returns></returns>
        public DataSet getBatchItemDetails()
        {
            try
            {
                DataSet dsXLSDefinition = new DataSet();
                DataSet dsBatchItemDetails = new DataSet();

                Connect();

                OracleCommand cmd = GenerateCommand("bpa_apps.pkg_ia.select_ia_xls_definition", CommandType.StoredProcedure);

                DBUtilities.CreateAndAddParameter("p_in_department_name", BatchDetail.Department, OracleType.VarChar, ParameterDirection.Input, cmd);                
                DBUtilities.CreateAndAddParameter("p_out_result", OracleType.VarChar, ParameterDirection.Output, 255, cmd);
                DBUtilities.CreateAndAddParameter("p_out_error_message", OracleType.VarChar, ParameterDirection.Output, 4000, cmd);
                DBUtilities.CreateAndAddParameter("p_out_ref_cursor", DBNull.Value, OracleType.Cursor, ParameterDirection.Output, cmd);

                using (OracleDataReader dataReader = cmd.ExecuteReader())
                {
                    if (cmd.Parameters["p_out_result"].Value.ToString().ToUpper() != "SUCCESSFUL")
                    {
                        throw new Exception("-266088529; Procedure Error: " +
                           cmd.Parameters["p_out_result"].Value.ToString() +
                           "; Oracle Error: " + cmd.Parameters[
                           "p_out_error_message"].Value.ToString());
                    }
                    else
                    {
                        if (dataReader.HasRows)
                        {
                            DataTable dt = new DataTable("Results");
                            dsXLSDefinition.Tables.Add(dt);

                            //Get XLS definition
                            dsXLSDefinition.Load(dataReader, LoadOption.PreserveChanges, dsXLSDefinition.Tables[0]);

                            //Build input params string from the table fetched above
                            string strInputParams = string.Empty;

                            for (int i = 0; i < dsXLSDefinition.Tables[0].Rows.Count; i++)
                            {
                                if (i == dsXLSDefinition.Tables[0].Rows.Count - 1)
                                {
                                    strInputParams = strInputParams + dsXLSDefinition.Tables[0].Rows[i][2].ToString() + " AS \"" + dsXLSDefinition.Tables[0].Rows[i][1].ToString() + "\"";
                                }
                                else if (i == 6)
                                {
                                    //
                                    strInputParams = strInputParams + "TO_CHAR(" + dsXLSDefinition.Tables[0].Rows[i][2].ToString() + ",\'MM/DD/YYYY\')" + " AS \"" + dsXLSDefinition.Tables[0].Rows[i][1].ToString() + "\", ";
                                }
                                else
                                {
                                    strInputParams = strInputParams + dsXLSDefinition.Tables[0].Rows[i][2].ToString() + " AS \"" + dsXLSDefinition.Tables[0].Rows[i][1].ToString() + "\", ";
                                }                                
                            }

                            string strSelectQuery;

                            strSelectQuery = "select " + strInputParams + " from bpa_apps.batch_item where batch_no = '" + BatchDetail.BatchNo + "'";

                            OracleCommand cmd1 = GenerateCommand(strSelectQuery, CommandType.Text);

                            OracleDataAdapter oda = new OracleDataAdapter(cmd1);

                            DataTable dt1 = new DataTable("Results");

                            oda.Fill(dt1);

                            dsBatchItemDetails.Tables.Add(dt1);

                            return dsBatchItemDetails;                            
                        }
                        else
                        {
                            Disconnect();
                            return null;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("DataHandler.DataAccess.getBatchItemDetails: " + ex.Message);
            }
        }
           
        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            crypto = null;            
            _connection = null;
            _connectionString = null;
            _transaction = null;
            _DSN = string.Empty;
            _DBUser = string.Empty;
            _DBPass = string.Empty;
        }
        #endregion
    }
}

