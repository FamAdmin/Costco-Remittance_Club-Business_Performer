using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using CostcoRemittance_ClubBusiness_Performer;
using FasterLookUpApproachesTest;
using UiPath.CodedWorkflows;

namespace Vendor_InvoiceProcessing_Performer.Framework
{
    public class ExcelWorkbookSheetNameExt : CodedWorkflow
    {
        [Workflow]
        public List<string> Execute(string excelFilePath)
        {
            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                throw new ArgumentException("Excel file path cannot be null or empty.");
            }

            List<string> sheetNames = new List<string>();

            string connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';";

            using (var connection = new OleDbConnection(connStr))
            {
                connection.Open();
                DataTable schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (schema != null)
                {
                    sheetNames = schema.AsEnumerable()
                        .Select(row => row["TABLE_NAME"].ToString())
                        .Where(name => name.EndsWith("$") || name.EndsWith("$'"))
                        .Select(name => name.Trim('\''))
                        .ToList();
                }
            }

            return sheetNames;
        }
    }
}