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
     // This coded workflow retrieves all valid sheet names from a given Excel file.
    public class ExcelWorkbookSheetNameExt : CodedWorkflow
    {
        [Workflow]
        public List<string> Execute(string excelFilePath)
        {
            // Step 1: Validate input
            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                throw new ArgumentException("Excel file path cannot be null or empty.");
            }

            List<string> sheetNames = new List<string>();
            
            
            // Step 2: Define OLEDB connection string for Excel
            // ===============================================
            // Provider: Microsoft.ACE.OLEDB.12.0 (must be installed)
            // Extended Properties: "Excel 12.0 Xml;HDR=YES"
            // HDR=YES means the first row is treated as headers
            
            
    
            string connStr = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';";
            
             // Step 3: Connect to Excel file and read schema

            using (var connection = new OleDbConnection(connStr))
            {
                connection.Open();
                DataTable schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (schema != null)
                {
                    
                     // Filter only sheet names that end with "$" or "$'"
                    // These represent actual Excel sheets
                    
                    sheetNames = schema.AsEnumerable()
                        .Select(row => row["TABLE_NAME"].ToString())
                        .Where(name => name.EndsWith("$") || name.EndsWith("$'"))
                        .Select(name => name.Trim('\''))
                        .ToList();
                }
            }

            // Step 4: Return list of cleaned sheet names
            return sheetNames;
        }
    }
}