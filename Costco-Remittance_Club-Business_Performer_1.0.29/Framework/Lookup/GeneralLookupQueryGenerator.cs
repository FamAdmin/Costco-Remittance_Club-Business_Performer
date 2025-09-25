using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using CostcoRemittance_ClubBusiness_Performer;
using UiPath.CodedWorkflows;

namespace FasterLookUpApproachesTest.Frameworks
{
    // This class defines a coded workflow for dynamic data lookup using SQL-style queries.
    public class GeneralLookup : CodedWorkflow
    {
        // The main method that executes the workflow and returns the SQL query as a string
        [Workflow]
        public string Execute(
            List<string> in_LookupValuesList,
            string in_LookupKey,                       // Column name to match against the lookup values
            List<string> in_LookupSheetsList,          // List of Excel sheets to look into
            List<string> in_LookupReturnColumnsList,    // List of column names to retrieve
            string range                                // Excel cell range (e.g., "A1:Z100")           
        )
        {
                        // Step 1: Format the lookup values for SQL IN clause
            string lookupValueStr = string.Join(", ", in_LookupValuesList
                .Select(val => $"'{val.Trim().Replace("'", "")}'")
                .Where(s => !string.IsNullOrWhiteSpace(s))
            );

                        // Step 2: Begin building the SQL query
            var queryBuilder = new StringBuilder();
            queryBuilder.Append("SELECT ");
            
                        // Add each return column to the SELECT clause
            foreach (var col in in_LookupReturnColumnsList)
            {
                queryBuilder.Append($"[{col}], ");
            }

            queryBuilder.Append("SourceSheet\nFROM (\n");
                       // Step 3: Build SELECT statements for each sheet with UNION
            for (int i = 0; i < in_LookupSheetsList.Count; i++)
            {
                string sheetName = in_LookupSheetsList[i].TrimEnd('$');
                queryBuilder.Append("    SELECT ");
                
                // Add all return columns for this sheet
                foreach (var col in in_LookupReturnColumnsList)
                {
                    queryBuilder.Append($"[{col}], ");
                }
                // Add literal value for sheet name as "SourceSheet" column
                queryBuilder.Append($"'{sheetName}' AS SourceSheet FROM [{sheetName}$" + string.Format("{0}]", range));

                // Add UNION ALL between sheet queries, except after the last one
                if (i < in_LookupSheetsList.Count - 1)
                    queryBuilder.Append("\n    UNION ALL\n");
                else
                    queryBuilder.Append("\n");
            }

            queryBuilder.Append(") AS CombinedData\n");
            
            // Step 4: Add WHERE clause based on lookup key and values

            if (!string.IsNullOrWhiteSpace(in_LookupKey) && !string.IsNullOrWhiteSpace(lookupValueStr))
            {
                // If there's only one value, use equality
                
                if(in_LookupValuesList.Count == 1){
                    queryBuilder.Append($"WHERE Format([{in_LookupKey}], '') = {lookupValueStr}");
                }else{
                    queryBuilder.Append($"WHERE Format([{in_LookupKey}], '') IN ({lookupValueStr})");
                }
                
            }
                // Step 5: Return the final query string
            return queryBuilder.ToString();
        }
    }
}
