using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using CostcoRemittance_ClubBusiness_Performer;
using UiPath.CodedWorkflows;

namespace FasterLookUpApproachesTest.Frameworks
{
    public class GeneralLookup : CodedWorkflow
    {
        [Workflow]
        public string Execute(
            List<string> in_LookupValuesList,
            string in_LookupKey,
            List<string> in_LookupSheetsList,
            List<string> in_LookupReturnColumnsList,
            string range
        )
        {
        
            string lookupValueStr = string.Join(", ", in_LookupValuesList
                .Select(val => $"'{val.Trim().Replace("'", "")}'")
                .Where(s => !string.IsNullOrWhiteSpace(s))
            );

            var queryBuilder = new StringBuilder();
            queryBuilder.Append("SELECT ");

            foreach (var col in in_LookupReturnColumnsList)
            {
                queryBuilder.Append($"[{col}], ");
            }

            queryBuilder.Append("SourceSheet\nFROM (\n");

            for (int i = 0; i < in_LookupSheetsList.Count; i++)
            {
                string sheetName = in_LookupSheetsList[i].TrimEnd('$');
                queryBuilder.Append("    SELECT ");
                foreach (var col in in_LookupReturnColumnsList)
                {
                    queryBuilder.Append($"[{col}], ");
                }
                queryBuilder.Append($"'{sheetName}' AS SourceSheet FROM [{sheetName}$" + string.Format("{0}]", range));

                if (i < in_LookupSheetsList.Count - 1)
                    queryBuilder.Append("\n    UNION ALL\n");
                else
                    queryBuilder.Append("\n");
            }

            queryBuilder.Append(") AS CombinedData\n");

            if (!string.IsNullOrWhiteSpace(in_LookupKey) && !string.IsNullOrWhiteSpace(lookupValueStr))
            {
                if(in_LookupValuesList.Count == 1){
                    queryBuilder.Append($"WHERE Format([{in_LookupKey}], '') = {lookupValueStr}");
                }else{
                    queryBuilder.Append($"WHERE Format([{in_LookupKey}], '') IN ({lookupValueStr})");
                }
                
            }

            return queryBuilder.ToString();
        }
    }
}
