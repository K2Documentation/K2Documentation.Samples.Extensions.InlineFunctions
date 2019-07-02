using System;
using System.Linq;
using System.Drawing.Design;
using SourceCode.Framework;
using SourceCode.Framework.Design;
using SourceCode.Workflow.Functions;
using SourceCode.Workflow.Functions.Design;
using System.Web.Services.Protocols;
using System.Text.RegularExpressions;

namespace K2Documentation.Samples.Extensions.InlineFunctions
{
    [Category("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "ExcelCategory", typeof(ExcelWebService))]
    public class ExcelWebService
    {
        [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionName", typeof(ExcelWebService)),
        Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionDescription", typeof(ExcelWebService)),
        K2Icon(typeof(ExcelWebService), "Resources.Excel.png")]
        public static string GetCell(
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionParm1Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionParm1Description", typeof(ExcelWebService))]
            string workbookUrl,
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionParm2Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionParm2Description", typeof(ExcelWebService))]
            string sheetName,
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionParm3Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellFunctionParm3Description", typeof(ExcelWebService))]
            string returnCell)
        {
            try
            {
                // get workbook URI to construct the web service URL
                System.Uri workbookUri = new System.Uri(workbookUrl);
                string wsPage = "/_vti_bin/ExcelService.asmx";
                string wsUrl = workbookUri.Scheme + "://" + workbookUri.Host + wsPage;
                // http://portal.denallix.com/_vti_bin/ExcelService.asmx

                Console.WriteLine("K2Documentation.Samples.Extensions.InlineFunctions.ExcelWebService.GetCell - Web Service URL: {0}; Workbook URL: {1}; SheetName: {2}; ReturnCell: {3}",
                    wsUrl, workbookUrl, sheetName, returnCell);

                ExcelServiceReference.ExcelService es = new ExcelServiceReference.ExcelService();
                ExcelServiceReference.Status[] outStatus;
                
                es.Url = wsUrl;
                es.Credentials = System.Net.CredentialCache.DefaultCredentials;

                Regex expColumn = new Regex(@"(?i:)\D+");
                Regex expRow = new Regex(@"(?i:)\d+");

                int returnColumn = GetColumnNumber(expColumn.Match(returnCell).Value) - 1;
                int returnRow = int.Parse(expRow.Match(returnCell).Value) - 1;

                string sessionId = es.OpenWorkbook(workbookUrl, "en-US", "en-US", out outStatus);

                // get cell value
                object rangeResult = es.GetCell(sessionId, sheetName, returnRow, returnColumn, false, out outStatus);
                es.CloseWorkbook(sessionId);

                Console.WriteLine("K2Documentation.Samples.Extensions.InlineFunctions.ExcelWebService.GetCell - ReturnValue: " + rangeResult.ToString());

                return rangeResult.ToString();
            }
            catch (SoapException e)
            {
                Console.WriteLine("K2Documentation.Samples.Extensions.InlineFunctions.ExcelWebService.GetCell: SOAP Exception Message: {0}", e.Message);
                return string.Empty;
            }
        }

        [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionName", typeof(ExcelWebService)),
        Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionDescription", typeof(ExcelWebService)),
        K2Icon(typeof(ExcelWebService), "Resources.Excel.png")]
        public static string GetCellWithInput(
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm1Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm1Description", typeof(ExcelWebService))]
            string workbookUrl,
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm2Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm2Description", typeof(ExcelWebService))]
            string sheetName,
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm3Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm3Description", typeof(ExcelWebService))]
            string inputCell,
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm4Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm4Description", typeof(ExcelWebService))]
            object inputValue,
            [DisplayName("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm5Name", typeof(ExcelWebService)),
            Description("K2Documentation.Samples.Extensions.InlineFunctions.Properties.Resources", "GetCellWithInputFunctionParm5Description", typeof(ExcelWebService))]
            string returnCell)
        {
            try
            {
                // get workbook URI to construct the web service URL
                System.Uri workbookUri = new System.Uri(workbookUrl);
                string wsPage = "/_vti_bin/ExcelService.asmx";
                string wsUrl = workbookUri.Scheme + "://" + workbookUri.Host + wsPage;
                // http://portal.denallix.com/_vti_bin/ExcelService.asmx

                Console.WriteLine("K2Documentation.Samples.Extensions.InlineFunctions.ExcelWebService.GetCell - Web Service URL: {0}; Workbook URL: {1}; SheetName: {2}; InputCell: {3}; InputValue: {4}; ReturnCell: {5}",
                    wsUrl, workbookUrl, sheetName, inputCell, inputValue, returnCell);

                ExcelServiceReference.ExcelService es = new ExcelServiceReference.ExcelService();
                ExcelServiceReference.Status[] outStatus;
                es.Url = wsUrl;
                es.Credentials = System.Net.CredentialCache.DefaultCredentials;

                Regex expColumn = new Regex(@"(?i:)\D+");
                Regex expRow = new Regex(@"(?i:)\d+");

                int inputColumn = GetColumnNumber(expColumn.Match(inputCell).Value) - 1;
                int inputRow = int.Parse(expRow.Match(inputCell).Value) - 1;
                int returnColumn = GetColumnNumber(expColumn.Match(returnCell).Value) - 1;
                int returnRow = int.Parse(expRow.Match(returnCell).Value) - 1;

                string sessionId = es.OpenWorkbook(workbookUrl, "en-US", "en-US", out outStatus);

                // set cell in row, column format
                es.SetCell(sessionId, sheetName, inputRow, inputColumn, inputValue);
                // calc the workbook
                es.Refresh(sessionId, "TempConnectionName");
                // get cell value
                object rangeResult = es.GetCell(sessionId, sheetName, returnRow, returnColumn, false, out outStatus);
                es.CloseWorkbook(sessionId);

                Console.WriteLine("K2Documentation.Samples.Extensions.InlineFunctions.ExcelWebService.GetCell - ReturnValue: " + rangeResult.ToString());

                return rangeResult.ToString();
            }
            catch (SoapException e)
            {
                Console.WriteLine("K2Documentation.Samples.Extensions.InlineFunctions.ExcelWebService.GetCell: SOAP Exception Message: {0}", e.Message);
                return string.Empty;
            }
        }

        // Function to convert an Excel column name to it's column number
        private static int GetColumnNumber(string name)
        {
            return name.ToUpper().Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);
        }
    }
}
