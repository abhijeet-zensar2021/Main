using ClosedXML.Excel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System.Collections.Generic;
using System.Data;
//using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace CalendarRule
{
    class Program
    {
        public static void Main(string[] args)
        {
            var connectionString = "Url=https://ztmsdev.crm8.dynamics.com/; Username=Harshal@ztmsd365.onmicrosoft.com; Password=1qaz@WSX3edc; AuthType=Office365";
            var crmSvcClient = new CrmServiceClient(connectionString);

            //fetch Calendar Rule Records
            System.Data.DataTable Dt = getCalendarRuleRecord(crmSvcClient);

            //fetch Calendar Records
            EntityCollection calendarRecords = getCalendarRecords(crmSvcClient);

            //System.Data.DataTable Dt5 = convert(calendarRecords);

            // create excel file for Calendar Rule
            bool flagCalendarRule = convertToExcelCalendarRule(Dt);

            //create excel file for Calendar
            bool flagCalendar = convertToExcelCalendar(calendarRecords);

        }

        private static System.Data.DataTable convert(EntityCollection calendarRecords)
        {
            EntityCollection accountRecords = calendarRecords;
            System.Data.DataTable dTable = new System.Data.DataTable();
            int iElement = 0;

           

            //Defining the ColumnName for the datatable
            for (iElement = 0; iElement <= accountRecords.Entities[0].Attributes.Count - 1; iElement++)
            {
                string columnName = accountRecords.Entities[0].Attributes.Keys.ElementAt(iElement);
                dTable.Columns.Add(columnName);
            }

            foreach (Entity entity in accountRecords.Entities)
            {
                DataRow dRow = dTable.NewRow();
                for (int i = 0; i <= entity.Attributes.Count - 1; i++)
                {
                    string colName = entity.Attributes.Keys.ElementAt(i);
                    dRow[colName] = entity.Attributes.Values.ElementAt(i);
                }
                dTable.Rows.Add(dRow);
            }
            return dTable;
        }

        private static System.Data.DataTable getCalendarRuleRecord(CrmServiceClient crmSvcClient)
        {
            var fetch = @"<fetch>
                          <entity name='calendarrule' >
                            <attribute name='calendarid' />
                            <attribute name='description' />
                            <attribute name='name' />
                            <attribute name='starttime' />
                          </entity>
                        </fetch>";

            var executeFetchReq = new ExecuteFetchRequest
            {
                FetchXml = fetch
            };

            //Works
            EntityCollection collection = null;

            var crmSvcExecuteFetchResponse = crmSvcClient.Execute(executeFetchReq);

            string xmlString = crmSvcExecuteFetchResponse["FetchXmlResult"].ToString();

            XDocument doc = new XDocument();
            System.Data.DataTable Dt = new System.Data.DataTable("ssss");

            XmlDocument doc1 = new XmlDocument();
            doc1.Load(new StringReader(xmlString));
            XmlNode NodoEstructura = doc1.FirstChild.FirstChild;
            //  Table structure (columns definition) 

            Dt.Columns.Add("Calendar Rule");
            Dt.Columns.Add("Calendar");
            Dt.Columns.Add("Name");
            Dt.Columns.Add("Description");
            Dt.Columns.Add("Start");

            XmlNode Filas = doc1.FirstChild;
            //  Data Rows 
            foreach (XmlNode Fila in Filas.ChildNodes)
            {
                DataRow dr = Dt.NewRow();
                List<string> Valores = new List<string>();
                foreach (XmlNode Columna in Fila.ChildNodes)
                {
                    if (Columna.Name == "calendarruleid")
                        dr["Calendar Rule"] = Columna.InnerText;
                    if (Columna.Name == "calendarid")
                        dr["Calendar"] = Columna.InnerText;
                    if (Columna.Name == "name")
                        dr["Name"] = Columna.InnerText;
                    if (Columna.Name == "description")
                        dr["Description"] = Columna.InnerText;
                    if (Columna.Name == "starttime")
                        dr["Start"] = Columna.InnerText;
                }
                Dt.Rows.Add(dr);
            }
            return Dt;
           
        }

        private static bool convertToExcelCalendar(EntityCollection calendarRecords)
        {
            var workbook = new XLWorkbook();
            workbook.AddWorksheet("sheetName");

            var ws = workbook.Worksheet("sheetName");
            //Recorrer el objecto
            int row1 = 1;
            int row = 2;

            ws.Cell("A" + row1.ToString()).Value = "calendarid";
            ws.Cell("B" + row1.ToString()).Value = "name";
            ws.Cell("C" + row1.ToString()).Value = "businessunitid";
            ws.Cell("D" + row1.ToString()).Value = "type";
            ws.Cell("E" + row1.ToString()).Value = "holidayschedulecalendarid";

            foreach (Entity entObj in calendarRecords.Entities)
            {
                if (entObj.Attributes.Contains("calendarid"))
                {
                    ws.Cell("A" + row.ToString()).Value = entObj.Attributes["calendarid"].ToString();
                }
                if (entObj.Attributes.Contains("name"))
                {
                    ws.Cell("B" + row.ToString()).Value = entObj.Attributes["name"].ToString();
                }
                if (entObj.Attributes.Contains("businessunitid"))
                {
                    EntityReference entref = (EntityReference)entObj.Attributes["businessunitid"];
                    var LookupId = entref.Id;
                    ws.Cell("C" + row.ToString()).Value = LookupId;
                }
                if (entObj.Attributes.Contains("type"))
                {
                    OptionSetValue optionSet = (OptionSetValue)entObj.Attributes["type"];
                    ws.Cell("D" + row.ToString()).Value = optionSet;
                }
                if (entObj.Attributes.Contains("holidayschedulecalendarid"))
                {
                    ws.Cell("E" + row.ToString()).Value = entObj.Attributes["holidayschedulecalendarid"].ToString();
                }
                row++;
            }

            workbook.SaveAs("calender.xlsx");

            return true;
        }

        private static EntityCollection getCalendarRecords(CrmServiceClient crmSvcClient)
        {
            var fetch = @"<fetch top='50' >
                              <entity name='calendar' >
                                <attribute name='holidayschedulecalendarid' />
                                <attribute name='description' />
                                <attribute name='name' />
                                <attribute name='businessunitid' />
                                <attribute name='calendarid' />
                                <attribute name='businessunitidname' />
                                <attribute name='holidayschedulecalendaridname' />
                                <attribute name='type' />
                                <filter>
                                  <condition attribute='createdon' operator='today' />
                                </filter>
                              </entity>
                            </fetch>";


            EntityCollection collection = null;
            collection = crmSvcClient.RetrieveMultiple(new FetchExpression(fetch));
           
            return collection;
        }

        private static bool convertToExcelCalendarRule(System.Data.DataTable Dt)
        {
           
            var workbook = new XLWorkbook();
            workbook.AddWorksheet("sheetName");

            var ws = workbook.Worksheet("sheetName");
            XLWorkbook wb = new XLWorkbook();

            wb.Worksheets.Add(Dt, "WorksheetName");
            wb.SaveAs("CalendarRule.xlsx");
            return true;
        }

    }

}
