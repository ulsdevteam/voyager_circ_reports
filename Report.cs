using System;
using System.Data;
using System.Text;
using System.Globalization;
using Dapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Oracle.ManagedDataAccess.Client;

namespace voyager_circ_reports
{
    public class Report 
    {
        public Location Location { get; set; }
        public DateTime ReportDate { get; set; }
        public int Month { get; set; }
        public int Year { get; set; }
        public int ActiveChargeTransactions { get; set; }
        public int ArchivedChargeTransactions { get; set; }
        public int TotalChargeTransactions => ActiveChargeTransactions + ArchivedChargeTransactions;
        public int ActiveRenewTransactions { get; set; }
        public int ArchivedRenewTransactions { get; set; }
        public int TotalRenewTransactions => ActiveRenewTransactions + ArchivedRenewTransactions;
        public int TotalChargeAndRenewTransactions => TotalChargeTransactions + TotalRenewTransactions;
        public int TotalDischargeTransactions { get; set; }
        public string Title => 
            $"{(Location is null ? "" : Location.DisplayName + ' ')}Monthly Circ Stats for {new DateTimeFormatInfo().GetMonthName(Month)}, {Year}";

        public static Report Run(Options options)
        {
            using var connection = new OracleConnection(options.ConnectionString);                    
            connection.Open();                    
            return new Report
            {
                ReportDate = DateTime.Now,
                Month = options.Month,
                Year = options.Year,
                Location = GetLocation(connection, options.Location),
                ActiveChargeTransactions = Query(connection, options, "circ_transactions", "charge"),
                ArchivedChargeTransactions = Query(connection, options, "circ_trans_archive", "charge"),
                ActiveRenewTransactions = Query(connection, options, "renew_transactions", "renew"),
                ArchivedRenewTransactions = Query(connection, options, "renew_trans_archive", "renew"),
                TotalDischargeTransactions = Query(connection, options, "circ_trans_archive", "discharge"),
            };
        }

        public void ExportToExcel(string path)
        {
            using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet {
                Id = workbookPart.GetIdOfPart(worksheetPart), 
                SheetId = 1, 
                Name = "Sheet1"
            };
            sheets.Append(sheet);
            sheetData.AppendChild(new Row()).AppendChild(new Cell {
                DataType = CellValues.String,
                CellValue = new CellValue(ReportDate.ToString("MM/dd/yyyy H:mm:ss"))
            });
            sheetData.AppendChild(new Row()).AppendChild(new Cell {
                DataType = CellValues.String,
                CellValue = new CellValue(Title)
            });
            sheetData.AppendChild(new Row());
            if (Location != null) {
                AddDataRow(CellValues.String, "Circ Happening Location Code", Location.Code);
                AddDataRow(CellValues.String, "Circ Happening Location Name", Location.Name);
            }
            AddDataRow(CellValues.Number, "Active Charge Transactions", ActiveChargeTransactions);
            AddDataRow(CellValues.Number, "Archived Charge Transactions", ArchivedChargeTransactions);
            AddDataRow(CellValues.Number, "Total Charge Transactions", TotalChargeTransactions);
            AddDataRow(CellValues.Number, "Active Renew Transactions", ActiveRenewTransactions);
            AddDataRow(CellValues.Number, "Archived Renew Transactions", ArchivedRenewTransactions);
            AddDataRow(CellValues.Number, "Total Renew Transactions", TotalRenewTransactions);
            AddDataRow(CellValues.Number, "Total Charge + Renew Transactions", TotalChargeAndRenewTransactions);
            AddDataRow(CellValues.Number, "Total Discharge Transactions", TotalDischargeTransactions);

            void AddDataRow(CellValues dataType, string name, object data)
            {
                var row = new Row();
                row.AppendChild(new Cell {
                    DataType = CellValues.String,
                    CellValue = new CellValue(name)
                });
                row.AppendChild(new Cell {
                    DataType = dataType,
                    CellValue = new CellValue(data.ToString())
                });
                sheetData.AppendChild(row);
            }
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.AppendLine(ReportDate.ToString("MM/dd/yyyy H:mm:ss"));
            sb.AppendLine(Title);
            sb.AppendLine();
            if (Location != null) {
                sb.AppendLine($"Circ Happening Location Code\t{Location.Code}");
                sb.AppendLine($"Circ Happening Location Name\t{Location.Name}");
            }
            sb.AppendLine($"Active Charge Transactions\t{ActiveChargeTransactions}");
            sb.AppendLine($"Archived Charge Transactions\t{ArchivedChargeTransactions}");
            sb.AppendLine($"Total Charge Transactions\t{TotalChargeTransactions}");
            sb.AppendLine($"Active Renew Transactions\t{ActiveRenewTransactions}");
            sb.AppendLine($"Archived Renew Transactions\t{ArchivedRenewTransactions}");
            sb.AppendLine($"Total Renew Transactions\t{TotalRenewTransactions}");
            sb.AppendLine($"Total Charge + Renew Transactions\t{TotalChargeAndRenewTransactions}");
            sb.AppendLine($"Total Discharge Transactions\t{TotalDischargeTransactions}");
            return sb.ToString();
        }
        
        static Location GetLocation(IDbConnection connection, string locationCode)
        {
            if (locationCode is null) return null;
            const string sql = @"
                select location_code as Code, location_name as Name, location_display_name as DisplayName
                from pittdb.location where location_code = :locationCode";
            return connection.QuerySingle<Location>(sql, new { locationCode });
        }

        static int Query(IDbConnection connection, Options options, string table, string chargeOrRenew)
        {
            var sql = new StringBuilder($"select count(*) from pittdb.{table} t ");
            if (options.Location is null) {
                sql.Append("where ");
            } else {
                sql.Append($"join pittdb.location l on t.{chargeOrRenew}_location = l.location_id where l.location_code = :location and ");
            }
            sql.Append($"trunc(t.{chargeOrRenew}_date) between to_date(:firstDay, 'yyyy/mm/dd') and to_date(:lastDay, 'yyyy/mm/dd')");            
            var firstDay = $"{options.Year}/{options.Month}/1";
            var lastDay = $"{options.Year}/{options.Month}/{DateTime.DaysInMonth(options.Year, options.Month)}";
            return connection.QuerySingle<int>(sql.ToString(), new { firstDay, lastDay, location = options.Location });
        }        
    }
}
