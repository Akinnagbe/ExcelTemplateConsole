using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;

namespace ExcelTemplateConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            LoadTemplate();
        }

        static void LoadTemplate()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var template = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelTemplateConsole.Templates.mzSDoCTemp.xlsx");
            using (var excel = new ExcelPackage(template))
            {
                var ws = excel.Workbook.Worksheets["Blank SDoC"];

                ws.Cells["E7"].Value = "123456";
                SetStyles(ws.Cells["E7"]);

                ws.Cells["E9"].Value = "Olamide James";
                SetStyles(ws.Cells["E9"]);

                ws.Cells["E11"].Value = "EkoBalls";
                SetStyles(ws.Cells["E11"]);

                var tt = ws.Cells["E14"];
                tt.Value = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc blandit, nunc in interdum tincidunt, arcu lorem tempor nibh, eget laoreet mi";
               
                //ws.Cells[tt.Start.Row, tt.Start.Column, tt.End.Row, tt.End.Column].Merge = true;
                SetStyles(tt);

                var items = GetItems();
                var itemsCells = ws.Cells["E16"].LoadFromCollection(items, false, OfficeOpenXml.Table.TableStyles.None);
                var rowNumber = itemsCells.Rows;
                // ws.Cells[ws.Dimension.Address].AutoFitColumns();
                var itemsAddress = items.Count + itemsCells.Rows;
                var declarationCellAddress = $"C{itemsAddress}";
                var declarationCell = ws.Cells[declarationCellAddress];
                declarationCell.Value = "The object of declaration above is in conformity with the requirement of the following documents:";
                declarationCell.Style.Font.Bold = true;
                declarationCell.Style.Font.Color.SetColor(Color.Black);
                declarationCell.Style.Locked = true;
                //declarationCell.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                var itemNumber4 = ws.Cells[$"B{declarationCell.Start.Row + 1}"];
                itemNumber4.Value = "4)";


                var titleRowNumber = itemNumber4.Start.Row + 1;
                var staticHeader = ws.Cells[$"D{ titleRowNumber}"];
                staticHeader.Value = "Document No.:";
                staticHeader.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                int staticHeaderRow = staticHeader.Start.Row + 1;

                ws.Cells[$"D{ titleRowNumber}:E{ titleRowNumber}"].Merge = true;

                ws.Cells[$"G{ titleRowNumber}"].Value = "Title";
                ws.Cells[$"G{ titleRowNumber}:P{ titleRowNumber}"].Merge = true;

                ws.Cells[$"R{ titleRowNumber}"].Value = "Edition";
                ws.Cells[$"R{ titleRowNumber}:S{ titleRowNumber}"].Merge = true;

                ws.Cells[$"U{ titleRowNumber}"].Value = "Date of Issue";
                ws.Cells[$"U{ titleRowNumber}:V{ titleRowNumber}"].Merge = true;

                var documentDetails = GetDocumentDetails();
                int index = 0;
                foreach (var doc in documentDetails)
                {
                    //int emptyRow = staticHeaderRow + i
                    //empty row
                    ws.Cells[$"D{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"].Merge = true;

                    ws.Cells[$"G{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"].Merge = true;

                    ws.Cells[$"R{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"].Merge = true;

                    ws.Cells[$"U{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"].Merge = true;

                    index++;

                    #region Content
                   
                    #region DocumentNo.:
                    ws.Cells[$"D{staticHeaderRow + index}"].Value = doc.DocumentNo;
                    ws.Cells[$"D{staticHeaderRow + index}"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"].Merge = true;
                    SetStyles(ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"]);
                    //ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"].Style.Fill.SetBackground(Color.YellowGreen);
                    #endregion

                    #region Tittle
                    ws.Cells[$"G{staticHeaderRow + index}"].Value = doc.Title;
                    ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"].Merge = true;
                    SetStyles(ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"]);
                    //ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"].Style.Fill.SetBackground(Color.YellowGreen);

                    #endregion

                    #region Editor
                    ws.Cells[$"R{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"].Merge = true;
                    SetStyles(ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"]);
                    //ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"].Style.Fill.SetBackground(Color.YellowGreen);
                    #endregion

                    #region Date of Issue

                    ws.Cells[$"U{staticHeaderRow + index}"].Value = doc.DateOfIssue;
                    ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"].Merge = true;
                    SetStyles(ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"]);
                    //ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    //ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    //ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"].Style.Fill.SetBackground(Color.YellowGreen);
                    #endregion

                    #endregion


                    index++;

                    ws.Cells[$"D{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"D{staticHeaderRow + index}:E{staticHeaderRow + index}"].Merge = true;

                    ws.Cells[$"G{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"G{staticHeaderRow + index}:P{staticHeaderRow + index}"].Merge = true;

                    ws.Cells[$"R{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"R{staticHeaderRow + index}:S{staticHeaderRow + index}"].Merge = true;

                    ws.Cells[$"U{staticHeaderRow + index}"].Value = "";
                    ws.Cells[$"U{staticHeaderRow + index}:V{staticHeaderRow + index}"].Merge = true;

                }

               

                excel.SaveAs(new System.IO.FileInfo(@"C:\Users\akinnagbeeo\Documents\Ekoballs\mzdoc.xlsx"));
            }

           
        }

        static void SetStyles(ExcelRange excelRange)
        {
            excelRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            excelRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            excelRange.Style.Fill.SetBackground(Color.YellowGreen);
        }

        static List<Item> GetItems()
        {
            return new List<Item>
            {
                new Item{ItemName ="Item 1"},
                new Item{ItemName ="Item 2"},
                new Item{ItemName ="Item 3"},
                new Item{ItemName ="Item 4"},
                new Item{ItemName ="Item 5"},
                new Item{ItemName ="Item 6"},
                new Item{ItemName ="Item 7"},
                new Item{ItemName ="Item 8"},
                new Item{ItemName ="Item 9"},
                new Item{ItemName ="Item 10"},
                new Item{ItemName ="Item 11"},
                new Item{ItemName ="Item 12"},
                new Item{ItemName ="Item 13"},
                new Item{ItemName ="Item 14"},
                new Item{ItemName ="Item 15"},
                new Item{ItemName ="Item 16"},
                new Item{ItemName ="Item 17"},

            };
        }

        static List<DocumentDetail> GetDocumentDetails()
        {
            return new List<DocumentDetail>
            {
                new DocumentDetail{DocumentNo = 1, DateOfIssue = "5/15/2015", Edition = "", Title = " IMO Guidelines in Resolution MEPC.269(68)"},
                new DocumentDetail{DocumentNo = 2, DateOfIssue = "11/20/2013", Edition = "", Title = "Regulation EU No. 1257/2013"},
                new DocumentDetail{DocumentNo = 3, DateOfIssue = "10/28/2016", Edition = "", Title = "EMSA's Best Practice Guidance on the IHM"},
                new DocumentDetail{DocumentNo = 4, DateOfIssue = "5/19/2009", Edition = "", Title = "SR/CONF/45 The Hong Kong International Convention for the Safe and Environmentally Sound Recycling of Ships"},

            };
        }

    }

    public class Item
    {
        public string ItemName { get; set; }
    }

    public class DocumentDetail
    {
        public int DocumentNo { get; set; }
        public string Title { get; set; }
        public string Edition { get; set; }
        public string DateOfIssue { get; set; }
    }
}
