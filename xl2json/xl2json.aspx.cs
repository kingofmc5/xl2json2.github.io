using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.Routing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Excel;
using System.Text.Json;
using System.Text.Json.Serialization;
using Newtonsoft.Json.Serialization;

using System.Web.Script.Serialization;

namespace jsonapp
{
    public partial class xl2json : System.Web.UI.Page
    {
        string seller = "";
        string selleraddress = "";
        string sellergstinuin = "";
        string sellerstatename = "";
        string sellercin = "";
        string selleremail = "";

        string buyer = "";
        string buyeraddress = "";
        string buyergstinuin = "";
        string buyerstatename = "";

        string invoiceno,
            ewaybillno,
            invoicedate,
            deliverynote,
            modetermsofpayment, 
            suppliersref, 
            customerponodate, 
            buyersorderno, 
            orderdate, 
            despatchdocumentno, 
            deliverynotedate, 
            despatchedthrough, 
            destination = "";


        string[][] dog;

        string dogjson,gstjson,hsnsacjson = string.Empty;
        string totalquantity,totalamount,totalamtwords,totaltaxamtwords,companyPAN,declaration
            = string.Empty;



        protected void Page_Load(object sender, EventArgs e)
        {
            //alert("hi");
        }

        [Serializable]
        class SerializableDataTable
        {
            public string TableName { get; set; }
            public string[] ColumnNames { get; set; }
            public object[][] Rows { get; set; }

            
        }

        public string dt2json(System.Data.DataTable dataTable)
        {
            try
            {
                {
                    JavaScriptSerializer jsSerializer = new JavaScriptSerializer();
                    List<Dictionary<string, object>> parentRow = new List<Dictionary<string, object>>();
                    Dictionary<string, object> childRow;
                    foreach (DataRow row in dataTable.Rows)
                    {
                        childRow = new Dictionary<string, object>();
                        foreach (DataColumn col in dataTable.Columns)
                        {
                            childRow.Add(col.ColumnName, row[col]);
                        }
                        parentRow.Add(childRow);
                    }
                    return jsSerializer.Serialize(parentRow);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        protected void alert(string alertmsg) {
            ScriptManager.RegisterStartupScript(this, GetType(), "ShowAlert", "alert('"+alertmsg+"');", true);
        }

        protected void upload_click(object sender, EventArgs e)
        {
            //try {
            //    if (FileUpload1.HasFile)
            //    {
            //        DataTable dataTable = ConvertExcelToDataTable(FileUpload1.PostedFile.InputStream);
            //        // Now you have the data in the 'dataTable' variable
            //        // You can bind it to a GridView or perform other operations
            //    }

            //}
            //catch (Exception ex)
            //{
            //    alert(ex.Message);
            //}
            //return;
            try
            {
                string saveaspath;
                HttpPostedFile postedFile = FileUpload1.PostedFile;
                if (postedFile.FileName.Trim() == "")
                {

                    //ScriptManager.RegisterStartupScript(this.Page, this.Page.GetType(), "alert", "<script type='text/javascript'>alert('Please select a file');</script>", true);
                    ScriptManager.RegisterStartupScript(this, GetType(), "Please select a file", "alertMessage();", true);


                    return;
                }

                string[] FileExt = System.IO.Path.GetFileName(postedFile.FileName).Split('.');
                if (FileExt[FileExt.Length - 1].ToUpper() == "XLSX" || FileExt[FileExt.Length - 1].ToUpper() == "XLS")
                {
                }
                else
                {

                }
                string oldfilename = Path.GetFileNameWithoutExtension(postedFile.FileName);
                string xtension = Path.GetExtension(postedFile.FileName);
                string datetimenamepart = DateTime.Now.ToString("yyyyMMddHHmmssffff");
                string newfilename = oldfilename + "_" + datetimenamepart + xtension;

                string folderpath = System.Configuration.ConfigurationManager.AppSettings["uploads"];
                saveaspath = folderpath + newfilename;
                postedFile.SaveAs(Server.MapPath(saveaspath));

                string excelcontent = ReadCells(saveaspath);

                string finaljson =

                "{\n" +
                    "\"Seller\": \"" + seller + "\",\n" +
                    "\"Buyer\": \"" + buyer + "\",\n" +
                    "\"Invoice No.\": \"" + invoiceno + "\",\n" +
                    "\"e-Way Bill No.\": \"" + ewaybillno + "\",\n" +
                    "\"Invoice Dated\": \"" + invoicedate + "\",\n" +
                    "\"Delivery Note\": \"" + deliverynote + "\",\n" +
                    "\"Mode/Terms of Payment\": \"" + modetermsofpayment + "\",\n" +
                    "\"Supplier's Ref.\": \"" + suppliersref + "\",\n" +
                    "\"Customer PO No & Date\": \"" + customerponodate + "\",\n" +
                    "\"Buyer's Order No.\": \"" + buyersorderno + "\",\n" +
                    "\"Order Date\": \"" + orderdate + "\",\n" +
                    "\"Despatch Document No.\": \"" + despatchdocumentno + "\",\n" +
                    "\"Delivery Note Date\": \"" + deliverynotedate + "\",\n" +
                    "\"Despatched through\": \"" + despatchedthrough + "\",\n" +
                    "\"Destination\": \"" + destination + "\",\n" +
                    "\"Description Of Goods\": " + dogjson + ",\n" +
                    "\"Quantity Of Goods\": \"" + totalquantity + "\",\n" +
                    "\"GST Details\": " + gstjson + ",\n" +
                    "\"Amount Chargeable\": \"" + totalamount + "\",\n" +
                    "\"Amount Chargeable (in words)\": \"" + totalamtwords + "\",\n" +
                    //"\"GST Details\": " + gstjson==""?"{}":gstjson + ",\n" +
                    "\"HSN/SAC Details\": " + hsnsacjson + ",\n" +
                    "\"Total Tax Amount (in words)\": \"" + totaltaxamtwords + "\",\n" +
                    "\"Company's PAN\": \"" + companyPAN + "\",\n" +
                    "\"Declaration\": \"" + declaration + "\"\n" +

                "}";

              

                //string invoiceno, ewaybillno, invoicedate, deliverynote, modetermsofpayment, suppliersref, customerponodate,
                //buyersorderno, orderdate, despatchdocumentno, deliverynotedate, despatchedthrough, destination = "";
                txt_output.Text = finaljson ;


                //DataTable dataTable = ConvertExcelToDataTable(FileUpload1.PostedFile.InputStream);
                //DataTable dataTable = ConvertExcelToDataTable(Server.MapPath(saveaspath));

                //return;
                return;
                System.Data.DataTable dt = ConvertToDataTable(Server.MapPath(saveaspath));
                int i = 0;
                //////////////////////////string[,] excelData = ReadExcelData(Server.MapPath(saveaspath));
                //////////////////////////string text1 = "";
                //////////////////////////int rows = excelData.GetLength(0);
                //////////////////////////int cols = excelData.GetLength(1);

                //////////////////////////for (int i = 0; i < rows; i++)
                //////////////////////////{
                //////////////////////////    for (int j = 0; j < cols; j++)
                //////////////////////////    {
                //////////////////////////        string value = excelData[i, j];
                //////////////////////////        //Console.Write(value + " ");
                //////////////////////////        text1 = text1 + value + " ";
                //////////////////////////    }
                //////////////////////////    //Console.WriteLine();
                //////////////////////////}
                //////////////////////////string jsonfilesrc1 = folderpath + oldfilename + "_" + datetimenamepart + ".txt";
                //////////////////////////using (StreamWriter writer = new StreamWriter(Server.MapPath(jsonfilesrc1)))
                //////////////////////////{
                //////////////////////////    writer.WriteLine(text1);

                //////////////////////////}
                //////////////////////////Response.Clear();
                //////////////////////////Response.ContentType = "application/octet-stream";
                //////////////////////////Response.AppendHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(jsonfilesrc1));
                //////////////////////////Response.TransmitFile(jsonfilesrc1);
                //////////////////////////return;

                //ConvertExcelToJson(saveaspath, folderpath + oldfilename + "_" + datetimenamepart + ".txt");

                ////////////////////////////string text = ConvertExcelToJson(saveaspath);
                ////////////////////////////string jsonfilesrc = folderpath + oldfilename + "_" + datetimenamepart + ".txt";
                ////////////////////////////using (StreamWriter writer = new StreamWriter(Server.MapPath(jsonfilesrc)))
                ////////////////////////////{
                ////////////////////////////    writer.WriteLine(text);

                ////////////////////////////}


                ////////////////////////////////Response.ContentType = "application/octet-stream";
                //////////////////////////////////Response.Flush();

                ////////////////////////////////Response.TransmitFile(Server.MapPath(folderpath + oldfilename + "_" + datetimenamepart + ".txt"));
                ////////////////////////////////Response.Flush();
                //////////////////////////////Response.End();

                //////////////////////////////Response.Clear();
                //////////////////////////////Response.ClearHeaders();

                //////////////////////////////Response.AppendHeader("Content-Length", text.Length.ToString());
                //////////////////////////////Response.ContentType = "text/plain";
                //////////////////////////////Response.AppendHeader("Content-Disposition", "attachment;filename="+ oldfilename + "_" + datetimenamepart + ".txt");

                //////////////////////////////Response.Write(text);
                //////////////////////////////Response.End();
                ////////////////////////////Response.Clear();
                ////////////////////////////Response.ContentType = "application/octet-stream";
                ////////////////////////////Response.AppendHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(jsonfilesrc));
                ////////////////////////////Response.TransmitFile(jsonfilesrc);
                //Response.End();

            
            }
            catch (Exception ex)
            {

                //throw ex;
                alert(ex.Message);

            }

        }

        public string ReadCells(string filePath)
        {
            string jsonstring = "";
            string excelsheetstring = "";
            try
            {


                Excel.Application excelApp = null;
                Excel.Workbook workbook = null;
                Excel.Worksheet worksheet = null;

                try
                {
                    excelApp = new Excel.Application();
                    //workbook = excelApp.Workbooks.Open(filePath);
                    workbook = excelApp.Workbooks.Open(Server.MapPath(filePath));
                    worksheet = (Excel.Worksheet)workbook.Worksheets[1]; 

                    //next 2 lines : attempting to unmerge cells
                    //Excel.Range mergedCells = worksheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeConstants, Excel.XlSpecialCellsValue.xlTextValues);
                    //mergedCells.UnMerge();

                    // Read specific cell values
                    Excel.Range cell = worksheet.Cells[1, 1]; 
                    string cellValue = cell.Value2.ToString();
                    //Console.WriteLine("Cell A1 Value: " + cellValue);

                    // Read entire range of cells
                    Excel.Range usedRange = worksheet.UsedRange;
                    object[,] values = usedRange.Value2 as object[,];

                    System.Data.DataTable dt = new System.Data.DataTable();

                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; col++)
                    {
                        // Assuming you want to name the columns as Column1, Column2, Column3, etc.
                        dt.Columns.Add("Column" + col);
                    }
                    // Loop through rows
                    for (int row = 1; row <= worksheet.UsedRange.Rows.Count; row++)
                    {
                        DataRow dataRow = dt.NewRow();

                        // Loop through columns
                        for (int col = 1; col <= worksheet.UsedRange.Columns.Count; col++)
                        {
                            // Get the cell value
                            var cellval = (worksheet.Cells[row, col] as Range)?.Value2;

                            //if(col - 1 >= 0)
                            dataRow[col - 1] = cellval != null ? cellval.ToString() : string.Empty;
                        }

                        dt.Rows.Add(dataRow);
                    }
                    //return "";

                    System.Data.DataTable dtdog = new System.Data.DataTable();
                    dtdog.Columns.Add("SlNo.");
                    dtdog.Columns.Add("ID");
                    dtdog.Columns.Add("Description");
                    dtdog.Columns.Add("HSN/SAC");
                    dtdog.Columns.Add("Quantity");
                    dtdog.Columns.Add("Rate");
                    dtdog.Columns.Add("Discount");
                    dtdog.Columns.Add("Amount");

                    System.Data.DataTable dthsnsac = new System.Data.DataTable();
                    dthsnsac.Columns.Add("HSN/SAC");
                    dthsnsac.Columns.Add("Taxable Value");
                    dthsnsac.Columns.Add("Integrated Tax Rate");
                    dthsnsac.Columns.Add("Amount");
                    dthsnsac.Columns.Add("Total Tax Amount");

                    System.Data.DataTable dtgst = new System.Data.DataTable();

                    int rowCount = values.GetLength(0);
                    int colCount = values.GetLength(1);

                    bool concatenate = false;

                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            if (values[row, col] == null)
                                values[row, col] = "";
                        }
                    }

                        for (int row = 1; row <= rowCount; row++)
                        {
                            for (int col = 1; col <= colCount; col++)
                            {
                                object cellContent = values[row, col];
                                //Console.Write(cellContent.ToString() + "\t");
                                //Console.Write(cellContent!=null?(cellContent.ToString() + "\t"):("blankcell\t"));

                                excelsheetstring += (cellContent != null ? (cellContent.ToString() + "\t") : ("\t"));


                                //jsonstring += "";

                                if (cellContent != null && cellContent.ToString() == "Orpac Systems Private Limited")
                                {
                                    concatenate = true;
                                    for (int r = row; concatenate; r++)
                                    {
                                        if (values[r, col].ToString() == "Buyer") {
                                            concatenate = false;
                                            break;
                                        }

                                        seller += values[r, col].ToString().Trim() + (r == row?", ":" ");
                                        jsonstring += values[r, col].ToString().Trim() + " ";//useless line

                                    }
                                }

                                if (cellContent != null && cellContent.ToString().Trim() == "Buyer")
                                {
                                    concatenate = true;
                                    for (int r = row+1; concatenate; r++)
                                    {
                                        if (values[r, col].ToString().Trim() == "Sl")
                                        {
                                            concatenate = false;
                                            break;
                                        }

                                        buyer += (values[r, col].ToString().Trim() == "GSTIN/UIN:" || values[r, col].ToString().Trim() == "State Name :" || values[r, col].ToString().Trim() == "PAN/IT No :")
                                            ? (values[r, col].ToString().Trim() + " " + values[r, col + 3].ToString().Trim() + " ")
                                            : ((values[r, col].ToString().Trim() + " "));


                                        jsonstring += values[r, col].ToString().Trim() + " ";//useless line

                                    }
                                }

                                if (cellContent != null && cellContent.ToString().Trim() == "Invoice No.") 
                                {
                                    invoiceno = values[row+1, col].ToString().Trim();                                
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "e-Way Bill No.")
                                {
                                    ewaybillno = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Dated" && row==3)
                                {
                                    invoicedate = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Delivery Note")
                                {
                                    deliverynote = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Mode/Terms of Payment")
                                {
                                    modetermsofpayment = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Supplier's Ref.")
                                {
                                    suppliersref = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Customer PO No & Date")
                                {
                                    customerponodate = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Buyer's Order No.")
                                {
                                    buyersorderno = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Dated" && row==9)
                                {
                                    orderdate = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Despatch Document No.")
                                {
                                    despatchdocumentno = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Delivery Note Date")
                                {
                                    deliverynotedate = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Despatched through")
                                {
                                    despatchedthrough = values[row + 1, col].ToString().Trim();
                                }
                                if (cellContent != null && cellContent.ToString().Trim() == "Destination")
                                {
                                    destination = values[row + 1, col].ToString().Trim();
                                }

                                if (cellContent != null && cellContent.ToString() == "Description of Goods") {
                                    //alert(values[row + 2, col-1].ToString().Trim());
                                    int i = row + 2;//ID of good no.1 starts 2 rows after the title Description of goods
                                    int j = 0;
                                    int k = 0;
                                    concatenate = true;
                                    while (concatenate)
                                    {
                                        string x = values[i, 1].ToString();
                                        if (int.TryParse(x, out int anyint))
                                        {
                                            DataRow dr = dtdog.NewRow();
                                            dr["SlNo."] = anyint;
                                            dr["ID"] = values[i, 2].ToString();
                                            dr["HSN/SAC"] = values[i, 9].ToString();
                                            dr["Quantity"] = values[i, 10].ToString();
                                            dr["Rate"] = values[i, 11].ToString();
                                            dr["Discount"] = values[i, 13].ToString();
                                            dr["Amount"] = values[i, 14].ToString();

                                            dtdog.Rows.Add(dr);
                                        }
                                        //else if (values[i + 1, 3].ToString() == "IGST")
                                        else if (values[i + 2, 2].ToString() == "Total")
                                        {
                                            dogjson = dt2json(dtdog);
                                            concatenate = false;
                                            break;
                                        }
                                        else {
                                            string desc = dtdog.Rows[dtdog.Rows.Count - 1]["Description"].ToString();
                                            desc += " " + values[i, 2].ToString();
                                            dtdog.Rows[dtdog.Rows.Count - 1]["Description"] = desc;
                                        }
                                        i++;
                                        //dtdog.Columns.Add("SlNo.");
                                        //dtdog.Columns.Add("ID");
                                        //dtdog.Columns.Add("Description");
                                        //dtdog.Columns.Add("HSN/SAC");
                                        //dtdog.Columns.Add("Quantity");
                                        //dtdog.Columns.Add("Rate");
                                        //dtdog.Columns.Add("Discount");
                                        //dtdog.Columns.Add("Amount");

                                        ////if (values[row + 2, col + 1].ToString().Trim() == "IGST")
                                        ////if (values[i, col + 1].ToString().Trim() == "IGST")
                                        //if (values[i, col + 1].ToString().Trim() == "IGST" || values[i - 1, col + 1].ToString().Trim() == "IGST")
                                        //{
                                        //    concatenate = false;
                                        //    break;
                                        //}
                                        //else if (values[i, col - 1] != "" || values[i, col - 1] != null)
                                        //{
                                        //    dog[j][0] = values[i, col].ToString().Trim();
                                        //    dog[j][1] = values[i, col + 7].ToString().Trim();
                                        //    dog[j][2] = values[i, col + 8].ToString().Trim();
                                        //    dog[j][3] = values[i, col + 9].ToString().Trim();
                                        //    dog[j][4] = values[i, col + 11].ToString().Trim();
                                        //    dog[j][5] = values[i, col + 12].ToString().Trim();

                                        //    k = i+1;
                                        //    //for (k = i+1; (values[k, col - 1] == null); k++)
                                        //    while (values[k, col - 1] == ""|| values[k, col - 1] == null) 
                                        //    {
                                        //        dog[j][5] += values[k, col].ToString();
                                        //        if (values[k+2, col + 1].ToString().Trim() == "IGST")
                                        //        {
                                        //            concatenate = false;
                                        //            break;
                                        //        }
                                        //                                                }
                                        //    j++;


                                        //}
                                        //else i++;

                                    }
                                }
                            if (col ==2 & values[row, 2].ToString() == "Total")
                            {
                                totalquantity = values[row, 10].ToString();
                                totalamount =  values[row, 14].ToString();
                            }

                            if (values[row, col].ToString() == "Amount Chargeable (in words)")
                            { 
                                totalamtwords = values[row+1, col].ToString();
                            }
                            
                            if (values[row, col].ToString() == "Tax Amount (in words)  :")
                            { 
                                totaltaxamtwords = values[row, 7].ToString();
                            }
                            
                            if (values[row, col].ToString() == "Company's PAN :")
                            { 
                                companyPAN = values[row, 4].ToString();
                            }
                            if (values[row, col].ToString() == "Declaration")
                            { 
                                declaration = values[row+1, col].ToString();
                            }

                            if (col ==1 & values[row, 1].ToString() == "HSN/SAC")
                            {
                                //for (int i = row + 2; values[i + 1, col].ToString() != "Total"; i++)
                                for (int i = row + 2; values[i, col].ToString() != "Total"; i++)
                                {
                                    DataRow dr = dthsnsac.NewRow();
                                    dr["HSN/SAC"] = values[i, 1].ToString();
                                    dr["Taxable Value"] = values[i, 11].ToString();
                                    dr["Integrated Tax Rate"] = values[i, 12].ToString();
                                    dr["Amount"] = values[i, 13].ToString();
                                    dr["Total Tax Amount"] = values[i, 14].ToString();
                                    dthsnsac.Rows.Add(dr);
                                }


                            }
                            if (col == 1 & values[row, 1].ToString() == "Total")
                            {
                                DataRow dr = dthsnsac.NewRow();
                                dr["HSN/SAC"] = values[row, 1].ToString();
                                dr["Taxable Value"] = values[row, 11].ToString();
                                dr["Integrated Tax Rate"] = values[row, 12].ToString();
                                dr["Amount"] = values[row, 13].ToString();
                                dr["Total Tax Amount"] = values[row, 14].ToString();
                                dthsnsac.Rows.Add(dr);

                                hsnsacjson = dt2json(dthsnsac);

                            }

                            if (col == 3 & values[row, 3].ToString().Trim() == "IGST")
                            {
                                dtgst.Columns.Add("IGST");
                                if (dtgst.Rows.Count == 0)
                                {
                                    
                                    DataRow dr = dtgst.NewRow();
                                    dr["IGST"] = values[row, 14].ToString();
                                    dtgst.Rows.Add(dr);
                                }
                                else {
                                    dtgst.Rows[0]["IGST"] = values[row, 14].ToString();
                                }
                            }

                            if (col == 3 & values[row, 3].ToString().Trim() == "CGST")
                            {
                                dtgst.Columns.Add("CGST");
                                if (dtgst.Rows.Count == 0)
                                {

                                    DataRow dr = dtgst.NewRow();
                                    dr["CGST"] = values[row, 14].ToString();
                                    dtgst.Rows.Add(dr);
                                }
                                else
                                {
                                    dtgst.Rows[0]["CGST"] = values[row, 14].ToString();
                                }
                            }
                            if (col == 3 & values[row, 3].ToString().Trim() == "SGST")
                            {
                                dtgst.Columns.Add("SGST");
                                if (dtgst.Rows.Count == 0)
                                {

                                    DataRow dr = dtgst.NewRow();
                                    dr["SGST"] = values[row, 14].ToString();
                                    dtgst.Rows.Add(dr);
                                }
                                else
                                {
                                    dtgst.Rows[0]["SGST"] = values[row, 14].ToString();
                                }
                            }

                            gstjson = dt2json(dtgst);
                            //dthsnsac.Columns.Add("HSN/SAC");

                        }
                        //Console.WriteLine();
                        excelsheetstring += "\n";
                    }
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("Error: " + ex.Message);
                    alert("Error: " + ex.Message);
                }
                finally
                {
                    if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                    if (workbook != null) workbook.Close(false);
                    if (excelApp != null) excelApp.Quit();

                    Marshal.ReleaseComObject(excelApp);
                }
                return excelsheetstring;
                //return jsonstring;
            }
            catch (Exception ex) { alert(ex.Message); return excelsheetstring; }
        }

        private System.Data.DataTable ConvertExcelToDataTable(string fileName)
        {

            System.Data.DataTable dataTable = new System.Data.DataTable();

            try
            {


                using (OleDbConnection connection = new OleDbConnection())
                {
                    connection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+(Server.MapPath(fileName))+";Extended Properties=Excel 12.0 Xml;";
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = "SELECT * FROM [Sheet1$]";
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            adapter.Fill(dataTable);
                        }
                    }
                }
            }
            catch (Exception ex) { alert(ex.Message); }
            return dataTable;
        }

        //public string ConvertExcelToJson(string filePath, string destinationFile)
        public string ConvertExcelToJson(string filePath)
        {

            try
            {
                var data = new JArray();

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(Server.MapPath(filePath), false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        var item = new JObject();
                        int columnIndex = 0;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string header = GetCellValue(workbookPart, worksheetPart, cell);
                            string value = GetCellValue(workbookPart, worksheetPart, row.Elements<Cell>().ElementAt(columnIndex));

                            item[header] = value;
                            columnIndex++;
                        }
                        data.Add(item);
                    }
                }

                //Response.Clear();
                //Response.ContentType = "application/json";
                //Response.AppendHeader("Content-Disposition", "attachment; filename=" + fileDestination);
                //Response.TransmitFile(fileDestination);
                //Response.End();

                return JsonConvert.SerializeObject(data, Formatting.Indented);
            }
            catch (Exception ex)
            {
                //throw ex;
                alert(ex.Message);
                return "";
            }
        }

        

        public System.Data.DataTable ConvertToDataTable(string filePath)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    bool isFirstRow = true;
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        DataRow dataRow = dataTable.NewRow();

                        if (isFirstRow)
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string columnName = GetCellValue(workbookPart, cell);
                                dataTable.Columns.Add(columnName);
                            }
                            isFirstRow = false;
                        }
                        else
                        {
                            int columnIndex = 0;
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                dataRow[columnIndex] = GetCellValue(workbookPart, cell);
                                columnIndex++;
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }

            }
            catch (Exception ex) {
                alert(ex.Message);
            }
                return dataTable;
        }

        private string GetCellValue(WorkbookPart workbookPart, WorksheetPart worksheetPart, Cell cell)
        {
            try
            {
                if (cell.DataType == null)
                {
                    return "nalla";
                }
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int ssid = int.Parse(cell.CellValue.Text);
                    SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(ssid);
                    return ssi.Text.Text;
                }
                else
                {
                    return cell.CellValue.Text;
                }
            }
            catch (Exception ex)
            {
                //throw ex;
                alert(ex.Message);
                return "";
            }
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            SharedStringTablePart sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (cell.DataType == null)
            {
                return "nalla";
            }
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringPart != null)
            {
                int ssid = int.Parse(cell.CellValue.Text);
                return sharedStringPart.SharedStringTable.ElementAt(ssid).InnerText;
            }
            else
            {
                return cell.CellValue.Text;
            }
            return cell.CellValue.Text.ToString();
        }

        public string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.DataType == null)
            {
                return "nalla";
            }
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int ssid = int.Parse(cell.CellValue.Text);
                SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(ssid);
                return ssi.Text.Text;
            }
            else
            {
                return cell.CellValue.Text;
            }
        }

        public int GetColumnCount(SheetData sheetData)
        {
            int maxColumnCount = 0;
            foreach (Row row in sheetData.Elements<Row>())
            {
                int currentColumnCount = row.Elements<Cell>().Count();
                maxColumnCount = Math.Max(maxColumnCount, currentColumnCount);
            }
            return maxColumnCount;
        }

        //public string[,] ReadExcelData(string filePath, string sheetName)
        //{
        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        //    {
        //        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        //        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

        //        if (sheet != null)
        //        {
        //            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        //            Worksheet worksheet = worksheetPart.Worksheet;

        //            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

        //            int rowCount = sheetData.Count();
        //            int columnCount = GetColumnCount(sheetData);

        //            string[,] excelData = new string[rowCount, columnCount];

        //            int rowIndex = 0;
        //            foreach (Row row in sheetData.Elements<Row>())
        //            {
        //                int columnIndex = 0;
        //                foreach (Cell cell in row.Elements<Cell>())
        //                {
        //                    string cellValue = GetCellValue(cell, workbookPart);
        //                    excelData[rowIndex, columnIndex] = cellValue;
        //                    columnIndex++;
        //                }
        //                rowIndex++;
        //            }

        //            return excelData;
        //        }
        //    }

        //    return null;
        //}

        public string[,] ReadExcelData(string filePath)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();

                if (sheet != null)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    DocumentFormat.OpenXml.Spreadsheet.Worksheet worksheet = worksheetPart.Worksheet;

                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    int rowCount = sheetData.Count();
                    int columnCount = GetColumnCount(sheetData);

                    string[,] excelData = new string[rowCount, columnCount];

                    int rowIndex = 0;
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        int columnIndex = 0;
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string cellValue = GetCellValue(cell, workbookPart);
                            excelData[rowIndex, columnIndex] = cellValue;
                            columnIndex++;
                        }
                        rowIndex++;
                    }

                    return excelData;
                }
            }

            return null;
        }




    }
}