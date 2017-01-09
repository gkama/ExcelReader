using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace XLSXReader
{
    class xlsxReader
    {
        //Local variables
        private string FilePath { get; set; }
        private Application xlApp { get; set; }

        private bool correctFormat = false;

        //Private stored values if functions are first
        private Dictionary<string, List<string>> _AllSheetsData { get; set; }
        private List<string> _SheetDataList { get; set; }

        //Public methods

        //Constructor
        public xlsxReader(string filePath)
        {
            //Check format
            if (Path.GetExtension(filePath) == ".xlsx")
            {
                correctFormat = true;

                FilePath = filePath;
                xlApp = new Application();

                _AllSheetsData = new Dictionary<string, List<string>>();
                _SheetDataList = new List<string>();
            }
        }

        //Get the Excel Data
        //First sheet's index is 1
        public string GetSheetData(int SheetIndex, string SeparatorString = "\r\n", string filePath = "")
        {
            try
            {
                if (correctFormat)
                {
                    //Optional path
                    if (filePath == "")
                        filePath = FilePath;

                    StringBuilder toReturn = new StringBuilder();

                    Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Worksheet xlWorksheet = (Worksheet)xlWorkbook.Sheets[SheetIndex];
                    Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    //Loop through the cells
                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            string tempVal = xlWorksheet.Cells[i, j].Text.ToString();
                            if (tempVal != null)
                                toReturn.Append(tempVal).Append(SeparatorString);
                        }
                    }
                    return toReturn.ToString();
                }
                else
                {
                    throw new Exception("Incorrect file format");
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        //Get sheet data as a list
        public List<string> GetSheetDataList(int SheetIndex, string filePath = "")
        {
            try
            {
                if (correctFormat)
                {
                    //Optional path
                    if (filePath == "")
                        filePath = FilePath;

                    List<string> toReturn = new List<string>();

                    Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    Worksheet xlWorksheet = (Worksheet)xlWorkbook.Sheets[SheetIndex];
                    Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    //Loop through the cells
                    for (int i = 1; i <= rowCount; i++)
                    {
                        for (int j = 1; j <= colCount; j++)
                        {
                            string tempVal = xlWorksheet.Cells[i, j].Text.ToString();
                            if (!string.IsNullOrEmpty(tempVal))
                                toReturn.Add(tempVal);
                        }
                    }
                    this._SheetDataList = toReturn;
                    return toReturn;
                }
                else
                {
                    throw new Exception("Incorrect file format");
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        //Get all sheets' data
        //Sheet name is the Key
        public Dictionary<string, List<string>> GetAllSheetsData(string filePath = "")
        {
            try
            {
                if (correctFormat)
                {
                    //Optional path
                    if (filePath == "")
                        filePath = FilePath;

                    Dictionary<string, List<string>> toReturn = new Dictionary<string, List<string>>();

                    Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
                    for (int i = 1; i < xlWorkbook.Worksheets.Count + 1; i++)
                    {
                        List<string> valList = new List<string>();

                        Worksheet xlWorksheet = (Worksheet)xlWorkbook.Sheets[i];
                        Range xlRange = xlWorksheet.UsedRange;

                        int rowCount = xlRange.Rows.Count;
                        int colCount = xlRange.Columns.Count;

                        //Loop through the cells
                        for (int k = 1; k <= rowCount; k++)
                        {
                            for (int j = 1; j <= colCount; j++)
                            {
                                string tempVal = xlWorksheet.Cells[k, j].Text.ToString();
                                if (!string.IsNullOrEmpty(tempVal))
                                    valList.Add(tempVal);
                            }
                        }
                        toReturn.Add(xlWorksheet.Name.ToString(), valList);
                    }
                    this._AllSheetsData = toReturn;
                    return toReturn;
                }
                else
                {
                    throw new Exception("Incorrect file format");
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        //Get all sheets' data as a string
        public string GetAllSheetsDataAsString(string SeparatorString = "\r\n", string filePath = "")
        {
            try
            {
                if (correctFormat)
                {
                    //Optional path
                    if (filePath == "")
                        filePath = FilePath;

                    //Variables
                    StringBuilder toReturn = new StringBuilder();
                    Dictionary<string, List<string>> sheetsData = new Dictionary<string, List<string>>();

                    //Check if function was already ran
                    if (this._AllSheetsData.Count == 0)
                        sheetsData = GetAllSheetsData(filePath);
                    else
                        sheetsData = this._AllSheetsData;

                    //Loop through all and store them all as a string
                    foreach (KeyValuePair<string, List<string>> value in sheetsData)
                    {
                        foreach (string s in value.Value)
                            toReturn.Append(s).Append(SeparatorString);
                    }
                    return toReturn.ToString();
                }
                else
                {
                    throw new Exception("Incorrect file format");
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        //Get all sheets data as a List<List<string>>
        public List<List<string>> GetAllSheetsDataAsLists(string filePath = "")
        {
            try
            {
                if (correctFormat)
                {
                    //Optional path
                    if (filePath == "")
                        filePath = FilePath;

                    //Variables
                    List<List<string>> toReturn = new List<List<string>>();
                    List<string> tempList = new List<string>();
                    Dictionary<string, List<string>> sheetsData = new Dictionary<string, List<string>>();

                    //Check if function was already ran
                    if (this._AllSheetsData.Count == 0)
                        sheetsData = GetAllSheetsData(filePath);
                    else
                        sheetsData = this._AllSheetsData;

                    //Loop through all and store them all as a string
                    foreach (KeyValuePair<string, List<string>> value in sheetsData)
                    {
                        foreach (string s in value.Value)
                            tempList.Add(s);
                        toReturn.Add(tempList);
                        tempList = new List<string>();
                    }
                    return toReturn;
                }
                else
                {
                    throw new Exception("Incorrect file format");
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
