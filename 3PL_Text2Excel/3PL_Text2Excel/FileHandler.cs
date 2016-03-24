using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Configuration;

namespace _3PL_Text2Excel
{
    public class FileHandler
    {
        private DataTable snapshotdatatable = null;
        private List<string> processedfilenamelist = null;
        private List<ThirdPartyFile> unprocessedfilenamelist = null;
        private List<string> fieldValueList = new List<string>();
        private string filename = string.Empty;
        private string sourcefolder = string.Empty;
        private string outputfolder = string.Empty;
        
        public FileHandler(string fileName)
        {
            this.filename = fileName;
            this.sourcefolder = ConfigFileUtility.GetValue("SourceFolder");
            this.outputfolder = ConfigFileUtility.GetValue("OutputFolder");
        }

        public FileHandler(string sourceFolder, string outputFolder)
        {
            this.sourcefolder = sourceFolder;
            this.outputfolder = outputFolder;
        }

        public void Process()
        {
            ReadProcessedFileNameList();
            GetUnProcessedFileNameList();
            BuildSnapShotDataTableStructure();
            ReadFileContent();
        }

        private void ReadProcessedFileNameList()
        {
            string textLine = string.Empty;
            string fileName = ConfigFileUtility.GetValue("3PLFileNameListFileName");

            try
            {
                Console.WriteLine(string.Format("[{0}] - Starting to get processed file name from {1} file...", 
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), fileName));
                MiscUtility.LogHistory(string.Format("Starting to get processed file name from {0} file...", fileName));

                using (StreamReader reader = new FileInfo(fileName).OpenText())
                {
                    processedfilenamelist = new List<string>();

                    while ((textLine = reader.ReadLine()) != null)
                    {
                        processedfilenamelist.Add(textLine);
                    }

                    reader.Close();
                }

                Console.WriteLine(string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                MiscUtility.LogHistory("Done!");
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ReadProcessedFileNameList>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void GetUnProcessedFileNameList()
        {
            bool flag = false;
            DirectoryInfo dir = null;
            string fileName = string.Empty;
            string textLine = string.Empty;
            string tempFileName = string.Empty;

            try
            {
                Console.WriteLine(string.Format("[{0}] - Starting to search all unprocessed 3PL Snap files...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                MiscUtility.LogHistory("Starting to search all unprocessed 3PL Snap files...");

                unprocessedfilenamelist = new List<ThirdPartyFile>();
                dir = new DirectoryInfo(this.sourcefolder);

                foreach (FileInfo fi in dir.GetFiles())
                {
                    fileName = fi.Name;
                    flag = false;

                    for (int index = 0; index < processedfilenamelist.Count; index ++)
                    {
                        if (fileName.Equals(processedfilenamelist[index]))
                            flag = true;
                    }

                    if (!flag && fileName.ToUpper().Contains("SNAP"))
                    {
                        ThirdPartyFile thirdPartyFile = new ThirdPartyFile();
                        thirdPartyFile.FileName = fi.Name;
                        thirdPartyFile.FullFillName = fi.FullName;
                        thirdPartyFile.FileExtension = fi.Extension;

                        unprocessedfilenamelist.Add(thirdPartyFile);
                    }
                }

                Console.WriteLine(string.Format("[{0}] - Done! Total of {1} files are found.", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), unprocessedfilenamelist.Count.ToString()));
                MiscUtility.LogHistory(string.Format("Done! Total of {0} files are found.", unprocessedfilenamelist.Count.ToString()));
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <GetUnProcessedFileNameList>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void ReadFileContent()
        {
            string textLine = string.Empty;
            
            try
            {
                foreach (ThirdPartyFile thirdPartyFile in this.unprocessedfilenamelist)
                {
                    Console.WriteLine(string.Format("[{0}] - Starting to read 3PL Snap text file - {1}...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), thirdPartyFile.FullFillName));
                    MiscUtility.LogHistory(string.Format("Starting to read 3PL Snap text file - {0}...", thirdPartyFile.FullFillName));

                    fieldValueList.Clear();

                    using (StreamReader reader = new FileInfo(thirdPartyFile.FullFillName).OpenText())
                    {
                        textLine = reader.ReadLine();   // Read title line
                        //foreach (string tempLine in textLine.Split('|'))
                        //{
                        //}

                        while ((textLine = reader.ReadLine()) != null)
                        {
                            foreach (string tempLine in textLine.Split('|'))
                            {
                                fieldValueList.Add(tempLine.Replace("\"", ""));
                            }
                        }

                        reader.Close();
                    }

                    Console.WriteLine(string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                    MiscUtility.LogHistory("Done!");

                    FillDataTable();
                    ExportToExcelFile(thirdPartyFile.FileName);
                }
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ReadFileContent>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void FillDataTable()
        {
            DataRow dataRow = null;

            snapshotdatatable.Clear();

            for (int index = 0; index < fieldValueList.Count; index += 6)
            {
                dataRow = snapshotdatatable.NewRow();

                dataRow["Date_Time"] = fieldValueList[index].ToString().Trim();
                dataRow["Site_ID"] = fieldValueList[index + 1].ToString().Trim();
                dataRow["Part"] = fieldValueList[index + 2].ToString().Trim();
                dataRow["Qty"] = fieldValueList[index + 3].ToString().Trim();
                dataRow["Service_Tag"] = fieldValueList[index + 4].ToString();
                dataRow["Part_Class"] = fieldValueList[index + 5].ToString();

                snapshotdatatable.Rows.Add(dataRow);
            }
        }

        private void ExportToExcelFile(string filename)
        {
            string thirdFileName = ConfigFileUtility.GetValue("3PLFileNameListFileName");
            string newFileName = string.Format("{0}.csv", filename.Remove(filename.Length - 4, 4));
            string fullFileName = Path.Combine(ConfigFileUtility.GetValue("OutputFolder"), newFileName);

            Console.WriteLine(string.Format("[{0}] - Starting to export data into Excel file - {1}...", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), filename));
            MiscUtility.LogHistory(string.Format("Starting to export data into Excel file - {0}...", filename));

            try
            {
                MiscUtility.ExportDataIntoCSVFile(fullFileName, snapshotdatatable);
                MiscUtility.SaveFile(thirdFileName, filename);

                Console.WriteLine(string.Format("[{0}] - Done!", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")));
                MiscUtility.LogHistory("Done!");
            }
            catch (Exception ex)
            {
                MiscUtility.LogHistory(string.Format("Function name: <ExportToExcelFile>, Source:{0},  Error:{1}", ex.Source, ex.Message));
                throw;
            }
        }

        private void BuildSnapShotDataTableStructure()
        {
            snapshotdatatable = new DataTable();

            DataColumn dc = new DataColumn();
            dc.ColumnName = "Date_Time";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Site_ID";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Part";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Qty";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Service_Tag";
            snapshotdatatable.Columns.Add(dc);

            dc = new DataColumn();
            dc.ColumnName = "Part_Class";
            snapshotdatatable.Columns.Add(dc);

            //return datatable;
        }
    }
}
