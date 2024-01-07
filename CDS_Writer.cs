using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace CDS_Writer
{
    class Program {
        
        public static Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

        public static string filePath = @"CDS_file.xlsx";
        public static Workbook wb;
        public static Worksheet ws;
       
        public static void ExcelFile(int sheet, string cell, string value)
        {
                ws = wb.Worksheets[sheet];

                Range cellRange = ws.Range[cell];
                cellRange.Value = value;
        }

        public static string SqlConnection(string filename)
        {
            using (SqlConnection conn = new SqlConnection(@"Data Source=sql_instance; Initial Catalog=dbase1; Persist Security Info=True; Integrated Security=SSPI;"))
            {
                conn.Open();
                Console.WriteLine("Connection Opened Sucessful");
                Console.WriteLine("\nDone!");

                FileInfo file = new FileInfo(filename);
                string script = file.OpenText().ReadToEnd();

                SqlCommand command = new SqlCommand(script, conn);

                command.Parameters.Add(new SqlParameter("@year", "1900"));

                using (SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        
                        filename = reader["VALUE"].ToString();
                     
                        
                    }

                    return filename;
                }
            }
        }

        public static void killExcel()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.ProcessName.Equals("EXCEL"))
                {
                    PK.Kill();

                  
                }
            }
        }
        static void Main(string[] args)
        {
            string FiletoRead = @"master_file.txt";
            using (StreamReader Readerobject = new StreamReader(FiletoRead))
            {
                String Line;

                wb = excel.Workbooks.Open(Filename: filePath, ReadOnly: false);

                while ((Line = Readerobject.ReadLine()) != null)
                {
                    
                    string[] values = Line.Split(',');

                   Console.WriteLine(values[0].ToString());
                     Console.WriteLine(values[1].ToString());
                     Console.WriteLine(values[2].ToString());
                     Console.WriteLine(values[3].ToString());

                    ExcelFile(Convert.ToInt16(values[0]), values[1], SqlConnection(values[2])); 
                }
               
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            
            wb.Save();
            wb.Close();
            Marshal.FinalReleaseComObject(wb);
            
            excel.Quit();
            excel.Visible = false;
            Marshal.FinalReleaseComObject(excel);
            
            Console.ReadLine();
            Console.WriteLine("Done!");

            /* killExcel(); */








        }
    }
}
