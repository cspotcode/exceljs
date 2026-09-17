using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelValidator
{
    internal class OpenErrors
    {
        public string normal { get; set; }
        public string repair { get; set; }
        public string extractData { get; set; }
    }

    internal class Result
    {
        // Which CorruptLoad mode Excel needed to successfully open the file:
        // "normal", "repair", "extractData", or null if every mode failed.
        public string successfulMode { get; set; }

        // Per-mode COM error message, populated for each mode that was attempted and
        // failed
        public OpenErrors errors { get; set; }
    }

    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.Error.WriteLine("Usage: excel-validator.exe <input.xlsx>");
                return 2;
            }

            Result result;
            try
            {
                result = Run(args[0]);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                return 1;
            }

            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions
            {
                WriteIndented = false,
            });
            Console.WriteLine(json);

            return 0;
        }

        // Opens <inputPath> with a 3-step fallback (normal load, then repair mode, then
        // data-extraction mode) and reports which mode succeeded, plus errors for
        // every mode that was attempted and failed.
        private static Result Run(string inputPathArg)
        {
            var result = new Result
            {
                successfulMode = null,
                errors = new OpenErrors(),
            };

            string inputPath;
            try
            {
                inputPath = Path.GetFullPath(inputPathArg);
            }
            catch (Exception ex)
            {
                throw new Exception($"Invalid input path: {ex.Message}", ex);
            }
            if (!File.Exists(inputPath))
            {
                throw new Exception($"Input file not found: {inputPath}");
            }

            Excel.Application excelApp = null;
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;

            try
            {
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                };

                workbooks = excelApp.Workbooks;

                // Internet sources suggest that opening a valid xlsx with
                // CorruptLoad: xlRepairFile may cause an error, so xlRepairFile
                // can't be used unconditionally.
                try
                {
                    workbook = OpenWorkbook(workbooks, inputPath, Excel.XlCorruptLoad.xlNormalLoad);
                    result.successfulMode = "normal";
                }
                catch (COMException normalEx)
                {
                    result.errors.normal = DescribeComException(normalEx);

                    try
                    {
                        workbook = OpenWorkbook(workbooks, inputPath, Excel.XlCorruptLoad.xlRepairFile);
                        result.successfulMode = "repair";
                    }
                    catch (COMException repairEx)
                    {
                        result.errors.repair = DescribeComException(repairEx);

                        try
                        {
                            workbook = OpenWorkbook(workbooks, inputPath, Excel.XlCorruptLoad.xlExtractData);
                            result.successfulMode = "extractData";
                        }
                        catch (COMException extractEx)
                        {
                            result.errors.extractData = DescribeComException(extractEx);
                        }
                    }
                }

                return result;
            }
            finally
            {
                try
                {
                    if (workbook != null)
                    {
                        workbook.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(workbook);
                    }
                }
                catch { /* best-effort cleanup */ }

                try
                {
                    if (workbooks != null)
                    {
                        Marshal.ReleaseComObject(workbooks);
                    }
                }
                catch { /* best-effort cleanup */ }

                try
                {
                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch { /* best-effort cleanup */ }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static string DescribeComException(COMException ex)
        {
            return $"HRESULT 0x{ex.HResult:X8} - {ex.Message}";
        }

        private static Excel.Workbook OpenWorkbook(Excel.Workbooks workbooks, string inputPath, Excel.XlCorruptLoad corruptLoad)
        {
            // Format/Delimiter/Converter params are only for opening text files
            // and apparently trigger errors if passed when loading an .xlsx.
            // So always omit them!
            return workbooks.Open(
                Filename: inputPath,
                UpdateLinks: 0,
                ReadOnly: false,
                IgnoreReadOnlyRecommended: false,
                Notify: false,
                AddToMru: false,
                CorruptLoad: corruptLoad
            );
        }
    }
}
