using XlsObserve.Class.configuration;
using XlsObserve.Class.logs;
using XlsObserve.Class.xlsServices;

namespace XlsObserve.Class.FileObserve
{
    /// <summary>
    /// FileObserve
    /// developed by Ruldin Ayala
    /// </summary>
    internal class ClsFileObserve
    {

        ClsConfiguration conf = new();
        XlsServices oxls = new();


        public void start()
        {

            //FileSystemWatcher uses its own thread, await not required 
            FileSystemWatcher observador = new FileSystemWatcher(conf.ObserveDirectory);
            observador.NotifyFilter = (NotifyFilters.FileName);

            observador.Changed += OnChange;
            observador.Created += OnChange;
            observador.Error += OnError;

            observador.EnableRaisingEvents = true;

            Console.WriteLine(@"

            ");

            Console.WriteLine("enter para terminar");
            Console.ReadLine();

        }



        private Boolean isXls(string fileName)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Extension.Contains("xls")) //include xlsx and xls
            {
                string destination = conf.OutPutProcessedDirectory + fileInfo.Name;

                //todo: evaluate to use thread for processing file
                //

                var dt = oxls.Excel_To_DataTable(fileName,conf.IndexSheetFromMasterWorkBook);

                if(dt != null)
                {
                    int rows = oxls.AppendData(conf.FileNameMasterWorkBook, dt, conf.IndexSheetFromMasterWorkBook);
                    ClsLogs.StatusTrace($"Append {rows} rows.",ClsLogs.info);
                }
                else
                {
                    ClsLogs.ErrorLog($"No data found at {fileName} file");
                }

                ClsLogs.StatusTrace($"moving XLS file {fileInfo.Name}", ClsLogs.actions);
                ClsFileManager.MoveFile(fileName, destination);

                return true;
            } else
            {
                string destination = conf.OutPutNAplicableDirectory + fileInfo.Name;
                ClsLogs.StatusTrace($"Moving Not applicable file... {fileInfo.Name}", ClsLogs.actions);
                ClsFileManager.MoveFile(fileName, destination);
            }

            return false;

        }

        private void OnChange(object source, FileSystemEventArgs e)
        {
            ClsLogs.StatusTrace($"Getting new file...", ClsLogs.info);
            WatcherChangeTypes KindOfChange = e.ChangeType;
            isXls(e.FullPath);
            
        }

        private void OnError(object source,ErrorEventArgs e)
        {

            ClsLogs.ErrorLog("e.GetException().Message");
            
        }



    }
}
