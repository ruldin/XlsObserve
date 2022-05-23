using System;
using Microsoft.Extensions.Configuration;
using XlsObserve.Class.FileObserve;
using XlsObserve.Class.logs;

namespace XlsObserve.Class.configuration
{
    internal  class ClsConfiguration
    {
        public string ObserveDirectory { get;  }
        public string OutPutProcessedDirectory { get; }
        public string OutPutNAplicableDirectory { get;  }
        public string ExtentionsForWatch { get; }
        public string TempDirectory { get; set; }
        public string FileNameMasterWorkBook { get; set; }
        public int IndexSheetFromMasterWorkBook { get; set; }


        //the option to obtain the information by environment variables is left enabled
        private IConfiguration Cf = new ConfigurationBuilder()
    .SetBasePath(Directory.GetParent(AppContext.BaseDirectory).FullName)
    .AddJsonFile("appSettings.json", optional: true, reloadOnChange: true)
.AddEnvironmentVariables().Build();



        public ClsConfiguration() {
            
            //for security reason, the file appSettings not shoud display on screen
            //like message, but, just for evaluation propose i put it
            ClsLogs.StatusTrace("Gettering configuration from appSettings.json ....", ClsLogs.info);
            ClsLogs.StatusTrace("checking for files and folders.....", ClsLogs.info);
            OutPutProcessedDirectory = Cf.GetSection("folderspath")["OutPutProcessedDirectory"];
            ObserveDirectory = Cf.GetSection("folderspath")["ObserveDirectory"];
            OutPutNAplicableDirectory = Cf.GetSection("folderspath")["OutPutNAplicableDirectory"];
            ExtentionsForWatch = Cf.GetSection("folderspath")["ExtentionsForWatch"];
            TempDirectory = Cf.GetSection("folderspath")["TempDirectory"];
            FileNameMasterWorkBook = Cf.GetSection("folderspath")["FileNameMasterWorkBook"];
            
            int result = 0;
            if (Int32.TryParse(Cf.GetSection("folderspath")["IndexSheetFromMasterWorkBook"], out result))
            {
                IndexSheetFromMasterWorkBook = result;
            }
            else
            {
                //if there is an error in the conversion, the index is set to 0
                IndexSheetFromMasterWorkBook = result;
            }

            if (!CheckEnvoronment())
            {
                ClsLogs.StatusTrace("Sorry, i need meet all requirements for running... Please check appSettings.json", ClsLogs.info);
                System.Environment.Exit(-1);
            }



            ClsLogs.StatusTrace("The Observer is working now, awaiting for processing files...", ClsLogs.info);

        }



        /// <summary>
        /// evaluate if all folders and files is in the correct place
        /// is important evaluate all params before run the app
        /// if some directory not exist, then is create automatly
        /// if there are error for create new directory, the program ends
        /// </summary>
        /// <returns>true= OK, false = there are a trouble</returns>
        private Boolean CheckEnvoronment()
        {

            if (!ClsFileManager.CheckExistDirectory(ObserveDirectory))
            {
                ClsLogs.ErrorLog("The Observe Directory not Found!!!");
                return false;
            }
            else
            {
                ClsLogs.StatusTrace("Directory to Observer is Ok.....", ClsLogs.info);
            }

            if (!ClsFileManager.CheckExistDirectory(OutPutProcessedDirectory))
            {
                ClsLogs.ErrorLog("The Output of Processed Directory not Found!!!");
                return false;
            }
            else
            {
                ClsLogs.StatusTrace("Directory 'Processed'is Ok.....", ClsLogs.info);
            }

            if (!ClsFileManager.CheckExistDirectory(OutPutNAplicableDirectory))
            {
                ClsLogs.ErrorLog("The 'Not Aplicable' Directory not Found!!!");
                return false;
            }
            else
            {
                ClsLogs.StatusTrace("The 'Not Aplicable' Directory is OK!", ClsLogs.info);
            }




            if (!ClsFileManager.CheckExistDirectory(TempDirectory))
            {
                //ClsLogs.ErrorLog("The 'Temporary' Directory not Found!!!");
                //return false;
                //is not mandatory, is just suggestion
            }
            else
            {
                //ClsLogs.StatusTrace("tmp folder is OK", ClsLogs.info);
            }



            if (!File.Exists(FileNameMasterWorkBook))
            {
                ClsLogs.ErrorLog("The Master WorkBook is not found, this file is ver important for the application");
                return false;
            }
            else
            {
                ClsLogs.StatusTrace("Master WorkBook is found!!!!", ClsLogs.info);
            }
            return true;
        }



    }
}
