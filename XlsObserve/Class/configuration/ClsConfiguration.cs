using System;
using Microsoft.Extensions.Configuration;
using XlsObserve.Class.FileObserve;

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




        public ClsConfiguration()
        {
            //IConfiguration Config = new ConfigurationBuilder().
            //    .AddJsonFile("appSettings.json")
            //    .Build();

            //var URL = Config.GetSection("URL").Value;
            //todo put this in configuration file

            //TODO verificar que el archivo termine con \ en los paths
            ObserveDirectory = @"c:\tmp2\xls";
            OutPutProcessedDirectory = @"c:\tmp2\xls\Processed\";
            OutPutNAplicableDirectory = @"c:\tmp2\xls\NotApplicable\";
            ExtentionsForWatch = "xls,xlsx";
            TempDirectory = @"c:\tmp2\xls\tmp\";
            FileNameMasterWorkBook = @"C:\tmp2\xls\consolidated\ConsolidateWB.xlsx";
            IndexSheetFromMasterWorkBook = 0;




            //evaluate if all run fine, else, stop the program
            ClsFileManager.CheckExistDirectory(OutPutProcessedDirectory);
            ClsFileManager.CheckExistDirectory(OutPutNAplicableDirectory);
            ClsFileManager.CheckExistDirectory(TempDirectory);


            Console.WriteLine("The Observer is working now, awaiting for processing files...");

        }




      



    }
}
