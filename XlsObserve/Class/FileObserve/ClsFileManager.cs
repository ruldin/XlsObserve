using XlsObserve.Class.logs;

namespace XlsObserve.Class.FileObserve
{
    public static class ClsFileManager
    {
        public static int MoveFile(string source, string target)
        {
            int status = -1;
            string? TargetDirName = Path.GetDirectoryName(target);
            try
            {
                if (CheckExistDirectory(TargetDirName))
                {
                    File.Move(source, target);
                    return 1;
                }
                
            } catch (Exception ex)
            {
                ClsLogs.ErrorLog($"Error Moving file {source} to {target} Err:{ex.Message}");
                status = -1;
            }
            return status;
        }



        public static Boolean CheckExistDirectory(string? DirectoryPath)
        {
            if (!Directory.Exists(DirectoryPath))
            {
                try
                {
                    Directory.CreateDirectory(DirectoryPath);
                    return true;
                }
                catch (Exception ex)
                {
                    //put log error and send status
                    return false;
                }

            }
            return true;
        }

    }
}
