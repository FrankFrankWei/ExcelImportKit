/******************************************************************
** auth: wei.huazhong
** date: 9/20/2018 11:53:39 AM
** desc:
******************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    public class RootDirectoryHelper
    {
        public static string GetFilePath(string relativeToAppRoot)
        {
            Uri sheetConfigUri = new Uri(GetAppRootDirectory(), relativeToAppRoot);
            return sheetConfigUri.LocalPath;
        }

        public static Uri GetAppRootDirectory()
        {
            Uri baseUri = new Uri(AppDomain.CurrentDomain.BaseDirectory);
            var relativePath = IsClientApp(baseUri) ? @"..\..\" : @".\";
            Uri rootUri = new Uri(baseUri, relativePath);
            return rootUri;
        }

        /// <summary>
        /// client app means windows form app or app with the same directory hierachy .
        /// windows form app's AppDomain.CurrentDomain.BaseDirectory is "ProjectName/bin/Debug/" or "ProjectName/bin/Release"
        /// our Configs folder which holds xml config files locate at "ProjectName/"
        /// so relative path is different between windows form app and web app
        /// </summary>
        /// <param name="uri"></param>
        /// <returns></returns>
        private static bool IsClientApp(Uri uri)
        {
            return uri.LocalPath.Contains("bin");
        }

    }
}
