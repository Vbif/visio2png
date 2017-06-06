using System;
using System.IO;
using System.Linq;

using Microsoft.Office.Interop.Visio;

namespace Visio2Png
{
    class Program
    {
        static void Main(string[] args)
        {
            var visOpenRO = 2;
            var visOpenMinimized = 16;
            var visOpenHidden = 64;
            var visOpenMacrosDisabled = 128;
            var visOpenNoWorkspace = 256;

            var currentDir = new DirectoryInfo(".");
            var fileList = currentDir.EnumerateFiles("*.vsd").ToList();

            var app = new Application();

            app.Application.Settings.SetRasterExportResolution(
                VisRasterExportResolution.visRasterUseCustomResolution,
                300,
                300, VisRasterExportResolutionUnits.visRasterPixelsPerInch);
            app.Application.Settings.SetRasterExportSize(VisRasterExportSize.visRasterFitToSourceSize);

            var flags = Convert.ToInt16(visOpenRO + visOpenMinimized + visOpenHidden + visOpenMacrosDisabled + visOpenNoWorkspace);

            foreach (var file in fileList)
            {
                Console.WriteLine("Open {0}", file.Name);
                var doc = app.Documents.OpenEx(file.FullName, flags);
                var pages = doc.Pages;

                for (var i = 0; i < pages.Count; i++) {

                    var p = doc.Pages[i + 1];
                    var resultName = "";
                    if (pages.Count == 1 && p.Name.Contains("Страница"))
                        resultName = System.IO.Path.GetFileNameWithoutExtension(file.Name) + ".png";
                    else
                        resultName = System.IO.Path.GetFileNameWithoutExtension(file.Name) + " " + p.Name + ".png";

                    Console.WriteLine("\tSave {0}", resultName);
                    p.Export(System.IO.Path.Combine(file.Directory.FullName, resultName));
                }

                doc.Close();
            }


            app.Quit();
        }
    }
}
