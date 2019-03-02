using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Reflection;
using System.IO;
using System.Xml.Linq;
using System.IO.Compression;
using System.Diagnostics;

namespace dotnetcore_officeaddin_toolbox
{
    class Program
    {
        //https://github.com/OfficeDev/office-toolbox/blob/master/src/util.ts#L499
        //office-toolbox sideload -m manifest.xml -a Excel
        static void Main(string[] args)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                if (args.Length < 3)
                {
                    Console.WriteLine("当前该工具支持三个参数，第一个参数是命令名，当前仅支持sideload，第二个参数是manifest的文件位置，第三个参数为应用程序名称，目前支持Word，Excel，PowerPoint");
                    return;
                    //TODO: 这里如何更好地解析命令参数
                }

                var command = args[0];//命令
                var manifest = Path.GetFullPath(args[1]);//manifest文件位置
                var app = args[2];//应用程序

                //增加一个注册表条目
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Wef\Developer", true))
                {
                    key.SetValue(manifest, manifest);
                }

                //获取模板目录
                var dir = new DirectoryInfo(Assembly.GetExecutingAssembly().Location);
                var templatesPath = dir.Parent.Parent.Parent.GetDirectories("templates")[0];

                //参数设置
                var settings = new
                {
                    Word = new
                    {
                        TaskPaneApp = new
                        {
                            WebExtensionPath = "word/webextensions/webextension.xml",
                            TemplateName = "DocumentWithTaskPane.docx"
                        }
                    },
                    Excel = new
                    {
                        TaskPaneApp = new
                        {
                            WebExtensionPath = "xl/webextensions/webextension.xml",
                            TemplateName = "BookWithTaskPane.xlsx"
                        },
                        ContentApp = new
                        {
                            WebExtensionPath = "xl/webextensions/webextension.xml",
                            TemplateName = "BookWithContent.xlsx"
                        }
                    },
                    PowerPoint = new
                    {
                        TaskPaneApp = new
                        {
                            WebExtensionPath = "ppt/webextensions/webextension.xml",
                            TemplateName = "PresentationWithTaskPane.pptx"
                        },
                        ContentApp = new
                        {
                            WebExtensionPath = "ppt/slides/udata/data.xml",
                            TemplateName = "PresentationWithContent.pptx"
                        }
                    }
                };

                var info = GetManifestInfo(manifest);


                var templateFileName = string.Empty;
                var extensionPath = string.Empty;
                var extension = string.Empty;
                var generatedFileName = string.Empty;

                switch (app.ToLower())
                {
                    case "word":
                        if (info.Type == "TaskPaneApp")
                        {
                            templateFileName = settings.Word.TaskPaneApp.TemplateName;
                            extensionPath = settings.Word.TaskPaneApp.WebExtensionPath;
                        }
                        break;
                    case "excel":
                        if (info.Type == "TaskPaneApp")
                        {
                            templateFileName = settings.Excel.TaskPaneApp.TemplateName;
                            extensionPath = settings.Excel.TaskPaneApp.WebExtensionPath;
                        }
                        else if (info.Type == "ContentApp")
                        {
                            templateFileName = settings.Excel.ContentApp.TemplateName;
                            extensionPath = settings.Excel.ContentApp.WebExtensionPath;
                        }
                        break;
                    case "powerpoint":
                        if (info.Type == "TaskPaneApp")
                        {
                            templateFileName = settings.PowerPoint.TaskPaneApp.TemplateName;
                            extensionPath = settings.PowerPoint.TaskPaneApp.WebExtensionPath;
                        }
                        else if (info.Type == "ContentApp")
                        {
                            templateFileName = settings.PowerPoint.ContentApp.TemplateName;
                            extensionPath = settings.PowerPoint.ContentApp.WebExtensionPath;
                        }
                        break;
                    default:
                        break;
                }
                extension = Path.GetExtension(templateFileName);//扩展名
                generatedFileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + extension);//临时的文件名
                templateFileName = Path.Combine(templatesPath.FullName, templateFileName);//模板的完整路径

                File.Copy(templateFileName, generatedFileName);
                using (ZipArchive zip = ZipFile.Open(generatedFileName, ZipArchiveMode.Update))
                {
                    var entry = zip.GetEntry(extensionPath);
                    var entrystream = entry.Open();
                    var text = new StreamReader(entrystream).ReadToEnd();
                    text = text.Replace("00000000-0000-0000-0000-000000000000", info.Id);
                    text = text.Replace("1.0.0.0", info.Version);

                    entrystream.Close();
                    entry.Delete();//删除这个entry

                    var newEntry = zip.CreateEntry(extensionPath);
                    using (StreamWriter writer = new StreamWriter(newEntry.Open()))
                    {
                        writer.WriteLine(text);
                        writer.Flush();
                    }
                }

                //启动这个文件
                var p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    UseShellExecute = true,
                    FileName = generatedFileName
                };
                p.Start();

            }
            else
            {
                Console.WriteLine("这个工具仅支持在Windows上面运行.This tool is only support for Windows.");
            }

        }

        static ManifestInfo GetManifestInfo(string path)
        {
            XNamespace xn = "http://schemas.microsoft.com/office/appforoffice/1.1";
            var doc = XDocument.Load(path);
            var id = doc.Root.Element(xn + "Id").Value;
            var version = doc.Root.Element(xn + "Version").Value;
            XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
            var type = doc.Root.Attribute(xsi + "type").Value;

            return new ManifestInfo()
            {
                Type = type,
                Id = id,
                Version = version
            };
        }
    }

    struct ManifestInfo
    {
        internal string Type { get; set; }
        internal string Id { get; set; }
        internal string Version { get; set; }
    }

}
