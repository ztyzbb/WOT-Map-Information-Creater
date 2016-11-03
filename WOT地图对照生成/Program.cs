using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Xml;
using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
namespace WOT地图对照生成
{
    class Program
    {
        static int Main(string[] args)
        {
            Console.WriteLine("*********************************************************************");
            Console.WriteLine("*                         WOT地图对照生成 v1.0                      *");
            Console.WriteLine("*                    在坦克世界v0.9.16.0下测试通过                  *");
            Console.WriteLine("*  该程序修改并使用了来自World of Tanks Mod Tools Library的部分代码 *");
            Console.WriteLine("*                  http://wottoolslib.codeplex.com/                 *");
            Console.WriteLine("*                wottoolslib是修改自WoTModTools的类库               *");
            Console.WriteLine("*              https://github.com/katzsmile/WoTModTools             *");
            Console.WriteLine("*             同时该程序也调用了GNU gettext的反编译程序             *");
            Console.WriteLine("*               http://www.gnu.org/software/gettext/                *");
            Console.WriteLine("*       该程序从游戏文件为每个地图生成一些录像解析必要的参数        *");
            Console.WriteLine("*                    具体使用方法详见Readme.md                      *");
            Console.WriteLine("*                         本软件遵循GPLv3协议                       *");
            Console.WriteLine("*                           祝您使用愉快！                          *");
            Console.WriteLine("*                      ztyzbb 2016.11.03 敬上！                     *");
            Console.WriteLine("*********************************************************************");

            Console.Write("请输入坦克世界安装目录（如C:\\Games\\World_of_Tanks）：");
            string path_org = Console.ReadLine();
            if (path_org.EndsWith("\\"))
                path_org = path_org.TrimEnd('\\');
            string path = path_org + "\\res\\text\\LC_MESSAGES\\arenas.mo";

            ProcessStartInfo p = null;
            Process Proc = null;

            Console.WriteLine();
            Console.WriteLine("当前arenas.mo路径：" + path);
            Console.WriteLine("请务必确保路径合法有效，是否继续？如否，请右上角");
            Console.ReadLine();

            p = new ProcessStartInfo("msgunfmt.exe", path + " -o arenas.po");
            p.RedirectStandardOutput = true;
            p.UseShellExecute = false;
            try
            {
                Proc = Process.Start(p);
            }
            catch (Exception e)
            {
                Console.WriteLine();
                Console.WriteLine("启动msgunfmt.exe失败");
                Console.WriteLine("Command=msgunfmt.exe " + path + " -o arenas.po");
                Console.WriteLine(e.Message);
                Console.ReadLine();
                return -1;
            }

            if (!Proc.WaitForExit(1000))
            {
                Console.WriteLine();
                Console.WriteLine("mo解密超时！将继续等待");
                Proc.WaitForExit();
            }

            if (Convert.ToBoolean(Proc.ExitCode))
            {
                Console.WriteLine();
                Console.WriteLine("mo反编译失败，请根据提示排除故障！");
                Console.WriteLine("Command=msgunfmt.exe " + path + "-o arenas.po");
                Console.WriteLine("返回值=" + Proc.ExitCode);
                Console.ReadLine();
                return -2;
            }
            Console.WriteLine("成功反编译arenas.mo");

            path = path_org + "\\res\\scripts\\arena_defs\\_list_.xml";
            Console.WriteLine();
            Console.WriteLine("当前arena_defs\\_list_.xml路径：" + path);
            Console.WriteLine("请务必确保路径合法有效，是否继续？如否，请右上角");
            Console.ReadLine();

            p.FileName = "wottoolslib.exe";
            p.Arguments = path;
            try
            {
                Proc = Process.Start(p);
            }
            catch (Exception e)
            {
                Console.WriteLine();
                Console.WriteLine("启动wottoolslib.exe失败");
                Console.WriteLine("Command=wottoolslib.exe " + path);
                Console.WriteLine(e.Message);
                Console.ReadLine();
                return -3;
            }

            if (!Proc.WaitForExit(1000))
            {
                Console.WriteLine();
                Console.WriteLine("xml解密超时！将继续等待");
                Proc.WaitForExit();
            }

            string result = Proc.StandardOutput.ReadToEnd();
            if (Convert.ToBoolean(Proc.ExitCode))
            {
                Console.WriteLine();
                Console.WriteLine(result);
                Console.WriteLine("arena_defs\\_list_.xml解析失败，请根据提示排除故障！");
                Console.WriteLine("Command=wottoolslib.exe " + path);
                Console.WriteLine("返回值=" + Proc.ExitCode);
                Console.ReadLine();
                return -4;
            }
            XmlDocument xml = new XmlDocument();
            try
            {
                xml.LoadXml(result);
            }
            catch (Exception e)
            {
                Console.WriteLine();
                Console.WriteLine(e.Message);
                Console.WriteLine("arena_defs\\_list_.xml解析失败，请根据提示排除故障！");
            }

            Console.WriteLine("选择存储方式：");
            Console.WriteLine("0.json");
            Console.WriteLine("1.excel");
            Console.Write("输入任意值以使用excel,直接回车以使用json");
            int mapcounter = 0,namecounter=0;
            if (Console.ReadLine().Length == 0)
            {
                JArray data = new JArray();
                XmlNode root = xml.SelectSingleNode("_list_.xml");
                XmlNodeList maps = root.ChildNodes;
                
                foreach (XmlNode currentmap in maps)
                {
                    data.Add(new JObject());
                    data[mapcounter]["mapid"] = Int32.Parse(currentmap.SelectSingleNode("id").FirstChild.Value.Remove(0, 1));
                    data[mapcounter]["mapidname"] = currentmap.SelectSingleNode("name").FirstChild.Value;
                    mapcounter++;
                }
                using (StreamReader file = new StreamReader("arenas.po"))
                {
                    string line = file.ReadLine();
                    while (!file.EndOfStream)
                    {
                        if (!line.StartsWith("msgid") || !line.EndsWith("name\""))
                        {
                            line = file.ReadLine();
                            continue;
                        }
                        line = line.Remove(0, 7);
                        line = line.Remove(line.Length - 1);
                        line = line.Split('/')[0];
                        for (int j = 0; j < data.Count; j++)
                        {
                            if ((string)(data[j]["mapidname"]) == line)
                            {
                                line = file.ReadLine();
                                line = line.Remove(0, 8);
                                line = line.Remove(line.Length - 1);
                                data[j]["mapname"] = line;
                                namecounter++;
                                break;
                            }
                        }
                    }
                }
                try
                {
                    StreamWriter sw = new StreamWriter("maps.json", false, Encoding.UTF8);
                    sw.Write(JsonConvert.SerializeObject(data));
                    sw.Close();
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                    Console.ReadLine();
                    return -5;
                }
            }

            else
            {
                Excel.Application excelapp;
                Excel.Workbook excelbook;
                Excel.Worksheet excelsheet;
                
                try
                {
                    Object Nothing = System.Reflection.Missing.Value;
                    excelapp = new Excel.Application();
                    excelapp.Visible = true;
                    excelbook = excelapp.Workbooks.Add(Nothing);
                    excelsheet = excelbook.Sheets[1];
                    XmlNode root = xml.SelectSingleNode("_list_.xml");
                    XmlNodeList maps = root.ChildNodes;

                    foreach (XmlNode currentmap in maps)
                    {
                        mapcounter++;
                        excelsheet.Cells[mapcounter, 1].Value = Int32.Parse(currentmap.SelectSingleNode("id").FirstChild.Value.Remove(0, 1));
                        excelsheet.Cells[mapcounter, 2].Value = currentmap.SelectSingleNode("name").FirstChild.Value;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.ReadKey();
                    return -6;
                }
                using (StreamReader file = new StreamReader("arenas.po"))
                {
                    string line = file.ReadLine();
                    while (!file.EndOfStream)
                    {
                        if (!line.StartsWith("msgid") || !line.EndsWith("name\""))
                        {
                            line = file.ReadLine();
                            continue;
                        }
                        line = line.Remove(0, 7);
                        line = line.Remove(line.Length - 1);
                        line = line.Split('/')[0];
                        for (int j = 1; j <= mapcounter; j++)
                        {
                            if (excelsheet.Cells[j, 2].Value == line)
                            {
                                line = file.ReadLine();
                                line = line.Remove(0, 8);
                                line = line.Remove(line.Length - 1);
                                excelsheet.Cells[j, 3].Value = line;
                                namecounter++;
                                break;
                            }
                        }
                    }
                }
                excelbook.SaveAs(Directory.GetCurrentDirectory() + "\\maps.xlsx");
                excelbook.Close();
                excelapp.Quit();
            }
            Console.WriteLine();
            if (namecounter != mapcounter)
                Console.WriteLine("检测到潜在的数据缺失！请手动检查导出的文件！");
            Console.WriteLine("文件已经导出到" + Directory.GetCurrentDirectory());
            Console.WriteLine("如要保留arenas.po，请输入任意值，否则直接回车");
            if (Console.ReadLine().Length == 0)
                File.Delete("arenas.po");
            Console.WriteLine("完成！ztyzbb于2016.11.02敬上！");
            Console.ReadLine();
            return 0;
        }
    }
}
