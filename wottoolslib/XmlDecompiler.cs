/*
 * 这是World of Tanks Mod Tools Library
 * http://wottoolslib.codeplex.com/
 * 是World of Tanks Mod Tools的类库版
 * https://github.com/katzsmile/WoTModTools
 * 在引入WOT对照生成时修改成控制台程序用于被主程序调用
 * 所有输出全部输出到stdout，经过pipe被主程序接收
 * 在WoT 0.9.15.1版本下测试通过
 * 如有任何疑问，请加QQ741073274
 * ztyzbb 2016.09.05
*/
using System;
using System.Collections.Generic;
using System.Linq;
using Packed_Section_Reader;
using Primitive_File_Reader;
using System.IO;
using System.Xml;

namespace wottoolslib
{
    public class XmlDecompiler
    {
        private static int currentstate = 0;
        public static int Main(string[] args)
        {
            XmlDecompiler XD = Instance;
            if (args[0].Length != 0)
                Console.Write(XD.GetFileXml(args[0]));
            else
            {
                Console.Write("调用命令错误，未能接收到参数");
                return 1;
            }
            return currentstate;
        }
        private XmlDecompiler() { }
        private static XmlDecompiler _instance = new XmlDecompiler();
        public static XmlDecompiler Instance
        {
            get
            {
                return _instance;
            }
        }
        public Packed_Section PackedSection = new Packed_Section();
        public Primitive_File PackedFile = new Primitive_File();

        public string GetFileXml(string fileName)
        {
            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (BinaryReader reader = new BinaryReader(fileStream))
                    {
                        int head = reader.ReadInt32();
                        if (head == Packed_Section.Packed_Header)
                        {
                            return ReadPackedFileAsXml(reader, fileName);
                        }
                        else if (head == Primitive_File.Binary_Header)
                        {
                            return ReadPrimitiveFileAsXml(reader, fileName);
                        }
                        else
                        {
                            currentstate = 5;
                            Console.Write("File was not determined to be a packed or primitive file.");
                            return string.Empty;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                currentstate = 2;
                Console.Write("Could not open file '" + fileName + "' " + ex.Message);
            }
            return string.Empty;
        }

        private string ReadPackedFileAsXml(BinaryReader reader, string fileName)
        {
            try
            {
                string file = Path.GetFileName(fileName).ToLower();
                reader.ReadSByte();
                List<string> dictionary = PackedSection.readDictionary(reader);
                if (dictionary.Any())
                {
                    XmlDocument xDoc = new XmlDocument();
                    XmlNode xmlroot = xDoc.CreateNode(XmlNodeType.Element, file, "");

                    PackedSection.readElement(reader, xmlroot, xDoc, dictionary);

                    xDoc.AppendChild(xmlroot);
                    return xDoc.OuterXml;
                }
            }
            catch (Exception ex)
            {
                currentstate = 3;
                Console.Write("Could not read packed file '" + fileName + "' " + ex.Message);
            }

            return string.Empty;
        }

        private string ReadPrimitiveFileAsXml(BinaryReader reader, string fileName)
        {
            try
            {
                string file = Path.GetFileName(fileName).ToLower();
                XmlDocument xDoc = new XmlDocument();
                XmlNode xmlPrimitives = xDoc.CreateNode(XmlNodeType.Element, "primitives", "");
                PackedFile.ReadPrimitives(reader, xmlPrimitives, xDoc);
                xDoc.AppendChild(xmlPrimitives);
                return xDoc.OuterXml;
            }
            catch (Exception ex)
            {
                currentstate = 4;
                Console.Write("Could not read primitive file '" + fileName + "' " + ex.Message);
            }
            return string.Empty;
        }
    }
}
