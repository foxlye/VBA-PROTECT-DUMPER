using System;
using System.IO;
using System.Text;
using System.Xml;
using ICSharpCode.SharpZipLib.Core;

namespace VBA_PROTECT_DUMPER
{
    internal class Program
    {
        static string g_VBA_target = null;
        static string g_filename = null;

        static void Main(string[] args)
        {
            Console.Title = "VBA PROTECT DUMPER";
            EntryPoint();
            Console.ReadKey();
        }


        static void EntryPoint()
        {
            ConsoleColorController(ConsoleColor.Gray);
            logger("Drop file to this directory.");
            logger("Enter filename (with extension): ");
            if (!UnpackFile(Console.ReadLine()))
            {
                ConsoleColorController(ConsoleColor.Red);
                logger("Failed to unpack file.");
                logger("---------------------");
                EntryPoint();
            }

            ConsoleColorController(ConsoleColor.Gray);
            logger("Getting VBA Target..");

            if (!GetVBATarget())
            {
                ConsoleColorController(ConsoleColor.Red);
                logger("Failed to get VBA target.");
                logger("---------------------");
                EntryPoint();
            }

            ConsoleColorController(ConsoleColor.Gray);
            logger("Patching VBA Target..");

            if (!PatchVBATarget())
            {
                ConsoleColorController(ConsoleColor.Red);
                logger("Failed to patch VBA target.");
                logger("---------------------");
                EntryPoint();
            }

            ConsoleColorController(ConsoleColor.Gray);
            logger("Creating patched file..");

            if (!CreatePatchedFile())
            {
                ConsoleColorController(ConsoleColor.Red);
                logger("Failed to create patched file.");
                logger("---------------------");
                EntryPoint();
            }

            ConsoleColorController(ConsoleColor.Green);
            logger("Project successfully unprotected.");
        }


        static void logger(string log)
        {
            Console.WriteLine($"[github.com/foxlye] {log}");
        }

        static void ConsoleColorController(ConsoleColor color)
        {
            Console.ForegroundColor = color;
        }

        static bool UnpackFile(string filename)
        {
            try
            {
                ConsoleColorController(ConsoleColor.Gray);
                logger("Unpacking file..");

                if (!File.Exists(Environment.CurrentDirectory + $"\\{filename}"))
                    return false;

                byte[] buffer = new byte[4096];
                buffer = File.ReadAllBytes(Environment.CurrentDirectory + $"\\{filename}");
                Stream stream = new MemoryStream(buffer);
                stream.Position = 0;
                if (Directory.Exists($"{Environment.CurrentDirectory}\\Temp"))
                    Directory.Delete($"{Environment.CurrentDirectory}\\Temp", true);

                ExtractZipFile(stream, "", @"Temp");

                g_filename = filename.Split('.')[0] + "_patched." + filename.Split('.')[1];

                return true;
            }
            catch (Exception ex)
            {
                logger(ex.Message);
                return false;
            }

        }



        static bool GetVBATarget()
        {
            try
            {
                if (!Directory.Exists($"{Environment.CurrentDirectory}\\Temp"))
                    return false;

                if (!File.Exists($"{Environment.CurrentDirectory}\\Temp\\xl\\_rels\\workbook.xml.rels"))
                    return false;

                XmlDocument parser = new XmlDocument() { XmlResolver = null };

                parser.Load($"{Environment.CurrentDirectory}\\Temp\\xl\\_rels\\workbook.xml.rels");

                XmlElement xml_root = parser.DocumentElement;

                if (xml_root == null)
                    return false;

                foreach (XmlElement xml_node in xml_root)
                {
                    string type = null;

                    foreach (XmlAttribute xml_elements in xml_node.Attributes)
                    {
                        if (xml_elements.Name == "Type")
                        {
                            if (xml_elements.InnerText.Contains("vbaProject"))
                                type = xml_elements.InnerText;
                        }

                        if (type != null)
                        {
                            if (xml_elements.Name == "Target")
                            {
                                g_VBA_target = xml_elements.InnerText;
                                break;
                            }
                        }
                    }

                    if (g_VBA_target != null)
                        break;
                }


                if (g_VBA_target == null)
                    return false;


                return true;
            }
            catch (Exception ex)
            {
                logger(ex.Message);
                return false;
            }


        }


        static bool PatchVBATarget()
        {
            try
            {
                if (g_VBA_target == null)
                    return false;

                if (!File.Exists($"{Environment.CurrentDirectory}\\Temp\\xl\\{g_VBA_target}"))
                    return false;

                byte[] vbaTargetData = File.ReadAllBytes($"{Environment.CurrentDirectory}\\Temp\\xl\\{g_VBA_target}");

                if (vbaTargetData == null)
                    return false;

                bool[] changed = { false, false, false };

                string string_data = Encoding.Default.GetString(vbaTargetData);


                if (string_data.Contains("CMG="))
                {
                    StringBuilder sb = new StringBuilder(string_data);
                    sb.Replace("CMG=", "CMC=");
                    string_data = sb.ToString(); ;


                    if (!string_data.Contains("CMG="))
                        changed[0] = true;
                }



                if (string_data.Contains("DPB="))
                {
                    StringBuilder sb = new StringBuilder(string_data);
                    sb.Replace("DPB=", "DPD=");
                    string_data = sb.ToString();


                    if (!string_data.Contains("DPB="))
                        changed[1] = true;
                }


                if (string_data.Contains("GC="))
                {
                    StringBuilder sb = new StringBuilder(string_data);
                    sb.Replace("GC=", "CC=");
                    string_data = sb.ToString();


                    if (!string_data.Contains("GC="))
                        changed[2] = true;
                }


                if (!changed[0] || !changed[1] || !changed[2])
                    return false;

                if (File.Exists($"{Environment.CurrentDirectory}\\Temp\\xl\\{g_VBA_target}"))
                {
                    File.Delete($"{Environment.CurrentDirectory}\\Temp\\xl\\{g_VBA_target}");
                }

                byte[] PatchedBytes = Encoding.Default.GetBytes(string_data);


                using (var fs = new FileStream($"{Environment.CurrentDirectory}\\Temp\\xl\\{g_VBA_target}", FileMode.Create, FileAccess.Write))
                {
                    fs.Write(PatchedBytes, 0, PatchedBytes.Length);
                    fs.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                logger(ex.Message);
                return false;
            }

        }

        static bool CreatePatchedFile()
        {
            try
            {
                using (Ionic.Zip.ZipFile zip = new Ionic.Zip.ZipFile($"{g_filename}"))
                {
                    zip.AddDirectory($"{Environment.CurrentDirectory}\\Temp");
                    zip.Save();
                }

                if (Directory.Exists($"{Environment.CurrentDirectory}\\Temp"))
                    Directory.Delete($"{Environment.CurrentDirectory}\\Temp", true);


                return true;
            }
            catch (Exception ex)
            {
                logger(ex.Message);
                return false;
            }
        }


        public static void ExtractZipFile(Stream inputStream, string password, string outFolder)
        {
            ICSharpCode.SharpZipLib.Zip.ZipFile zf = null;
            try
            {
                zf = new ICSharpCode.SharpZipLib.Zip.ZipFile(inputStream);
                if (!string.IsNullOrEmpty(password))
                {
                    zf.Password = password;
                }
                foreach (ICSharpCode.SharpZipLib.Zip.ZipEntry zipEntry in zf)
                {
                    if (!zipEntry.IsFile)
                    {
                        continue;
                    }
                    String entryFileName = zipEntry.Name;


                    byte[] buffer = new byte[4096];
                    Stream zipStream = zf.GetInputStream(zipEntry);

                    String fullZipToPath = Path.Combine(outFolder, entryFileName);
                    string directoryName = Path.GetDirectoryName(fullZipToPath);
                    if (directoryName.Length > 0)
                        Directory.CreateDirectory(directoryName);


                    using (FileStream streamWriter = File.Create(fullZipToPath))
                    {
                        StreamUtils.Copy(zipStream, streamWriter, buffer);
                    }
                }
            }
            finally
            {
                if (zf != null)
                {
                    zf.IsStreamOwner = true;
                    zf.Close();
                }
            }
        }
    }
}
