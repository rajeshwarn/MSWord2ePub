using System;
using System.Configuration;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using Ionic.Zip;
using System.Xml;

namespace Word2HTML4ePub
{
    public partial class ePubTools
    {
        static string ePubCheckPack = "epubcheck-3.0.1.zip";

        private static string pSettingFile;
        public static string SettingFile
        {
            get
            {
                if (String.IsNullOrEmpty(pSettingFile))
                {

                    //Get the assembly information
                    System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();

                    //Location is where the assembly is run from 
                    string assemblyLocation = assemblyInfo.Location;

                    ////CodeBase is the location of the ClickOnce deployment files
                    //Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
                    //string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());

                    pSettingFile = assemblyInfo.CodeBase + ".config";
                    //                    pSettingFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ePubTools");
                    //pSettingFile = Application.CommonAppDataPath;
                    //if (!Directory.Exists(pSettingFile))
                    //{
                    //    Directory.CreateDirectory(pSettingFile);
                    //}

                    //pSettingFile = Path.Combine(pSettingFile, Path.GetFileName(Environment.GetCommandLineArgs()[0]) + ".config");

                    if (!File.Exists(pSettingFile))
                    {
                        File.Copy(Path.GetFileName(Environment.GetCommandLineArgs()[0]) + ".config", pSettingFile);
                    }
                }
                return pSettingFile;
            }
        }
        


        private static string pJavaPath;
        public static string JavaPath
        {
            get
            {
                if (string.IsNullOrEmpty(pJavaPath))
                {

                    //On essaye de démarrer java sans fournir le path.
                    System.Diagnostics.Process proc = new System.Diagnostics.Process();// processus de verification
                    proc.StartInfo = new System.Diagnostics.ProcessStartInfo("java.exe");
                    proc.StartInfo.CreateNoWindow = true; // ne pas afficher de fenetre
                    proc.StartInfo.UseShellExecute = false; // On désactive le shell
                    proc.StartInfo.RedirectStandardOutput = true; // On redirige la sortie standard
                    proc.StartInfo.RedirectStandardError = true;

                    try
                    {
                        proc.Start(); // Démarrage du processus    
                        proc.WaitForExit();
                        proc.Close();
                        //Ca marche, on se passera de la lecture du path
                        pJavaPath = "java.exe";
                        return pJavaPath;
                    }
                    catch (Exception e)
                    {//Ca ne marche pas, 
                        
                    }

                    //On va essayer de lire le Settings file
                    if (!string.IsNullOrEmpty(SettingsFile.Default.javaPathSettings))
                        if (File.Exists(SettingsFile.Default.javaPathSettings))
                        {
                            pJavaPath = SettingsFile.Default.javaPathSettings;
                            return pJavaPath;
                        }
                    
                    try
                    {
                        pJavaPath = ReadDeportedAppFileKey("javaPath");
                        if (String.IsNullOrEmpty(pJavaPath))
                        {
                            MessageBox.Show("Mauvais fichier de configuration: " + SettingFile, "Message", MessageBoxButtons.OK);
                            return null;
                            //throw new Exception("Mauvais fichier de configuration: " + SettingFile);
                        }
                        //pJavaPath = ConfigurationManager.AppSettings["javaPath"];
                        if (!File.Exists(pJavaPath))
                        {
                            MessageBox.Show("java introuvable!", "Message", MessageBoxButtons.OK);
                            return null;
                            //throw new Exception("java introuvable!");
                        }
                    }
                    catch (Exception e)
                    {
                        OpenFileDialog file1 = new OpenFileDialog();
                        file1.Title = "Merci de sélectionner l'application java.exe";
                        file1.Filter = "(java.exe)|java.exe";
                        file1.Multiselect = false;
                        if (file1.ShowDialog() != DialogResult.OK)
                            MessageBox.Show("Impossible de trouver java.exe. est-il installé sur cet ordinateur?", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        else
                            JavaPath = file1.FileName;

                    }
                }
                return pJavaPath;
            }
            set
            {
                //Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                //config.AppSettings.Settings.Remove("javaPath");
                //config.AppSettings.Settings.Add("javaPath", value);
                //config.Save(ConfigurationSaveMode.Modified);
                //ConfigurationManager.RefreshSection("appSettings");
                SettingsFile.Default.javaPathSettings = value;
                SettingsFile.Default.Save();
                WriteDeportedAppFileKey("javaPath", value);
                pJavaPath = value;
            }
        }

        private static string pEpubCheckPath;
        public static string EpubCheckPath
        {
            get
            {
                if (string.IsNullOrEmpty(pEpubCheckPath))
                {
                    try
                    {
                        pEpubCheckPath = ReadDeportedAppFileKey("epubCheckerPath");
                        if (String.IsNullOrEmpty(pEpubCheckPath))
                        {
                            MessageBox.Show("Mauvais fichier de configuration: " + SettingFile, "Message", MessageBoxButtons.OK);
                            return null;
                            //throw new Exception("Mauvais fichier de configuration: " + SettingFile);
                        }

                        if (!File.Exists(pEpubCheckPath))
                        {
                            MessageBox.Show("ePubcheck introuvable!", "Message", MessageBoxButtons.OK);
                            return null;
                            //throw new Exception("ePubcheck introuvable!");
                        }
                            
                    }
                    catch (Exception e)
                    {
                        //On va chercher si le répertoire temp contient une archive de epubCheck
                        string ePubCheckerPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(ePubCheckPack));

                        if (Directory.Exists(ePubCheckerPath))
                        {
                            string[] lof = Directory.GetFiles(ePubCheckerPath, "epubcheck*.jar", SearchOption.AllDirectories);
                            if (lof.Length != 0)
                            { 
                                EpubCheckPath = lof[0];
                                return pEpubCheckPath;
                            }
                        }

                        { 
                            System.Reflection.Assembly _assembly = System.Reflection.Assembly.GetExecutingAssembly();
                            Stream _ePubCheckFileStream = _assembly.GetManifestResourceStream("Word2HTML4ePub." + ePubCheckPack);
                            Ionic.Zip.ZipFile zip = Ionic.Zip.ZipFile.Read(_ePubCheckFileStream);
                        
                            //Décompresser le modèle dans un sous folder
                            Directory.CreateDirectory(ePubCheckerPath);
                            zip.ExtractAll(Directory.GetParent(ePubCheckerPath).ToString(), Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);

                            string[] lof = Directory.GetFiles(ePubCheckerPath, "epubcheck*.jar", SearchOption.AllDirectories);
                            if (lof.Length != 0)
                            {
                                EpubCheckPath = lof[0];
                                return pEpubCheckPath;
                            }
                        }

                        //string epubCheckPackFileName = (from FileInfo file in new DirectoryInfo(Application.StartupPath).GetFiles("epubcheck*.zip") 
                        //                                orderby file.CreationTime descending select file.FullName).FirstOrDefault();

                        //if (string.IsNullOrEmpty(epubCheckPackFileName))
                        //{
                        //    OpenFileDialog file1 = new OpenFileDialog();
                        //    file1.Title = "Merci de sélectionner l'application ePubChecker";
                        //    file1.Filter = "(epubcheck*.jar)|epubcheck*.jar";
                        //    file1.Multiselect = false;
                        //    if (file1.ShowDialog() != DialogResult.OK)
                        //        MessageBox.Show("Impossible de trouver epubcheck. est-il installé sur cet ordinateur?", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //    else
                        //        EpubCheckPath = file1.FileName;
                        //}
                        //else
                        //{   //Decompresser le package
                        //    ZipFile zip = ZipFile.Read(epubCheckPackFileName);
                        //    zip.ExtractAll(Application.CommonAppDataPath, ExtractExistingFileAction.OverwriteSilently);
                            
                        //    //Rechercher le epubcheck*.jar et le stocker
                        //    EpubCheckPath = (from FileInfo file in new DirectoryInfo(Application.CommonAppDataPath).GetFiles("epubcheck*.jar", SearchOption.AllDirectories)
                        //                                    orderby file.CreationTime ascending
                        //                                    select file.FullName).LastOrDefault();
                        //}
                    }
                }
                return pEpubCheckPath;
            }
            set
            {
                WriteDeportedAppFileKey("epubCheckerPath", value);
                pEpubCheckPath = value;
            }
        }

        private static bool pCheckPostGen;
        public static bool CheckPostGen
        {
            get
            {
                try
                {
                    string val = ReadDeportedAppFileKey("CheckPostGen");
                    if (String.IsNullOrEmpty(val))
                    {
                        MessageBox.Show("Mauvais fichier de configuration: " + SettingFile, "Message", MessageBoxButtons.OK);
                        return false;
                        //throw new Exception("Mauvais fichier de configuration: " + SettingFile);
                    }
                    pCheckPostGen = Convert.ToBoolean(val);
                }
                catch (Exception e)
                {
                    CheckPostGen = true;
                }
                return pCheckPostGen;
            }
            set
            {
                WriteDeportedAppFileKey("CheckPostGen", value.ToString());
                pCheckPostGen = value;
            }
        }


        private static bool pAutoRessources;
        public static bool AutoRessources
        {
            get
            {
                try
                {
                    string val = ReadDeportedAppFileKey("AutoRemplirRessources");
                    if (String.IsNullOrEmpty(val))
                    {
                        MessageBox.Show("Mauvais fichier de configuration: " + SettingFile, "Message", MessageBoxButtons.OK);
                        return false;
                        //throw new Exception("Mauvais fichier de configuration: " + SettingFile);
                    }
                    pAutoRessources = Convert.ToBoolean(val);
                }
                catch (Exception e)
                {
                    AutoRessources = true;
                }
                return pAutoRessources;
            }
            set
            {
                WriteDeportedAppFileKey("AutoRemplirRessources", value.ToString());
                pAutoRessources = value;
            }
        }

        private static bool pTocNcx;
        public static bool TocNcx
        {
            get
            {
                try
                {
                    string val = ReadDeportedAppFileKey("TocNcx");
                    if (String.IsNullOrEmpty(val))
                    {
                        MessageBox.Show("Mauvais fichier de configuration: " + SettingFile, "Message", MessageBoxButtons.OK);
                        return false;
                        //throw new Exception("Mauvais fichier de configuration: " + SettingFile);
                    } 
                    
                    pTocNcx = Convert.ToBoolean(val);
                }
                catch (Exception e)
                {
                    TocNcx = true;
                }
                return pTocNcx;
            }
            set
            {
                WriteDeportedAppFileKey("TocNcx", value.ToString());
                pTocNcx = value;
            }
        }


        protected static string ReadDeportedAppFileKey(string Key)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(SettingFile);
            XmlElement rootnode = doc.DocumentElement;
            if (!rootnode.Name.Contains("configuration"))
                throw new Exception("Mauvais fichier de configuration: " + SettingFile);
            
            XmlNode appSettingsNode = (from XmlNode node in rootnode.ChildNodes where node.Name.Contains("appSettings") select node).FirstOrDefault();
            if (appSettingsNode == null)
                throw new Exception("Mauvais fichier de configuration: " + SettingFile);

            return (from XmlElement elem in appSettingsNode.ChildNodes where elem.GetAttribute("key").Equals(Key) select elem.GetAttribute("value")).FirstOrDefault();
        }

        protected static bool WriteDeportedAppFileKey(string Key, string value)
        {
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(SettingFile);
                XmlElement rootnode = doc.DocumentElement;
                if (!rootnode.Name.Contains("configuration"))
                    throw new Exception("Mauvais fichier de configuration: " + SettingFile);

                XmlNode appSettingsNode = (from XmlNode node in rootnode.ChildNodes where node.Name.Contains("appSettings") select node).FirstOrDefault();
                if (appSettingsNode == null)
                    throw new Exception("Mauvais fichier de configuration: " + SettingFile);

                XmlElement paramNode = (from XmlElement elem in appSettingsNode.ChildNodes where elem.GetAttribute("key").Equals(Key) select elem).FirstOrDefault();
                if (paramNode == null)
                    throw new Exception("Mauvais fichier de configuration: " + SettingFile);

                paramNode.SetAttribute("value", value);
                doc.Save(SettingFile);
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
    }

}