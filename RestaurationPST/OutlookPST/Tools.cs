using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Permissions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Office.Core;
using System.Configuration;
using System.DirectoryServices;

namespace OutlookPST
{
    public class Tools
    {
        private static Thread ThreadDiscover { get; set; }

        public enum ChoixUseLocalPST
        {
            Oui,
            Non,
            Annule
        }

        public static string repo;
        public string localTmpPath;
        private const string AdrRepertoire = "Mes_Archives";
        private const string AdrLog = "Outlook_Mes_Archives.log";
        
       

        /// <summary>
        /// Création du repertoire pour la Log
        /// </summary>
        public static void CreerRepertoireLog()
        {
            try
            {
                repo = Path.Combine("C:\\Users\\Public\\BDF\\", AdrRepertoire);
                if (!Directory.Exists(repo))
                    Directory.CreateDirectory(repo);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Erreur lors de la création de la Log :{0}\n{1}\n{2}", ex.InnerException, ex.Source, ex.StackTrace), "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>Action de Log dans le ficher de log
        /// </summary> this.ctrler.Logger(string.Join("\r\n", msg));
        /// <param name="msg">Le message à écrire</param>
        public static void Logger(string msg)
        {
            try
            {
                // Création des répertoire pour la Log Si necessaire
                CreerRepertoireLog();

                DateTime datetime = DateTime.Now;
                if (!File.Exists(Path.Combine(repo, AdrLog)))
                {
                    System.IO.FileStream f = File.Create(Path.Combine(repo, AdrLog));
                    f.Close();
                }

                StreamWriter writter = File.AppendText(Path.Combine(repo, AdrLog));
                //writter.WriteLine("\n\r\n\r" + datetime.ToString("yyyy-MMM-dd HH:mm.fff") + " > " + msg);
                writter.WriteLine(datetime.ToString("yyyy-MMM-dd HH:mm.fff") + " > " + msg);
                writter.Flush();
                writter.Close();
                PerformFileTrim(Path.Combine(repo, AdrLog));
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Erreur lors de la création de la Log :{0}\n{1}\n{2}", ex.InnerException, ex.Source, ex.StackTrace), "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Formate les infos pour la Log
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="ex"></param>
        public static void LogMessage(string msg, System.Exception ex = null, bool AffichON = false, string methode="")
        {
            try
            {
                // Methode sollicité sur Erreur (Exception)
                if (ex != null)
                {
                    // Affiche un message is necessaire
                    if (AffichON)
                        System.Windows.Forms.MessageBox.Show(msg, string.Format("{0} - {1}", GetAppVersion(),"Erreur :"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    // Ajout A la log
                    Logger(string.Format("Méthode :{0}\r\n\"{1}\" \r\n Erreur :{2} \r\n StackTrace: {3}", methode, msg, ex.Message, ex.StackTrace));
                }
                else
                {
                    // Methode sollicité pour Afficher une information seulement
                    // Ajout A la log
                    Logger(string.Format("{0} - {1}", GetAppVersion(), msg));
                }
            }
            catch (Exception ex2)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Erreur PerformFiletri :{0}\n{1}\n{2}", ex2.InnerException, ex2.Source, ex2.StackTrace), "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }


        /// <summary>Action de réduction de la taille du fichier de log pour ne pas qu'il dépasse une certaine taille en MO
        /// </summary>
        /// <param name="filename">Emplacement du fichier de log.</param>
        private static void PerformFileTrim(string filename)
        {
            try
            {
                var FileSize = Convert.ToDecimal((new System.IO.FileInfo(filename)).Length);

                //10 megabytes maximum
                if (FileSize > 5000000)
                {
                    var file = File.ReadAllLines(filename).ToList();
                    var AmountToCull = (int)(file.Count * 0.33);
                    var trimmed = file.Skip(AmountToCull).ToList();
                    File.WriteAllLines(filename, trimmed);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(String.Format("Erreur PerformFiletri :{0}\n{1}\n{2}", ex.InnerException, ex.Source, ex.StackTrace), "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        

        public static void CurrentExplorer_Event()
        {
            try
            {
                

            }
            catch (Exception ex)
            {
                LogMessage("Erreur lors de l'actualisation." + MethodBase.GetCurrentMethod().Name, ex, true);
            }
        }


        /// <summary>
        /// Termine le thread passé en paramètre.
        /// </summary>
        /// <param name="thread"></param>
        [SecurityPermission(SecurityAction.Demand, ControlThread = true)]
        public static void KillThread(Thread thread)
        {
            if (thread != null)
            {
                if (thread.IsAlive)
                {
                    thread.Abort();
                }
            }
        }


        /// <summary>
        /// Test si l'office est 2016 (32b) ou 365 / 2019 (64b)
        /// </summary>
        /// <returns></returns>
        public static string OfficeVersion()
        {
            string version = "x64";

            // Recherche Si Office 2016 (32b) est sur le poste ?
            try
            {
                if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"))
                {
                    version = "x86";
                }
            }
            catch
            {

            }

            // Recherche Si Office 365 (64b) est sur le poste ?
            try
            {
                RegistryKey repcle = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common\Licensing\LicensingNext\", false);
                if (repcle != null)
                {
                    object clelicOff = repcle.GetValue("o365proplusretail");
                    if (clelicOff != null)
                    {
                        version = "x64";
                    }

                }
            }
            catch 
            {

            }

            // Sinon (64b) ?
            return version;
        }


        public static bool isConnected()
        {
            bool res = true;

            try
            {
                //System.Windows.Forms.MessageBox.Show(Globals.ThisAddIn.Application.Session.ExchangeConnectionMode.ToString());
                if (Globals.ThisAddIn.Application.Session.ExchangeConnectionMode==Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olCachedOffline || Globals.ThisAddIn.Application.Session.ExchangeConnectionMode == Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olOffline || Globals.ThisAddIn.Application.Session.ExchangeConnectionMode == Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olCachedDisconnected || Globals.ThisAddIn.Application.Session.ExchangeConnectionMode == Microsoft.Office.Interop.Outlook.OlExchangeConnectionMode.olDisconnected)
                {
                    res = false;
                }                
            }
            // Si non connecter ou impossible à recuperer
            catch (Exception exc)
            {
                System.Windows.Forms.MessageBox.Show(exc.Message);
                    res = false;
                // rien faire
            }
            return res;
        }


        /// <summary>
        /// En millisecondes, délai alloué à la découverte d'un réseau TCP/IP.
        /// Positionne une valeur par défaut si la valeur est absente ou illisible
        /// depuis le fichier de configuration (App.config).
        /// </summary>
        public static int NETWORK_DISCOVERY_TIME
        {
            get
            {
                int delay;

                try
                {
                    delay = 5000;
                }
                catch
                {
                    delay = 5000;
                }

                return delay;
            }
        }


        /// <summary>
        /// Méthode de découverte de connectivité au domaine ADBDF,
        /// appelée dans un thread séparé.
        /// </summary>
        private static void DiscoverNAS()
        {

            
        }


        public static bool testReseau()
        {
            bool res = true;
            int NetworksToDiscover = 4;

            try
            {

                ThreadDiscover = new Thread(start: new ThreadStart(DiscoverNAS));
                ThreadDiscover.Start();

                //  On attend tant que l'on n'a pas vérifié la joignabilité de tous les annuaires.
                //  Le temps d'attente est celui alloué à la découverte, augmenté de 10 %.
                int discoverGlobalTimeOut = 0;  //  délai incrémenté en millisecondes
                int stepDelay = 10;             //  délai de rafraîchissement en millisecondes

                while (NetworksToDiscover != 0 && discoverGlobalTimeOut < NETWORK_DISCOVERY_TIME * 1.10)
                {
                    Thread.Sleep(millisecondsTimeout: stepDelay);

                    discoverGlobalTimeOut += stepDelay;
                }
                KillThread(thread: ThreadDiscover);
                ThreadDiscover = null;
            }
            // Si non connecter ou impossible à recuperer
            catch 
            {
                //System.Windows.Forms.MessageBox.Show(ex.Message);
                res = false;
                // rien faire
            }
            return res;
        }

        public static string GetAppVersion()
        {
            Version ver = null;

            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)

                ver = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion;

            else

                ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

            return string.Format("{0}", ver);
            //return string.Format("{0}/{1}", " V:", ver);
        }
    }
}
