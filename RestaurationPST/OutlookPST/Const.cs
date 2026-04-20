
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookPST
{
    

    public static class Const
    {
        // limit du nombre de PST installable
        public static int limitNbPST = 15;

        // Info pour le ruban : Represente la bal selectionné
        // Par default
        public static bool isTenYear = false; 
        public static int anneeDebut = DateTime.Now.AddYears(-5).Year; 

        // Indicateur d'annulation
        // **** Annulation des téléchargements
        public static CancellationTokenSource cancellationTokenSourceForDownloads;
        public static bool stopThread = false;

        public static bool downloadsInUseThread = false;
        public static int tempo = 30;

        // Toast Message
        public static string ToastsMessage = "";
        // Background worker pour Toast 
        public static System.ComponentModel.BackgroundWorker worker;

        // Utilisateur Connecté
        public static string userlogon = Environment.GetEnvironmentVariable(variable: "Username").ToLower();
        // ****** Chemin pour les PST *****
        public static string path_Local_PST = "";
        // Chemin du Flag pour le fichier PST -> C:\Users\{0}\Documents confidentiels (local)\PST\
        public static string path_Local_Flg_PST = "";
        // Chemin du Flag pour l'indicateur du RUBAN pour le PST -> ex: C:\Users\{0}\Documents confidentiels (local)\PST\outlook\
        public static string path_Local_Flg_For_Session = "";


        // Utiliser pour l'actualisation des bt de la boite séléctionné 
        // Permet de limiter l'actualisation au changement de boite
        public static string currentMailBox = "";

        public static List<FlgPSTClass> lstFlagPST = new List<FlgPSTClass>();
        public static List<FlgRubanClass> lstFlagRuban = new List<FlgRubanClass>();

        public static ParamDemande param = new ParamDemande();        

        //**************************************
        //Information pour AWS S3
        //**************************************
        public static string pathNAS = Environment.GetEnvironmentVariable(variable: "Userdomain").ToLower()=="adbdfint" ? ConfigurationManager.AppSettings[name: "pathNAS_IN"].ToString() : ConfigurationManager.AppSettings[name: "pathNAS_PR"].ToString();
        public static string stockObj_ServiceURL = Environment.GetEnvironmentVariable(variable: "Userdomain").ToLower() == "adbdfint" ? ConfigurationManager.AppSettings[name: "stockObj_ServiceURL_IN"].ToString() : ConfigurationManager.AppSettings[name: "stockObj_ServiceURL_PR"].ToString();
        public static string stockObj_RegionEndpoint = Environment.GetEnvironmentVariable(variable: "Userdomain").ToLower() == "adbdfint" ? ConfigurationManager.AppSettings[name: "stockObj_RegionEndpoint_IN"].ToString() : ConfigurationManager.AppSettings[name: "stockObj_RegionEndpoint_PR"].ToString();
        public static string stockObj_BucketName = Environment.GetEnvironmentVariable(variable: "Userdomain").ToLower() == "adbdfint" ? ConfigurationManager.AppSettings[name: "stockObj_BucketName_IN"].ToString() : ConfigurationManager.AppSettings[name: "stockObj_BucketName_PR"].ToString();

        public static string tmp7z = string.Format(ConfigurationManager.AppSettings[name: "tmp7z"], Environment.GetEnvironmentVariable(variable: "username").ToLower());

        public static List<string> lstPST_InSession = new List<string>();

        public static string publicBloc= "iNEQ";
        public static string finBloc = "c$v";

        public static bool isConnected;

        public static string GetAccueil
        {
            get
            {
                return ConfigurationManager.AppSettings["Accueil"];
            }
        }

        public static string GetModop
        {
            get
            {
                return ConfigurationManager.AppSettings["Modop"];
            }
        }
    }
}
