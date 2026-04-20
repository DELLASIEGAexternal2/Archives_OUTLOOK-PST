using OutlookPST.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPST
{
    public static class NASClass
    {
        /// <summary>
        /// Récupération des clés dustockage StocObj
        /// </summary>
        /// <param name="usInfo"></param>
        /// <param name="userDemande">Utilisateur qui fair la demande</param>
        /// <returns></returns>
        public static bool GetNASInformation(StocObjInfoUser usInfo,string userDemande)
        {
            try
            {
                Tools.LogMessage(string.Format("NAS : Récupération des clés de l'utilisateur :{0}",userDemande), null, false, MethodBase.GetCurrentMethod().Name);
                //string fullPathFileNas = Path.Combine(Const.pathNAS, "_AccessKeys", Const.userlogon, string.Format("{0}-akey.txt", Const.userlogon));
                string fullPathFileNas = Path.Combine(Const.pathNAS, "_AccessKeys", userDemande, string.Format("{0}-akey.txt", userDemande));
                if (!File.Exists(fullPathFileNas))
                {
                    return false;
                }


                // Lecture de toutes les lignes
                string[] lignesDuFichier = File.ReadAllLines(fullPathFileNas);

                //  Dernière ligne
                string derniereLigne = lignesDuFichier[lignesDuFichier.Length - 1];
                // Décomposition / Separateur TAB (9)
                string[] elem = derniereLigne.Split(Convert.ToChar(9));

                // Structure a ce jours
                string accesKey = ValSansGuillement(elem[0]);
                string createDate = ValSansGuillement(elem[1]);
                string secretKey = ValSansGuillement(elem[2]);
                string status = ValSansGuillement(elem[3]);
                string username = ValSansGuillement(elem[4]);

                // Si la derniere ligne n'est pas avec un status actif et pour le meme matricule
                //if (username.ToLower() == Const.userlogon.ToLower())
                if (username.ToLower() == userDemande.ToLower())
                {
                    usInfo.StocObjAccessKey = accesKey;
                    usInfo.StocObjSecretKey = secretKey;
                }
                else
                {
                    // Si la derniere ligne n'est pas le meme matricule
                    return false;
                }

                return true;
            }
            catch (System.Exception ex)
            {
                Tools.LogMessage("Erreur d'accès au NAS.", ex, false, MethodBase.GetCurrentMethod().Name);
                return false;
            }

        }

        private static string ValSansGuillement(string s)
        {
            string s2 = s;
            try
            {
                if (!string.IsNullOrEmpty(s) && s.StartsWith("\"") && s.EndsWith("\""))
                {
                    s2 = s.Substring(1, s.Length - 2);
                }
            }
            catch 
            {

            }
            return s2;
        }

    }
}
