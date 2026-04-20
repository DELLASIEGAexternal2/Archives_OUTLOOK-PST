using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices;
using System.Drawing.Printing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookPST
{

    public static class AD
    {
        //public const string AD_FQDN_PROD = "ADBDF.PRIVATE";
        public const string AD_PROD = "ADBDF";
        //public const string AD_FQDN_TEST = "ADBDFINT.PRIVATE";
        public const string AD_TEST = "ADBDFINT";


        static AD() {}

        /// <summary>
        /// Recupération de l'AD du samAccountName en Minuscule
        /// </summary>
        /// <param name="mailBox"></param>
        /// <returns></returns>
        public static string getsamAccountName(string mailBox)
        {
            string samAccountName = "";
            DirectoryEntry Ldap = null;
            try
            {
                if (!Tools.isConnected())
                {
                    //System.Windows.Forms.MessageBox.Show("Fonction non disponible en mode Hors ligne.","Mode Hors Ligne",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    return samAccountName;
                }


                string userLogon = Environment.GetEnvironmentVariable(variable: "Username");
                if (Environment.GetEnvironmentVariable(variable: "USERDOMAIN").ToUpper()== AD_TEST)
                {
                     Ldap = new DirectoryEntry("LDAP://ADBDFINT.PRIVATE", userLogon,null,AuthenticationTypes.Secure);
                }
                else {
                     Ldap = new DirectoryEntry("LDAP://ADBDF.PRIVATE", userLogon, null, AuthenticationTypes.Secure);
                }

                if (Ldap == null) return null;

                string LdapFilter = "(ANR="+mailBox+")";

                //  Instancie le DirectorySearcher
                DirectorySearcher ds = new DirectorySearcher()
                {
                    SearchRoot = Ldap,
                    Filter = LdapFilter,
                    SearchScope = SearchScope.Subtree,
                    PageSize = 1,
                    ServerTimeLimit = new TimeSpan(hours: 0, minutes: 0, seconds: 10)  
                };

                SearchResult resultds = ds.FindOne();
                if (resultds != null)
                { 
                DirectoryEntry DirEntry = resultds.GetDirectoryEntry();
                samAccountName = DirEntry.Properties["sAMAccountName"].Cast<string>().First();
                samAccountName= samAccountName.ToLower();
                }


            }
            catch (System.Exception ex)
            {
                Tools.LogMessage("Erreur lors de la récupération de la samAccountName de la Boite aux lettres selectionnée." + MethodBase.GetCurrentMethod().Name, ex, true);
                samAccountName = "";
            }
            return samAccountName;
        }

        /// <summary>
        /// Recupération de l'AD du samAccountName en Minuscule
        /// </summary>
        /// <param name="mailBox"></param>
        /// <returns></returns>
        public static string getmailNickName(string mailBox)
        {
            string mailNickName = "";
            DirectoryEntry Ldap = null;
            try
            {
                if (!Tools.isConnected())
                {
                    //System.Windows.Forms.MessageBox.Show("Fonction non disponible en mode Hors ligne.","Mode Hors Ligne",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    return mailNickName;
                }


                string userLogon = Environment.GetEnvironmentVariable(variable: "Username");
                if (Environment.GetEnvironmentVariable(variable: "USERDOMAIN").ToUpper() == AD_TEST)
                {
                    Ldap = new DirectoryEntry("LDAP://ADBDFINT.PRIVATE", userLogon, null, AuthenticationTypes.Secure);
                }
                else
                {
                    Ldap = new DirectoryEntry("LDAP://ADBDF.PRIVATE", userLogon, null, AuthenticationTypes.Secure);
                }

                if (Ldap == null) return null;

                string LdapFilter = "(ANR=" + mailBox + ")";

                //  Instancie le DirectorySearcher
                DirectorySearcher ds = new DirectorySearcher()
                {
                    SearchRoot = Ldap,
                    Filter = LdapFilter,
                    SearchScope = SearchScope.Subtree,
                    PageSize = 1,
                    ServerTimeLimit = new TimeSpan(hours: 0, minutes: 0, seconds: 10)
                };

                SearchResult resultds = ds.FindOne();
                if (resultds != null)
                {
                    DirectoryEntry DirEntry = resultds.GetDirectoryEntry();
                    mailNickName = DirEntry.Properties["mailNickName"].Cast<string>().First();
                }


            }
            catch (System.Exception ex)
            {
                Tools.LogMessage("Erreur lors de la récupération de la mailNickName de la Boite aux lettres selectionnée." + MethodBase.GetCurrentMethod().Name, ex, true);
                mailNickName = "";
            }
            return mailNickName;
        }

        /// <summary>
        /// getInfoRetention => Fixe la date de debut pour les retentions VIP/10ans et Defaut
        /// </summary>
        /// <param name="mailBox"></param>
        /// <returns>No error</returns>
        public static bool getInfoRetention(string mailBox)
        {
            DirectoryEntry Ldap = null;
            try
            {
                // plus necessaire
                //if (!Tools.isConnected())
                //{
                //    //System.Windows.Forms.MessageBox.Show("Fonction non disponible en mode Hors ligne.","Mode Hors Ligne",MessageBoxButtons.OK,MessageBoxIcon.Information);
                //    return res;
                //}


                string userLogon = Environment.GetEnvironmentVariable(variable: "Username");
                if (Environment.GetEnvironmentVariable(variable: "USERDOMAIN").ToUpper() == AD_TEST)
                {
                    Ldap = new DirectoryEntry("LDAP://ADBDFINT.PRIVATE", userLogon, null, AuthenticationTypes.Secure);
                }
                else
                {
                    Ldap = new DirectoryEntry("LDAP://ADBDF.PRIVATE", userLogon, null, AuthenticationTypes.Secure);
                }

                if (Ldap == null) return false;

                string LdapFilter = "(ANR=" + mailBox + ")";

                //  Instancie le DirectorySearcher
                DirectorySearcher ds = new DirectorySearcher()
                {
                    SearchRoot = Ldap,
                    Filter = LdapFilter,
                    SearchScope = SearchScope.Subtree,
                    PageSize = 1,
                    ServerTimeLimit = new TimeSpan(hours: 0, minutes: 0, seconds: 10)
                };

                SearchResult resultds = ds.FindOne();
                if (resultds != null)
                {
                    DirectoryEntry DirEntry = resultds.GetDirectoryEntry();
                    string retention = DirEntry.Properties["msExchMailboxTemplateLink"].Cast<string>().First();
                    if (retention != null)
                    {
                        // Ruban 10 ans pour les rétentions vip/10ans/defaut 
                        if ((retention.ToLower().Contains("vip") || retention.ToLower().Contains("10ans") || retention.ToLower().Contains("defaut")))
                        {
                            // Fixe les infos pour le ruban
                            Traitement.majAnneeDebutRuban(true);                           
                        }
                        else
                        {
                            Traitement.majAnneeDebutRuban(false);
                        }
                    }
                }

            }
            catch (System.Exception ex)
            {
                //Tools.LogMessage("Erreur lors de la récupération du type de rétention de la Boite aux lettres selectionnée." + MethodBase.GetCurrentMethod().Name, ex, false);
                Traitement.majAnneeDebutRuban(false);
                return false;
            }
            return true;
        }

    }
}
