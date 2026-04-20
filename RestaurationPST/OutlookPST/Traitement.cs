using Amazon.Auth.AccessControlPolicy;
using Amazon.S3;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Win32;
using OutlookPST.Model;
using SevenZipExtractor;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OutlookPST.Tools;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace OutlookPST
{
    public static class Traitement
    {
        /// INIT : Création des Dossiers pour les fichier PST et FLAG du ruban
        /// <summary>
        /// createFolderForPSTandFlag : Donne le chemin des PST et flag
        /// </summary>
        public static void createFolderForPSTandFlag()
        {
            //string cheminLocal = "";
            string path_Racine = "";
            try
            {

                // Chemin dans Document Confidentiels > C:\\Users\\{0}\\Documents confidentiels (local)\\
                path_Racine = string.Format(ConfigurationManager.AppSettings["cheminRacinePST"], Environment.GetEnvironmentVariable(variable: "Username"));
                if (!Directory.Exists(path_Racine))
                {
                    // Création nécessite un compte OT2
                    //Directory.CreateDirectory(path_Racine);
                    System.Windows.Forms.MessageBox.Show(string.Format("Répertoire inexistant :\n{0}\n\nVeuillez demander la création à votre correspondant informatique.", path_Racine), string.Format("{0} - {1}", GetAppVersion(),"Poste non conforme :"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
                // Si il n'existe pas Outlook Absent
                if (Directory.Exists(path_Racine))
                {
                    Const.path_Local_PST = Path.Combine(path_Racine, @"PST\");
                    Const.path_Local_Flg_PST = Const.path_Local_PST;
                    // Création du repertoire de stockage des PST et flgPST si il n'existe pas
                    if (!Directory.Exists(Const.path_Local_PST))
                        Directory.CreateDirectory(Const.path_Local_PST);

                    string profileOutlook = Globals.ThisAddIn.Application.Session.CurrentProfileName;
                    string ssDossierProfileOutlook = Path.Combine(Const.path_Local_PST, profileOutlook);
                    // Création du repertoire de stockage des flg de DownLoad si il n'existe pas
                    if (!Directory.Exists(ssDossierProfileOutlook))
                        Directory.CreateDirectory(ssDossierProfileOutlook);
                    Const.path_Local_Flg_For_Session = ssDossierProfileOutlook;
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors de la création des répertoires dans :/n{0}", path_Racine), ex, true, MethodBase.GetCurrentMethod().Name);
                System.Windows.Forms.MessageBox.Show(String.Format(@"Erreur lors de la création du répertoire dans ""{0}"" de stockage des PST :{1}\n{2}\n{3}", path_Racine + @"PST", ex.InnerException, ex.Source, ex.StackTrace), string.Format("{0} - {1}", GetAppVersion(),"Erreur"), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }


        public static bool initDemande(string bt, ParamDemande param)
        {
            try
            {
                // Recupération de la Current MailBox 
                param.btDemande = Traitement.getCurrentMailBox();
                if (string.IsNullOrEmpty(param.btDemande))
                    return false;                

                // Recupération de la Current samAccountName de la MailBox 
                param.btDemande_samAccountName = AD.getsamAccountName(param.btDemande);
                if (string.IsNullOrEmpty(param.btDemande_samAccountName))
                {
                    if (Tools.isConnected())
                    {
                        //System.Windows.Forms.MessageBox.Show("SamAccount de la boîte non résolu :" + param.btDemande, "Erreur dans la recherche AD", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return false;
                    }
                    else
                    {
                        // retourn true car pas necessaire en OffLigne
                        //return true;
                    }
                }

                // Recupération de la Current mailNickName de la MailBox 
                param.btDemande_mailNickName = AD.getmailNickName(param.btDemande);

                // Recupération de l'Année du bouton 
                param.anneeDemande = getAnneeBtn(bt);
                if (string.IsNullOrEmpty(param.anneeDemande))
                    return false;                

            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de l'initialisation de la demande.", ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Retourne le nom de l'Annee en fonction du bouton
        /// </summary>
        /// <param name="bt"></param>
        /// <returns></returns>
        private static string getAnneeBtn(string bt)
        {
            string Annee = "";
            try
            {
                // Annee
                switch (bt)
                {
                    case "btn1":
                        Annee = "2015";
                        break;
                    case "btn2":
                        Annee = "2016";
                        break;
                    case "btn3":
                        Annee = "2017";
                        break;
                    case "btn4":
                        Annee = "2018";
                        break;
                    case "btn5":
                        Annee = "2019";
                        break;
                    case "btn6":
                        Annee = "2020";
                        break;
                    case "btn7":
                        Annee = "2021";
                        break;
                    case "btn8":
                        Annee = "2022";
                        break;
                    case "btn9":
                        Annee = "2023";
                        break;
                    case "btn10":
                        Annee = "2024";
                        break;
                    default:
                        Annee = "2029";
                        break;
                }

            }
            catch 
            {
            }
            return Annee;
        }

        /// <summary>
        /// Récupération des informations du fichier statut.flg de 10ans pour le ruban (Dans le PST\Nom_Profil)
        /// </summary>
        /// <param name="btDemande">Boite Demandé</param>
        /// <param name="lstFlagDeStatutRuban">Récupération des Status sur les boites demandés</param>
        /// <returns>As t'on l'info de rentention dans le fichier Ruban</returns>
        public static bool initTenYearByFlagRuban(string btDemande)
        {
            bool res = false;
            try
            {
                if (string.IsNullOrEmpty(btDemande))
                    return false;

                Const.path_Local_PST=Path.Combine(string.Format(ConfigurationManager.AppSettings["cheminRacinePST"], Environment.GetEnvironmentVariable(variable: "Username")), @"PST\");
                Const.path_Local_Flg_For_Session = Path.Combine(Const.path_Local_PST, Globals.ThisAddIn.Application.Session.CurrentProfileName); ;
                string pathFichierFlgRuban = Path.Combine(Const.path_Local_Flg_For_Session, getNameFlg(btDemande));

                // Si il exist je le remmplace
                if (File.Exists(pathFichierFlgRuban))
                {
                    string[] lines = System.IO.File.ReadAllLines(pathFichierFlgRuban);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (i == 0)
                        {
                            string[] dt = lines[i].Split(';');
                            if (!string.IsNullOrEmpty(dt[0]))
                            {
                                if (dt.Count() > 4)
                                {
                                    if (dt[4] == "10ans")
                                    {
                                        majAnneeDebutRuban(true);
                                        return true;
                                    }
                                    else
                                    {
                                        majAnneeDebutRuban(false);
                                        return true;
                                    }
                                }
                                else
                                {
                                    // Pas d'info de retention dans le fichier
                                    return false;
                                }
                            }
                            break;
                        }
                    }
                }
                else
                {
                    // Pas encors de fichier de statut flag pour le ruban
                    return false;
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors du chargement des flag ruban afin de trouver l'info 10ans pour la boite {0} dans le profil {1}. \n", btDemande, Const.path_Local_Flg_For_Session), ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            return true;
        }


        /// <summary>
        /// Récupération des informations du fichier statut.flg pour le ruban (Dans le PST\Nom_Profil)
        /// </summary>
        /// <param name="btDemande"></param>
        /// <param name="lstFlagDeStatutRuban"></param>
        /// <returns></returns>
        public static bool readFlagRuban(string btDemande, List<FlgRubanClass> lstFlagDeStatutRuban)
        {
            try
            {
                if (string.IsNullOrEmpty(btDemande))
                    return false;

                string pathFichierFlgRuban = Path.Combine(Const.path_Local_Flg_For_Session, getNameFlg(btDemande));

                // Si il exist je le remmplace
                if (File.Exists(pathFichierFlgRuban))
                {
                    lstFlagDeStatutRuban.Clear();   
                    string[] lines = System.IO.File.ReadAllLines(pathFichierFlgRuban); 
                    for (int i = 0; i < lines.Length; i++) 
                    {
                        if (i == 0)
                        {
                            string[] dt = lines[i].Split(';');
                            if (!string.IsNullOrEmpty(dt[0]))
                            {
                                if (dt.Count() > 4)
                                {
                                    // Force que lorsque c'est 10 ans
                                    if (dt[4] == "10ans")
                                    {
                                        majAnneeDebutRuban(true);
                                    }
                                    //else
                                    //    majAnneeDebutRuban(false);
                                }
                                //else
                                //{
                                //    majAnneeDebutRuban(false);
                                //}
                            }
                        }


                        if (i>=1)
                        {
                        string[] dt=lines[i].Split(';');
                            if (!string.IsNullOrEmpty(dt[0]))
                            {
                                FlgRubanClass newFlgRuban= new FlgRubanClass();                            
                                newFlgRuban.MailBoxName = btDemande.Trim();
                                newFlgRuban.Annee = dt[0].Trim();
                                newFlgRuban.Archive = dt[1]=="X"?true:false;
                                newFlgRuban.Statut = convertFlagRuban(dt[2]);
                                newFlgRuban.Commentaire= dt[3].Trim();
                                lstFlagDeStatutRuban.Add(newFlgRuban);
                            }
                        }
                    }                    
                }
                else
                {
                    // Pas encors de fichier de statut flag pour le ruban
                    return false;
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors du chargement des flag ruban pour la boite {0} dans le profil {1}. \n", btDemande, Const.path_Local_Flg_For_Session), ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Recuperation des information dans le fichier statut .flg pour les PST
        /// </summary>
        /// <param name="btDemande"></param>
        /// <param name="lstFlagDeStatutPST"></param>
        /// <returns></returns>
        public static bool readFlagPST(string btDemande, List<FlgPSTClass> lstFlagDeStatutPST)
        {
            try
            {
                if (string.IsNullOrEmpty(btDemande))
                    return false;

                string pathFichierFlgRuban = Path.Combine(Const.path_Local_Flg_PST, getNameFlg(btDemande));

                // Si il exist je le remmplace
                if (File.Exists(pathFichierFlgRuban))
                {
                    lstFlagDeStatutPST.Clear();
                    string[] lines = System.IO.File.ReadAllLines(pathFichierFlgRuban);
                    for (int i = 0; i < lines.Length; i++)
                    {
                        if (i >= 1)
                        {
                        string[] dt = lines[i].Split(';');
                            if (!string.IsNullOrEmpty(dt[0]))
                            {
                                FlgPSTClass newFlgRuban = new FlgPSTClass();
                                newFlgRuban.MailBoxName = btDemande.Trim();
                                newFlgRuban.Annee = dt[0].Trim();
                                newFlgRuban.Archive = dt[1] == "X" ? true : false;
                                newFlgRuban.Statut = convertFlagPST(dt[2]);
                                newFlgRuban.Commentaire = dt[3].Trim();
                                lstFlagDeStatutPST.Add(newFlgRuban);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors du chargement des flag ruban pour la boite {0} dans le profil {1}. \n", btDemande, Const.path_Local_Flg_For_Session), ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Renvoi le statut du flag Ruban
        /// </summary>
        /// <param name="flg"></param>
        /// <returns></returns>
        private static Statut_ruban convertFlagRuban(string flg)
        {
            Statut_ruban res= Statut_ruban.None; 
            try
            {
                switch (flg.ToLower())
                {
                    case "none":
                        res = Statut_ruban.None;
                        break;

                    case "ok":
                        res = Statut_ruban.OK;
                        break;

                    case "hs":
                        res = Statut_ruban.HS;
                        break;

                    case "encours":
                        res = Statut_ruban.Encours;
                        break;

                    case "mount":
                        res = Statut_ruban.Mount;
                        break;

                    case "indisponible":
                        res = Statut_ruban.Indisponible;
                        break;

                    case "annule":
                        res = Statut_ruban.Annule;
                        break;

                    default:
                        res = Statut_ruban.None;
                        break;
                }
            }
            catch {
                res = Statut_ruban.None;
            }
            return res;
        }

        /// <summary>
        /// Renvoi un status pour le Flag PST
        /// </summary>
        /// <param name="flg"></param>
        /// <returns></returns>
        private static Statut_PST convertFlagPST(string flg)
        {
            Statut_PST res = Statut_PST.None;
            try
            {
                switch (flg.ToLower())
                {
                    case "none":
                        res = Statut_PST.None;
                        break;
            
                    case "hs":
                        res = Statut_PST.HS;
                        break;

                    case "encours":
                        res = Statut_PST.Encours;
                        break;

                    case "mount":
                        res = Statut_PST.Mount;
                        break;

                    case "indisponible":
                        res = Statut_PST.Indisponible;
                        break;

                    case "annule":
                        res = Statut_PST.Annule;
                        break;

                    default:
                        res = Statut_PST.None;
                        break;
                }
            }
            catch
            {
                res = Statut_PST.None;
            }
            return res;
        }

        /// <summary>
        /// Création ou Modification du Flag du Ruban
        /// </summary>
        /// <param name="lstFlagDeStatutRuban"></param>
        /// <param name="btSouhaite"></param>
        /// <param name="anneeSouhaite"></param>
        /// <param name="archiveSouhaite"></param>
        /// <param name="statut_souhaite"></param>
        /// <param name="commentaire_souhaite"></param>
        /// <returns></returns>
        public static bool createOrUpdateFlagRuban(List<FlgRubanClass> lstFlagDeStatutRuban, string btSouhaite, string anneeSouhaite, bool archiveSouhaite, Statut_ruban statut_souhaite, string commentaire_souhaite = "")
        {
            bool delete = false;
            try
            {
                string pathFichierFlgRuban = getFullPathFlg(btSouhaite);

                // Si il exist je le remmplace
                if (IsExistFlginLocal(btSouhaite))
                {
                    // Recupération des statut du ruban
                    readFlagRuban(btSouhaite, lstFlagDeStatutRuban);
                    delete = true;


                    // Mise a jours de la donnée
                    FlgRubanClass res = lstFlagDeStatutRuban.FirstOrDefault(c => c.MailBoxName.ToLower() == btSouhaite.ToLower() && c.Annee == anneeSouhaite && c.Archive == archiveSouhaite);
                    if (res != null)
                    {
                        res.Statut = statut_souhaite;
                        res.Commentaire = commentaire_souhaite;
                    }
                    else
                    {
                        // Ajout
                        FlgRubanClass dataF = new FlgRubanClass();
                        dataF.MailBoxName = btSouhaite;
                        dataF.Annee = anneeSouhaite;
                        dataF.Archive = archiveSouhaite;
                        dataF.Statut = statut_souhaite;
                        dataF.Commentaire = commentaire_souhaite;
                        lstFlagDeStatutRuban.Add(dataF);
                    }

                    if (delete) File.Delete(pathFichierFlgRuban);

                    // Sinon Création du fichier               
                    StreamWriter sw = new StreamWriter(pathFichierFlgRuban);

                    string dataFlg = "Annee;Archive;Statut;Commentaire";                    
                    if (Const.isTenYear)
                    {
                        dataFlg = "Annee;Archive;Statut;Commentaire;10ans";
                    }
                    sw.WriteLine(dataFlg);

                    foreach (FlgRubanClass itemFlgRuban in lstFlagDeStatutRuban)
                    {
                        dataFlg = string.Format("{0};{1};{2};{3}", itemFlgRuban.Annee, itemFlgRuban.Archive ? "X" : " ", itemFlgRuban.Statut.ToString(), itemFlgRuban.Commentaire);
                        sw.WriteLine(dataFlg);
                    }
                    sw.Close();
                }
                else
                {
                    lstFlagDeStatutRuban.Clear();
                    FlgRubanClass dataF = new FlgRubanClass();
                    dataF.MailBoxName = btSouhaite;
                    dataF.Annee = anneeSouhaite;
                    dataF.Archive = archiveSouhaite;
                    dataF.Statut = statut_souhaite;
                    dataF.Commentaire = commentaire_souhaite;
                    lstFlagDeStatutRuban.Add(dataF);

                    StreamWriter sw = new StreamWriter(pathFichierFlgRuban);
                    string dataFlg = "Annee;Archive;Statut;Commentaire";
                    if (Const.isTenYear)
                    {
                        dataFlg = "Annee;Archive;Statut;Commentaire;10ans";
                    }
                    sw.WriteLine(dataFlg);

                    foreach (FlgRubanClass itemFlgRuban in lstFlagDeStatutRuban)
                    {
                        dataFlg = string.Format("{0};{1};{2};{3}", itemFlgRuban.Annee, itemFlgRuban.Archive ? "X" : " ", itemFlgRuban.Statut.ToString(), itemFlgRuban.Commentaire);
                        sw.WriteLine(dataFlg);
                    }
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors de la création ou modification du flag ruban pour la boite {0} dans le profil {1}. \n", btSouhaite, Const.path_Local_Flg_For_Session), ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            return true;
        }


        /// <summary>
        /// Création ou Modification du Flag pour le PST
        /// </summary>
        /// <param name="lstFlagPST"></param>
        /// <param name="btDemande"></param>
        /// <param name="AnneeSouhaite"></param>
        /// <param name="ArchiveSouhaite"></param>
        /// <param name="statut_souhaite"></param>
        /// <param name="commentaire_souhaite"></param>
        /// <returns></returns>
        public static bool createOrUpdateFlagPST(List<FlgPSTClass> lstFlagPST, string btDemande, string AnneeSouhaite, bool ArchiveSouhaite, Statut_PST statut_souhaite, string commentaire_souhaite = "")
        {
            bool delete = false;
            try
            {
                string pathFichierFlgPST = getFullPathFlgPST(btDemande);

                // Si il exist je le remmplace
                if (IsExistFlgPSTinLocal(btDemande))
                {
                    // Recupération des statut du ruban
                    readFlagPST(btDemande, lstFlagPST);
                    delete = true;

                    // Mise a jours de la donnée
                    FlgPSTClass res = lstFlagPST.FirstOrDefault(c => c.MailBoxName.ToLower() == btDemande.ToLower() && c.Annee == AnneeSouhaite && c.Archive == ArchiveSouhaite);
                    if (res != null)
                    {
                        res.Statut = statut_souhaite;
                        res.Commentaire = commentaire_souhaite;
                    }
                    else
                    {
                        // Add 
                        FlgPSTClass dataF = new FlgPSTClass();
                        dataF.MailBoxName = btDemande;
                        dataF.Annee = AnneeSouhaite;
                        dataF.Archive = ArchiveSouhaite;
                        dataF.Statut = statut_souhaite;
                        dataF.Commentaire = commentaire_souhaite;
                        lstFlagPST.Add(dataF);
                    }
                    // Sinon Création du fichier               
                    if (delete) File.Delete(pathFichierFlgPST);
                    StreamWriter sw = new StreamWriter(pathFichierFlgPST);
                    string dataFlg = "Annee;Archive;Statut;Commentaire";
                    sw.WriteLine(dataFlg);

                    foreach (FlgPSTClass itemFlgPST in lstFlagPST)
                    {
                        dataFlg = string.Format("{0};{1};{2};{3}", itemFlgPST.Annee, itemFlgPST.Archive ? "X" : " ", itemFlgPST.Statut.ToString(), itemFlgPST.Commentaire);
                        sw.WriteLine(dataFlg);
                    }
                    sw.Close();
                }
                else
                {
                    lstFlagPST.Clear();
                    FlgPSTClass dataF = new FlgPSTClass();
                    dataF.MailBoxName = btDemande;
                    dataF.Annee = AnneeSouhaite;
                    dataF.Archive = ArchiveSouhaite;
                    dataF.Statut = statut_souhaite;
                    dataF.Commentaire = commentaire_souhaite;
                    lstFlagPST.Add(dataF);

                    StreamWriter sw = new StreamWriter(pathFichierFlgPST);

                    string dataFlg = "Annee;Archive;Statut;Commentaire";
                    sw.WriteLine(dataFlg);

                    foreach (FlgPSTClass itemFlgPst in lstFlagPST)
                    {
                        dataFlg = string.Format("{0};{1};{2};{3}", itemFlgPst.Annee, itemFlgPst.Archive ? "archive" : "", itemFlgPst.Statut.ToString(), itemFlgPst.Commentaire);
                        sw.WriteLine(dataFlg);
                    }
                    sw.Close();
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors de la création ou modification du flag ruban pour la boite {0} dans le profil {1}. \n", btDemande, Const.path_Local_Flg_For_Session), ex, true, MethodBase.GetCurrentMethod().Name);
                return false;
            }
            return true;
        }

        /// <summary>
        /// Message sur l'utilisation du fichier PST local
        /// </summary>
        /// <returns></returns>
        private static DialogResult msgChoixUsePSTOffLine()
        {
            DialogResult choix = DialogResult.No;
            try
            {
                choix = System.Windows.Forms.MessageBox.Show(MessageClass.message_ChoixOffLigne, string.Format("{0} - {1}", GetAppVersion(),"Remonter les données :"), System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            }
            catch
            {

            }
            return choix;
        }

        /// <summary>
        /// Message sur l'utilisation du fichier PST local
        /// </summary>
        /// <returns></returns>
        private static DialogResult msgChoixUsePST()
        {
            DialogResult choix = DialogResult.No;
            try
            {             
                choix = System.Windows.Forms.MessageBox.Show(MessageClass.message_Choix, string.Format("{0} - {1}", GetAppVersion(),"Télécharger les données ou remonter les données :"), System.Windows.Forms.MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);
            } 
            catch 
            {

            }
            return choix;
        }


        /// <summary>
        /// Monte le PST Local si il est présent suivant le choix Utilisteur
        /// Appelé pour installer le PST a la fin du téléchargement sans choix utilisateur. Via InstallForced
        /// </summary>
        /// <param name="btDemande"></param>
        /// <param name="annee"></param>
        /// <param name="InstallForced">True lors d'un appel en fin de téléchargement</param>
        /// <returns></returns>
        public static ChoixUseLocalPST useLocalPST(string btDemande, string annee, bool InstallForced=false) //, List<FlgPSTClass> lstFlagPST, ParamDemande param, Statut_PST statut_souhaite, string commentaire_souhaite = "")
        {
            List<FlgPSTClass> lstFlagPST=new List<FlgPSTClass>();
            List<FlgRubanClass> lstFlagRuban =new List<FlgRubanClass>();
            ParamDemande param=new ParamDemande();

            Statut_ruban statut_Ruban_BT = new Statut_ruban();
            Statut_ruban statut_Ruban_BT_Archive = new Statut_ruban();
            string commentaire_BT = "";
            string commentaire_BT_Archive = "";

            Statut_PST statut_Ruban_BT_PST = new Statut_PST();
            Statut_PST statut_Ruban_BT_Archive_PST = new Statut_PST(); ;
            string commentaire_BT_PST = "";
            string commentaire_BT_Archive_PST = "";

            bool resMountBT=false;
            bool resMountBT_Archive = false;

            ChoixUseLocalPST choixUtilisateur = ChoixUseLocalPST.Non;

            try
            {
               
                // Recherche du fichier de statut des PST
                string pathFichierFlgPST = getFullPathFlgPST(btDemande);

                // Si le fichier statut des PST, Récupération des éléments disponibles et indisponible  
                //if (File.Exists(pathFichierFlgPST))
                if (IsExistFlgPSTinLocal(btDemande))
                {
                    // Recupération des statut du ruban
                    readFlagPST(btDemande, lstFlagPST);

                    // Recherche les données dans le Flag PST
                    FlgPSTClass resBT = lstFlagPST.FirstOrDefault(c => c.MailBoxName.ToLower() == btDemande.ToLower() && c.Annee == annee && c.Archive == false);
                    FlgPSTClass resBTArchive = lstFlagPST.FirstOrDefault(c => c.MailBoxName.ToLower() == btDemande.ToLower() && c.Annee == annee && c.Archive == true);
                    // Les 2 info pour la BT principal et Archive sont dispos
                    if (resBT != null && resBTArchive!= null)
                    {
                        
                        // Si BT = Mount et BTArchive = Indisponible

                        // Test si le PST existe Je l'ajout a OUTLOOK
                        if (resBT.Statut==Statut_PST.Mount && resBTArchive.Statut==Statut_PST.Indisponible && File.Exists(getFullPathPST(resBT.MailBoxName, resBT.Annee, resBT.Archive)))
                        {
                            choixUtilisateur = ChoixUseLocalPST.Annule;
                            DialogResult choix;
                            if (InstallForced)
                                choix = DialogResult.Yes;
                            else
                            {
                                if (!Tools.isConnected())
                                {
                                    // Proposer de monter le PST
                                    choix = msgChoixUsePSTOffLine();
                                }
                                else
                                {
                                    // Proposer de monter le PST
                                    choix = msgChoixUsePST();
                                }
                            }

                            if (choix == DialogResult.Yes || choix == DialogResult.OK)
                            {
                                choixUtilisateur = ChoixUseLocalPST.Oui;

                                // Test la Limite OUTLOOK a 15 pst
                                if (Const.lstPST_InSession.Count <= (Const.limitNbPST-1))
                                {
                                    resMountBT = MountPST(resBT.MailBoxName, resBT.Annee, resBT.Archive, resBT.Statut);
                                    if (resMountBT)
                                    {
                                        statut_Ruban_BT = Statut_ruban.OK;
                                        commentaire_BT = MessageClass.Commentaite_Installe;
                                        statut_Ruban_BT_Archive = Statut_ruban.Indisponible;
                                        commentaire_BT_Archive = resBTArchive.Commentaire;
                                    }
                                    else
                                    {
                                        statut_Ruban_BT = Statut_ruban.HS;
                                        commentaire_BT = MessageClass.Commentaite_PST_HS;
                                        choixUtilisateur = ChoixUseLocalPST.Annule;
                                    }
                                // Mise a jours du ruban BT
                                createOrUpdateFlagRuban(Const.lstFlagRuban, resBT.MailBoxName, resBT.Annee, resBT.Archive, statut_Ruban_BT, commentaire_BT);
                                
                                    // Mise a jours du ruban BT Archive
                                createOrUpdateFlagRuban(Const.lstFlagRuban, resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, statut_Ruban_BT_Archive, commentaire_BT_Archive);

                                //YB 2029
                                if (statut_Ruban_BT == Statut_ruban.OK)
                                    MessageBox.Show(string.Format(MessageClass.message_Traitement_Termine, Const.param.anneeDemande), string.Format("{0} - {1}", GetAppVersion(),"Information :"),MessageBoxButtons.OK,MessageBoxIcon.Information);
                                }
                                else
                                {
                                // Si la limite est atteinte du nombre de PST
                                    System.Windows.Forms.MessageBox.Show(string.Format(MessageClass.LimiteNombrePst,Const.limitNbPST), "Alert - Limite OUTLOOK", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                                    statut_Ruban_BT = Statut_ruban.OK;
                                    commentaire_BT = MessageClass.Commentaire_PST_LimiteAtteinte;
                                    statut_Ruban_BT_Archive = Statut_ruban.Indisponible;
                                    commentaire_BT_Archive = resBTArchive.Commentaire;

                                    // Mise a jours du ruban BT
                                    createOrUpdateFlagRuban(Const.lstFlagRuban, resBT.MailBoxName, resBT.Annee, resBT.Archive, statut_Ruban_BT, commentaire_BT);

                                    // Mise a jours du ruban BT Archive
                                    createOrUpdateFlagRuban(Const.lstFlagRuban, resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, statut_Ruban_BT_Archive, commentaire_BT_Archive);
                                }

                            }
                            if (choix == DialogResult.No)
                                choixUtilisateur = ChoixUseLocalPST.Non;                           
                        }

                        // Si BT = Indisponible et BTArchive = Mount
                        if (resBT.Statut == Statut_PST.Indisponible && resBTArchive.Statut == Statut_PST.Mount && File.Exists(getFullPathPST(resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive)))
                        {
                            choixUtilisateur = ChoixUseLocalPST.Annule;
                            DialogResult choix;
                            if (InstallForced)
                                choix = DialogResult.Yes;
                            else
                            {
                                if (!Tools.isConnected())
                                {
                                    // Proposer de monter le PST
                                    choix = msgChoixUsePSTOffLine();
                                }
                                else
                                {
                                    // Proposer de monter le PST
                                    choix = msgChoixUsePST();
                                }
                            }
                            if (choix == DialogResult.Yes || choix == DialogResult.OK)
                            {
                                choixUtilisateur = ChoixUseLocalPST.Oui;
                                // Test la Limite OUTLOOK a 15 pst
                                if (Const.lstPST_InSession.Count <= (Const.limitNbPST-1))
                                {

                                    resMountBT_Archive = MountPST(resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, resBTArchive.Statut);
                                    if (resMountBT_Archive)
                                    {
                                        statut_Ruban_BT = Statut_ruban.Indisponible;
                                        statut_Ruban_BT_Archive = Statut_ruban.OK;
                                        commentaire_BT = resBTArchive.Commentaire;
                                        commentaire_BT_Archive = MessageClass.Commentaite_Installe;
                                    }
                                    else
                                    {
                                        statut_Ruban_BT_Archive = Statut_ruban.HS;
                                        commentaire_BT_Archive = MessageClass.Commentaite_PST_HS;
                                        choixUtilisateur = ChoixUseLocalPST.Annule;
                                    }

                                // Mise a jours du ruban BT
                                createOrUpdateFlagRuban(Const.lstFlagRuban, resBT.MailBoxName, resBT.Annee, resBT.Archive, statut_Ruban_BT, commentaire_BT);
                                // Mise a jours du ruban BT Archive
                                createOrUpdateFlagRuban(Const.lstFlagRuban, resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, statut_Ruban_BT_Archive, commentaire_BT_Archive);

                                //YB 2025
                                if (statut_Ruban_BT_Archive == Statut_ruban.OK)
                                    MessageBox.Show(string.Format(MessageClass.message_Traitement_Termine, Const.param.anneeDemande), string.Format("{0} - {1}", GetAppVersion(),"Information :"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    // Si la limite est atteinte du nombre de PST
                                    System.Windows.Forms.MessageBox.Show(string.Format(MessageClass.LimiteNombrePst, Const.limitNbPST), "Alert - Limite OUTLOOK", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                                    statut_Ruban_BT = Statut_ruban.Indisponible;
                                    commentaire_BT = resBTArchive.Commentaire;
                                    statut_Ruban_BT_Archive = Statut_ruban.OK;
                                    commentaire_BT_Archive = MessageClass.Commentaire_PST_LimiteAtteinte;

                                    // Mise a jours du ruban BT
                                    createOrUpdateFlagRuban(Const.lstFlagRuban, resBT.MailBoxName, resBT.Annee, resBT.Archive, statut_Ruban_BT, commentaire_BT);
                                    // Mise a jours du ruban BT Archive
                                    createOrUpdateFlagRuban(Const.lstFlagRuban, resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, statut_Ruban_BT_Archive, commentaire_BT_Archive);
                                }
                            }
                            if (choix == DialogResult.No)
                                choixUtilisateur = ChoixUseLocalPST.Non;
                        }



                        // Si BT = Mount et BTArchive = Mount
                        if (resBT.Statut == Statut_PST.Mount && resBTArchive.Statut == Statut_PST.Mount && File.Exists(getFullPathPST(resBT.MailBoxName, resBT.Annee, resBT.Archive)) && File.Exists(getFullPathPST(resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive)))
                        {
                            choixUtilisateur = ChoixUseLocalPST.Annule;
                            DialogResult choix;
                            if (InstallForced)
                                choix = DialogResult.Yes;
                            else
                                if (!Tools.isConnected())
                                    {
                                        // Proposer de monter le PST
                                        choix = msgChoixUsePSTOffLine();
                                    }
                                    else
                                    {
                                        // Proposer de monter le PST
                                        choix = msgChoixUsePST();
                                    }
                            if (choix == DialogResult.Yes || choix==DialogResult.OK)
                            {
                                choixUtilisateur = ChoixUseLocalPST.Oui;
                                // Test la Limite OUTLOOK a 15 pst
                                if (Const.lstPST_InSession.Count <= (Const.limitNbPST-2))
                                {
                                    
                                    resMountBT = MountPST(resBT.MailBoxName, resBT.Annee, resBT.Archive, resBT.Statut);
                                    if (resMountBT)
                                    {
                                        statut_Ruban_BT = Statut_ruban.OK;
                                        commentaire_BT = MessageClass.Commentaite_Installe;
                                    }
                                    else
                                    {
                                        statut_Ruban_BT = Statut_ruban.HS;
                                        commentaire_BT = MessageClass.Commentaite_PST_HS;
                                        choixUtilisateur=ChoixUseLocalPST.Annule;
                                    }
                               
                                    // Mise a jours du ruban BT
                                    createOrUpdateFlagRuban(Const.lstFlagRuban, resBT.MailBoxName, resBT.Annee, resBT.Archive, statut_Ruban_BT, commentaire_BT);

                                
                                    resMountBT_Archive = MountPST(resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, resBTArchive.Statut);
                                    if (resMountBT_Archive)
                                    {
                                        statut_Ruban_BT_Archive = Statut_ruban.OK;
                                        commentaire_BT_Archive = MessageClass.Commentaite_Installe;
                                    }
                                    else
                                    {
                                        statut_Ruban_BT_Archive = Statut_ruban.HS;
                                        commentaire_BT_Archive = MessageClass.Commentaite_PST_HS;
                                        choixUtilisateur = ChoixUseLocalPST.Annule;
                                    }

                                // Mise a jours du ruban BT Archive
                                createOrUpdateFlagRuban(Const.lstFlagRuban, resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, statut_Ruban_BT_Archive, commentaire_BT_Archive);

                                if (statut_Ruban_BT == Statut_ruban.OK && statut_Ruban_BT_Archive == Statut_ruban.OK)
                                    MessageBox.Show(string.Format(MessageClass.message_Traitement_Termine, Const.param.anneeDemande), string.Format("{0} - {1}", GetAppVersion(),"Information :"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                else
                                {
                                    // Si la limite est atteinte du nombre de PST
                                    System.Windows.Forms.MessageBox.Show(string.Format(MessageClass.LimiteNombrePst, Const.limitNbPST), "Alert - Limite OUTLOOK", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                                    statut_Ruban_BT = Statut_ruban.OK;
                                    commentaire_BT = MessageClass.Commentaire_PST_LimiteAtteinte;
                                    statut_Ruban_BT_Archive = Statut_ruban.OK;
                                    commentaire_BT_Archive = MessageClass.Commentaire_PST_LimiteAtteinte;
                                    choixUtilisateur = ChoixUseLocalPST.Annule;

                                    // Mise a jours du ruban BT Archive
                                    createOrUpdateFlagRuban(Const.lstFlagRuban, resBTArchive.MailBoxName, resBTArchive.Annee, resBTArchive.Archive, statut_Ruban_BT_Archive, commentaire_BT_Archive);
                                }
                            }
                            if (choix == DialogResult.No)
                                choixUtilisateur = ChoixUseLocalPST.Non;
                        }
                    }
                    else
                    {
                        // Si y a pas de flag PST Incomplet 
                        // Rien Faire


                    }
                }
                else
                {
                    // Fichier pas de Flag PST - ABSENT
                    // Test si les PST existe et création du Flag PST
                    if (File.Exists(getFullPathPST(btDemande, annee, false)) && (File.Exists(getFullPathPST(btDemande, annee, true))))
                    {
                        choixUtilisateur = ChoixUseLocalPST.Annule;
                        DialogResult choix;
                        if (InstallForced)
                            choix = DialogResult.Yes;
                        else
                        {
                            if (!Tools.isConnected())
                                choix = msgChoixUsePSTOffLine();
                            else
                                choix = msgChoixUsePST();
                        }

                            // Proposer de monter le PST Local
                        if (choix == DialogResult.Yes || choix == DialogResult.OK)
                        {
                            choixUtilisateur = ChoixUseLocalPST.Oui;
                            // Test la Limite OUTLOOK a 15 pst
                            if (Const.lstPST_InSession.Count < (Const.limitNbPST-2))
                            {
                                // Mise en place du Flag PST
                                resMountBT = MountPST(btDemande, annee, false, Statut_PST.Mount);
                                if (resMountBT)
                                {
                                    statut_Ruban_BT_PST = Statut_PST.Mount;
                                    commentaire_BT_PST = MessageClass.Commentaite_Installe;
                                }
                                else
                                {
                                    // Ne pas changer le statut pour pouvoir réessayer
                                }

                                createOrUpdateFlagPST(lstFlagPST, btDemande, annee, false, statut_Ruban_BT_PST, commentaire_BT_PST);


                                // Mise en place du Flag PST ARCHIVE
                                resMountBT_Archive = MountPST(btDemande, annee, true, Statut_PST.Mount);
                                if (resMountBT_Archive)
                                {
                                    statut_Ruban_BT_Archive_PST = Statut_PST.Mount;
                                    commentaire_BT_Archive_PST = MessageClass.Commentaite_Installe;
                                }
                                else
                                {
                                    // Ne pas changer le statut pour pouvoir réessayer
                                }


                                createOrUpdateFlagPST(lstFlagPST, btDemande, annee, true, statut_Ruban_BT_Archive_PST, commentaire_BT_Archive_PST);

                                // Mise en place du Flag Ruban
                                if (resMountBT)
                                {
                                    statut_Ruban_BT = Statut_ruban.OK;
                                    commentaire_BT = MessageClass.Commentaite_Installe;
                                }
                                else
                                {
                                    statut_Ruban_BT = Statut_ruban.OK;
                                    commentaire_BT = MessageClass.Commentaite_PST_HS;
                                    choixUtilisateur = ChoixUseLocalPST.Annule;
                                }

                                if (resMountBT_Archive)
                                {
                                    statut_Ruban_BT_Archive = Statut_ruban.OK;
                                    commentaire_BT_Archive = MessageClass.Commentaite_Installe;
                                }
                                else
                                {
                                    statut_Ruban_BT_Archive = Statut_ruban.HS;
                                    commentaire_BT_Archive = MessageClass.Commentaite_PST_HS;
                                    choixUtilisateur = ChoixUseLocalPST.Annule;
                                }

                            // Mise a jours du ruban BT
                            createOrUpdateFlagRuban(Const.lstFlagRuban, btDemande, annee, false, statut_Ruban_BT, commentaire_BT);
                            // Mise a jours du ruban BT Archive
                            createOrUpdateFlagRuban(Const.lstFlagRuban, btDemande, annee, true, statut_Ruban_BT_Archive, commentaire_BT_Archive);

                            if (statut_Ruban_BT == Statut_ruban.OK && statut_Ruban_BT_Archive == Statut_ruban.OK)
                                MessageBox.Show(string.Format(MessageClass.message_Traitement_Termine, Const.param.anneeDemande), string.Format("{0} - {1}", GetAppVersion(),"Information :"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                // Si la limite est atteinte du nombre de PST
                                System.Windows.Forms.MessageBox.Show(string.Format(MessageClass.LimiteNombrePst, Const.limitNbPST), "Alert - Limite OUTLOOK", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                                statut_Ruban_BT = Statut_ruban.OK;
                                commentaire_BT = MessageClass.Commentaire_PST_LimiteAtteinte;
                                statut_Ruban_BT_Archive = Statut_ruban.OK;
                                commentaire_BT_Archive = MessageClass.Commentaire_PST_LimiteAtteinte;
                                choixUtilisateur = ChoixUseLocalPST.Annule;

                                // Mise a jours du ruban BT
                                createOrUpdateFlagRuban(Const.lstFlagRuban, btDemande, annee, false, statut_Ruban_BT, commentaire_BT);
                                // Mise a jours du ruban BT Archive
                                createOrUpdateFlagRuban(Const.lstFlagRuban, btDemande, annee, true, statut_Ruban_BT_Archive, commentaire_BT_Archive);
                            }


                        }
                        if (choix == DialogResult.No)
                            choixUtilisateur = ChoixUseLocalPST.Non;

                    }
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors de la création ou modification du flag ruban pour la boite {0} dans le profil {1}. \n", btDemande, Const.path_Local_Flg_For_Session), ex, true, MethodBase.GetCurrentMethod().Name);
                return ChoixUseLocalPST.Annule;
            }
            return choixUtilisateur;
        }


        /// <summary>
        /// Actualisation du ruban
        /// </summary>
        /// <param name="ribbon"></param>
        /// <param name="btDemande"></param>
        /// <param name="lstFlagDeStatutRuban"></param>
        public static void refreshRuban(RibbonMesArchivesExplorer ribbon, string btDemande, List<FlgRubanClass> lstFlagDeStatutRuban)
        {
            try
            {
                FlgRubanClass infoRuban;
                FlgRubanClass btPrincipal;
                FlgRubanClass btArchive;
                bool btPrincipalInSession=false;
                bool btArchiveInSession=false;

                // Maj du control afin de savoir si la personne est connected
                //Const.isConnnectMode = Tools.ModeConnected();

                // Si Nous somme connecté actualisation de la liste des PST
                if (Tools.isConnected())
                    Traitement.PSTInSession();

                // Affiche : Nom de la boite au lettre séléctionné ou Mes archives mails, si l'élément est un PST local
                ribbon.grRestauration_av.Label = string.IsNullOrEmpty(btDemande) ? "Mes archives mails.": btDemande;

                // Complete le label du ruban
                if (!Tools.isConnected())
                    ribbon.grRestauration_av.Label += " (hors connexion).";

                // Si btDemande =="" desactiver les boutons
                if (string.IsNullOrEmpty(btDemande))
                {
                    // Desactivation de tous les boutons du ruban
                    visuelAllButton(ribbon,false);
                    return;
                }

                // Recupération pour chaque année des informations dans le flag _statut
                bool res=readFlagRuban(btDemande, lstFlagDeStatutRuban);
                if (res)
                {
                    //for (int annee = Const.anneeDebut; annee<2025; annee++)
                    for (int annee = 2015; annee < 2025; annee++)
                    {
                        btPrincipal = lstFlagDeStatutRuban.FirstOrDefault(c => c.MailBoxName.ToLower() == btDemande.ToLower() && c.Annee == annee.ToString() && c.Archive==false);
                        // Si les information n'existe pas dans le flag statut, définir les valeurs par défaut pour la BT Principal
                        if (btPrincipal==null)
                        {
                            btPrincipal=new FlgRubanClass();
                            btPrincipal.MailBoxName= btDemande;
                            btPrincipal.Annee=annee.ToString();
                            btPrincipal.Archive=false;
                            btPrincipal.Commentaire = "";
                            btPrincipal.Statut = Statut_ruban.None;
                        }

                        btArchive = lstFlagDeStatutRuban.FirstOrDefault(c => c.MailBoxName.ToLower() == btDemande.ToLower() && c.Annee == annee.ToString() && c.Archive == true);
                        // Si les information n'existe pas dans le flag statut, définir les valeurs par défaut pour la BT ARCHIVE
                        if (btArchive == null)
                        {
                            btArchive = new FlgRubanClass();
                            btArchive.MailBoxName = btDemande;
                            btArchive.Annee = annee.ToString();
                            btArchive.Archive = false;
                            btArchive.Commentaire = "";
                            btArchive.Statut = Statut_ruban.None;
                        }

                        // CAS : un des éléments est Annule
                        if (btPrincipal.Statut == Statut_ruban.Annule || btArchive.Statut == Statut_ruban.Annule)
                        {
                            // Mettre l'info comme non disponible
                            infoRuban = btPrincipal;
                            if (btArchive.Statut == Statut_ruban.Annule)
                            {
                                // Mettre l'info comme non disponible
                                infoRuban = btArchive;
                            }
                            // Selection du bouton
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);
                            // Next bouton
                            continue;
                        }

                        //****************************************
                        // CAS : Les éléments sont OK
                        //****************************************
                        if (btPrincipal.Statut == Statut_ruban.OK && btArchive.Statut == Statut_ruban.OK)
                        {
                            // Personnalisation du bouton sur le statut de la boite principal
                            infoRuban = btPrincipal;
                            infoRuban.Statut = btPrincipal.Statut;
                            infoRuban.Commentaire = btPrincipal.Commentaire+"\n"; //"";


                            if (!Tools.isConnected())
                            {
                                infoRuban.Commentaire += MessageClass.RechercheIndisponibleEnOffLigne;
                            }
                            else
                            {
                                // Les PST sont t'ils toujours present 
                                btPrincipalInSession = IsMountInSession(btPrincipal.MailBoxName, btPrincipal.Annee, btPrincipal.Archive);
                                btArchiveInSession = IsMountInSession(btArchive.MailBoxName, btArchive.Annee, btArchive.Archive);

                                // Si on est Connecter on recherche si le fichier es dans les PST de la session
                                if (!btPrincipalInSession)
                                {
                                    infoRuban.Statut = Statut_ruban.Partial;
                                    infoRuban.Commentaire += MessageClass.Commentaire_PST_ABS_Session;
                                }
                                if (!btArchiveInSession)
                                {
                                    infoRuban.Statut = Statut_ruban.Partial;
                                    infoRuban.Commentaire += MessageClass.Commentaire_PST_Archive_ABS_Session;
                                }

                                // Mettre le statut en dispo au lieu de partiel lors qu'aucun n'est present
                                if (!btPrincipalInSession && !btArchiveInSession)
                                {
                                    infoRuban.Statut = Statut_ruban.Mount;
                                    if (btPrincipal.Statut == Statut_ruban.Mount)
                                    {
                                        if (!IsExistPSTinLocal(btDemande, annee.ToString(), btPrincipal.Archive))
                                        {
                                            infoRuban.Statut = Statut_ruban.None;
                                            infoRuban.Commentaire += MessageClass.Commentaire_PST_ABS;
                                        }
                                    }
                                    if (btArchive.Statut == Statut_ruban.Mount && infoRuban.Statut != Statut_ruban.None)
                                    {
                                        if (!IsExistPSTinLocal(btDemande, annee.ToString(), btArchive.Archive))
                                        {
                                            infoRuban.Statut = Statut_ruban.None;
                                            infoRuban.Commentaire += MessageClass.Commentaire_PST_Archive_ABS;
                                        }
                                    }
                                }


                            }
                            
                            

                            // Selection du bouton
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);

                            // Next bouton
                            continue;
                        }


                        //****************************************
                        // CAS : HS
                        //****************************************
                        // CAS : Un des élément Principal est HS
                        if (btPrincipal.Statut == Statut_ruban.HS && btArchive.Statut != Statut_ruban.HS)
                        {
                            // Personnalisation du bouton sur le statut de la boite principal
                            infoRuban = btPrincipal;
                            infoRuban.Statut = Statut_ruban.HS; //Statut_ruban.Encours
                                                                // Selection du bouton
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);
                            // Next bouton
                            continue;
                        }

                        // CAS : Un des élément Archive est HS
                        if (btPrincipal.Statut != Statut_ruban.HS && btArchive.Statut == Statut_ruban.HS)
                        {
                            // Personnalisation du bouton sur le statut de la boite Archive
                            infoRuban = btArchive;
                            infoRuban.Statut = Statut_ruban.HS; //Statut_ruban.Encours
                                                                // Selection du bouton
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);
                            // Next bouton
                            continue;
                        }

                        //****************************************
                        // CAS : Disponible
                        //****************************************
                        // CAS : Un des élément Archive est Indisponible -> Prendre statut Principal
                        if (btPrincipal.Statut == Statut_ruban.Mount || btArchive.Statut == Statut_ruban.Mount)
                        {
                            
                            // Personnalisation du bouton sur le statut de la boite principal
                            infoRuban = btPrincipal;
                            infoRuban.Statut = btPrincipal.Statut;
                            // Particularite si Disponible ou OK -> Verification de la presence du Fichier
                            if (btPrincipal.Statut == Statut_ruban.Mount)
                            {
                                if (!IsExistPSTinLocal(btDemande, annee.ToString(), btPrincipal.Archive))
                                {
                                    infoRuban.Statut = Statut_ruban.HS;
                                    infoRuban.Commentaire += MessageClass.Commentaire_PST_ABS; ;
                                }
                            }

                            if(btArchive.Statut == Statut_ruban.Mount && infoRuban.Statut!= Statut_ruban.HS)
                            {
                                if (!IsExistPSTinLocal(btDemande, annee.ToString(), btPrincipal.Archive))
                                {
                                    infoRuban.Statut = Statut_ruban.HS;
                                    infoRuban.Commentaire += MessageClass.Commentaire_PST_Archive_ABS; ;
                                }
                            }

                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);
                            // Next bouton
                            continue;
                        }


                        //****************************************
                        // CAS : Indisponible
                        //****************************************
                        // CAS : Un des élément Principal est Indisponible -> Prendre statut Archive
                        if (btPrincipal.Statut == Statut_ruban.Indisponible || btArchive.Statut != Statut_ruban.Indisponible)
                        {
                            
                            // Personnalisation du bouton sur le statut de la boite principal
                            infoRuban = btArchive;

                            if (!Tools.isConnected())
                            {
                                infoRuban.Commentaire += MessageClass.RechercheIndisponibleEnOffLigne;
                            }
                            else
                            {
                                // Particularite si Disponible ou OK -> Verification de la presence du Fichier
                                if (btArchive.Statut == Statut_ruban.Mount || btArchive.Statut == Statut_ruban.OK)
                                {
                                    infoRuban.Statut = IsExistPSTinLocal(btDemande, annee.ToString(), btArchive.Archive) ? btArchive.Statut : Statut_ruban.None;
                                    
                                    // Les PST sont t'ils toujours present dans la session
                                    btArchiveInSession = IsMountInSession(btArchive.MailBoxName, btArchive.Annee, btArchive.Archive);                                   

                                    if (btArchive.Statut == Statut_ruban.OK && !btArchiveInSession)
                                    {
                                        infoRuban.Statut = Statut_ruban.Mount;
                                        infoRuban.Commentaire = MessageClass.Commentaire_PST_Archive_ABS_Session;
                                    }
                                }
                                else
                                    infoRuban.Statut = btArchive.Statut; //Statut_ruban.Encours
                                                                         // Selection du bouton
                            }
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);
                            // Next bouton
                            continue;
                        }

                        // CAS : Un des élément Archive est Indisponible -> Prendre statut Principal
                        if (btPrincipal.Statut != Statut_ruban.Indisponible || btArchive.Statut == Statut_ruban.Indisponible)
                        {
                            // Personnalisation du bouton sur le statut de la boite principal
                            infoRuban = btPrincipal;
                            if (!Tools.isConnected())
                            {
                                infoRuban.Commentaire += MessageClass.RechercheIndisponibleEnOffLigne;
                            }
                            else
                            { 
                                // Particularite si Disponible ou OK -> Verification de la presence du Fichier
                                if (btPrincipal.Statut == Statut_ruban.Mount || btPrincipal.Statut == Statut_ruban.OK)
                                {
                                    // Recherche si le PST est toujours present en Local
                                    infoRuban.Statut = IsExistPSTinLocal(btDemande, annee.ToString(), btPrincipal.Archive) ? btPrincipal.Statut : Statut_ruban.None;
                                    
                                    // Les PST sont t'ils toujours present 
                                    btPrincipalInSession = IsMountInSession(btPrincipal.MailBoxName, btPrincipal.Annee, btPrincipal.Archive);
                                    if (btPrincipal.Statut == Statut_ruban.OK && !btPrincipalInSession)
                                    {
                                        infoRuban.Statut = Statut_ruban.Mount;
                                        infoRuban.Commentaire = MessageClass.Commentaire_PST_ABS_Session;
                                    }
                               }
                                else
                                    infoRuban.Statut = btPrincipal.Statut; //Statut_ruban.Encours
                                                                           // Selection du bouton
                            }
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);
                            // Next bouton
                            continue;
                        }

                        // CAS : Statut identique -> Mettre le statut de Principal
                        if (btPrincipal.Statut == btArchive.Statut)
                        {
                            // Personnalisation du bouton sur le statut de la boite principal
                            infoRuban = btPrincipal;
                            infoRuban.Statut = btPrincipal.Statut;
                                
                            // Selection du bouton
                            RibbonButton bt1 = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                            // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                            visuelInButton(bt1, infoRuban);

                            // Next bouton
                            continue;
                        }

                        // ELSE CAS non traité : -> Mettre le statut de Principal
                        // Personnalisation du bouton sur le statut de la boite principal
                        infoRuban = btPrincipal;
                        infoRuban.Statut = btPrincipal.Statut;

                        // Selection du bouton
                        RibbonButton bt = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                        // Appliquer le visuel au bouton suivant les info des flags pour le ruban
                        visuelInButton(bt, infoRuban);

                        // Next bouton
                        continue;

                    }

                }
                else
                {
                    // Puis Actualisation Visuel du ruban avec les valeurs par Defaut
                    visuelAllButton(ribbon);
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de la Actualisation du ruban en fonction des informations du fichier .flg.", ex, true, MethodBase.GetCurrentMethod().Name);
            }
        }


        public static bool PSTInSession()
        {
            bool res = false;
            try
            {
                // Si je suis connecte
                if (Tools.isConnected())
                {
                    Const.lstPST_InSession.Clear();
                    foreach (Microsoft.Office.Interop.Outlook.Store itemStore in Globals.ThisAddIn.Application.Session.Stores)
                    {
                        // PST ou OST
                        if (itemStore.IsDataFileStore)
                        {
                            if (itemStore.FilePath.Trim().ToLower().EndsWith(".pst"))
                            {
                                Const.lstPST_InSession.Add(itemStore.FilePath.Trim().ToLower());                                
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log inutile de tracer
                if (ex.HResult== -2009857775)
                {
                    // PST supprimer du poste => Laisser que le message d'erreur OUTLOOK
                }
                else
                {
                    //Tools.LogMessage("Erreur lors de la création de la liste memoire des PST {0} de votre session OUTLOOK.", ex, false, MethodBase.GetCurrentMethod().Name);
                }
                
                res = false;
            }
            return res;
        }


        /// <summary>
        /// Contrôle la présence du PST dans la Session OUTLOOK
        /// </summary>
        /// <param name="mailBoxName"></param>
        /// <param name="annee"></param>
        /// <param name="archive"></param>
        /// <returns></returns>
        public static bool IsMountInSession(string mailBoxName, string annee, bool archive)
        {
            bool res = false;
            //string store;
            string cheminPST = "";
            try
            {


                // Chemin du PST a rechercher 
                if (!string.IsNullOrEmpty(mailBoxName) && !string.IsNullOrEmpty(annee))
                {
                    cheminPST = getFullPathPST(mailBoxName, annee, archive);
                    if (!string.IsNullOrEmpty(cheminPST))
                    {
                        cheminPST = cheminPST.Trim().ToLower();
                    }
                    else
                    {
                        return false;
                    }

                }

                return Const.lstPST_InSession.Contains(cheminPST);

            }
            catch (Exception ex)
            {
                Tools.LogMessage(string.Format("Erreur lors de la recherche du PST {0} dans votre session OUTLOOK.", cheminPST), ex, true, MethodBase.GetCurrentMethod().Name);
                res = false;
            }
            return res;
        }

        /// <summary>
        /// Fonction de Montage du PST dans la session OUTLOOK de l'utilisateur
        /// </summary>
        /// <param name="mailBoxName"></param>
        /// <param name="annee"></param>
        /// <param name="archive"></param>
        /// <param name="statut"></param>
        /// <returns></returns>
        private static bool MountPST(string mailBoxName, string annee, bool archive, Statut_PST statut )
        {
            bool res = false;
            string cheminFichierPST = "";
            try
            {

                // Chemin du PST
                cheminFichierPST = getFullPathPST(mailBoxName, annee, archive);
                // Test si le PST existe Je l'ajout a OUTLOOK
                if (File.Exists(cheminFichierPST) && statut == Statut_PST.Mount)
                {
                    // Ajout dans du PST
                    Globals.ThisAddIn.Application.Session.AddStore(cheminFichierPST);
                    //Globals.ThisAddIn.Application.GetNamespace("MAPI").AddStore(cheminFichierPST);
                    if (!Const.lstPST_InSession.Contains(cheminFichierPST))
                    {
                        Const.lstPST_InSession.Add(cheminFichierPST);
                    }
                    res = true;
                }
            }
            catch (Exception ex)
            {
                //// Si cela n'a pas fonctionné, je passe le flag de Session (profil) a a HS
                Tools.LogMessage(string.Format("Erreur lors de l'ajout du PST dans votre session OUTLOOK. {0}", cheminFichierPST), ex, true, MethodBase.GetCurrentMethod().Name);
                res = false;
            }
            return res;
        }


        /// <summary>
        /// Fixe le Visuel par defaut pour tous les boutons
        /// </summary>
        /// <param name="ribbon"></param>
        /// <param name="enanbleBt"></param>
        private static void visuelAllButton(RibbonMesArchivesExplorer ribbon, bool enanbleBt = true)
        {
            try
            {
                RibbonButton r;
                //for (int annee = Const.anneeDebut; annee < 2025; annee++)
                for (int annee = 2015; annee < 2025; annee++)
                {
                    r = Globals.Ribbons.RibbonDownloadExplorer.GetBtn(annee.ToString());
                    r.Visible = annee >= Const.anneeDebut ? true : false;
                    r.Enabled = enanbleBt;
                    r.Label = annee.ToString();
                    r.Image = OutlookPST.Properties.Resources.ico_Blanc;
                    r.SuperTip = "";
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de l'init des boutons du ruban.", ex, true, MethodBase.GetCurrentMethod().Name);
            }
        }


        /// <summary>
        /// Visuel Applique au bouton (bt) suivant les informations de restoration (infoflg)
        /// </summary>
        /// <param name="bt"></param>
        /// <param name="infoflg"></param>
        private static void visuelInButton(RibbonButton bt, FlgRubanClass infoflg)
        {
            try
            {
                string texthorsConnexion = "";
                if (!Tools.isConnected())
                    texthorsConnexion = " *";

                bt.Visible = int.Parse(infoflg.Annee) >= Const.anneeDebut ? true:false ;
                if (infoflg.Statut == Statut_ruban.None)
                {
                    //bt.Visible = true;
                    bt.Label = infoflg.Annee;
                    bt.Image = OutlookPST.Properties.Resources.ico_Blanc;
                    bt.Enabled = true;
                    bt.SuperTip = infoflg.Commentaire;
                }

                if (infoflg.Statut == Statut_ruban.OK)
                {
                    //bt.Visible = true;
                    bt.Label = infoflg.Annee + " Installé"+ texthorsConnexion;
                    bt.Image = OutlookPST.Properties.Resources.ico_Vert;
                    bt.Enabled = true;
                    bt.SuperTip = infoflg.Commentaire;
                }

                if (infoflg.Statut == Statut_ruban.HS)
                {
                    //bt.Visible = true;
                    bt.Label = infoflg.Annee + " - Erreur";
                    bt.Enabled = true;
                    bt.Image = OutlookPST.Properties.Resources.ico_Rouge;
                    bt.SuperTip = infoflg.Commentaire;
                }
                if (infoflg.Statut == Statut_ruban.Mount)
                {
                    //bt.Visible = true;
                    bt.Image = OutlookPST.Properties.Resources.ico_Jaune;
                    bt.Enabled = true;
                    bt.Label = infoflg.Annee + " - Disponible"+texthorsConnexion;
                    bt.SuperTip = infoflg.Commentaire;
                }
                if (infoflg.Statut == Statut_ruban.Encours)
                {
                    //bt.Visible = true;
                    bt.Image = OutlookPST.Properties.Resources.ico_Bleu;
                    bt.Enabled = true;
                    bt.Label = infoflg.Annee + " - " + Statut_ruban.Encours.ToString();
                    bt.SuperTip = infoflg.Commentaire;
                }
                if (infoflg.Statut == Statut_ruban.Indisponible)
                {
                    //bt.Visible = true;
                    bt.Image = OutlookPST.Properties.Resources.ico_Blanc;
                    bt.Enabled = true;
                    bt.Label = infoflg.Annee + " - " + Statut_ruban.Indisponible.ToString();
                    bt.SuperTip = infoflg.Commentaire;
                }
                if (infoflg.Statut == Statut_ruban.Annule)
                {
                    //bt.Visible = true;
                    bt.Image = OutlookPST.Properties.Resources.ico_Rouge;
                    bt.Enabled = true;
                    bt.Label = infoflg.Annee + " - Annulé par l'utilisateur";
                    bt.SuperTip = infoflg.Commentaire;
                }

                if (infoflg.Statut == Statut_ruban.Partial)
                {
                    //bt.Visible = true;
                    bt.Image = OutlookPST.Properties.Resources.ico_Jaune;
                    bt.Enabled = true;
                    bt.Label = infoflg.Annee + " - Partiel";
                    bt.SuperTip = infoflg.Commentaire;
                }

            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de la Actualisation du ruban en fonction de information du fichier .flg pour le ruban.", ex, true, MethodBase.GetCurrentMethod().Name);
            }
        }

        /// <summary>
        /// Renvoi la boite currente séléctionnée
        /// </summary>
        /// <returns></returns>
        public static string getCurrentMailBox()
        {
            string currentBoite = "";
            try
            {
                if (Globals.ThisAddIn.Application.ActiveExplorer() is null)
                {
                    return "";
                }
                if (Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder is null)
                {
                    return "";
                }
                // Les étapes suivantes permettent de récupérer la sélection actuelle de mail(s)
                currentBoite = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.FolderPath;
                //Tools.LogMessage("Folder selectionné "+ currentBoite, null, false, MethodBase.GetCurrentMethod().Name);
                if (!string.IsNullOrEmpty(currentBoite))
                {
                    //Tools.LogMessage("Boîte : "+ currentBoite, null, false, MethodBase.GetCurrentMethod().Name);
                    if (Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Store is null)
                    {
                        //Tools.LogMessage("CurrentFolder.Store is null", null, false, MethodBase.GetCurrentMethod().Name);
                        return "";    
                    }
                    else
                    {
                        // Si c'est une archive PST -> currentBoite = ""
                        if ((!string.IsNullOrEmpty(Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Store.FilePath)) && (Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Store.FilePath.LastIndexOf(".pst") > 0))
                        {
                            return "";
                        }
                        currentBoite = currentBoite.Replace("\\\\", "");
                        if (currentBoite.Contains(@"\"))
                        {
                            currentBoite = currentBoite.Substring(0, currentBoite.IndexOf(@"\"));
                        }
                    }
                }
                else
                {
                    // Si  Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.FolderPath ="" 
                    currentBoite = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.Name;
                    currentBoite = currentBoite.Replace("\\\\", "");
                    if (currentBoite.Contains(@"\"))
                    {
                        currentBoite = currentBoite.Substring(0, currentBoite.IndexOf(@"\"));
                    }
                }
                //System.Windows.Forms.MessageBox.Show(currentBoite);
                //isTenYear(currentBoite);
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de la récupération Boîte aux lettres séléctionnée.", ex, false, MethodBase.GetCurrentMethod().Name);
                currentBoite = "";
            }

            return currentBoite;
        }

        public static void isTenYear(string mail, bool writeifnecessarie=false)
        {
            try
            {
                // Recherche dans la liste
                List<string> lTenDay=new List<string>();
                lTenDay.Add("@acpr.banque-france.fr");
                // Pour les livraisons
                //lTenDay.Add("yannick.blangis@test-banque-france.fr");

                foreach (string itemMail in lTenDay)
                {
                    if (mail.ToLower().Contains(itemMail))
                    {
                        // Fixe l'année les infos pour le ruban 10 ans 
                        majAnneeDebutRuban(true);
                        return;
                    }
                }

                // puis sinon , On recherche dans le Flag Ruban 
                if (initTenYearByFlagRuban(mail))
                {
                    return;
                }

                // Sinon Recherche AD de la rétention dans l'AD si le poste est connecté
                if (Tools.isConnected())
                {
                    // Recupération dans l'AD de la rétention pour la Bal
                    if (AD.getInfoRetention(mail.ToLower()))
                    {
                        return;    
                    }
                }


                // Sinon Mise en place du ruban par defaut 5 ans
                majAnneeDebutRuban(false);
            }
            catch 
            {
                // Si erreur
                // Sinon Mise en place du ruban par defaut 5 ans
                majAnneeDebutRuban(false);
            }
    
     }


        /// <summary>
        /// Caclul la date de debut pour les boutons
        /// </summary>
        /// <returns></returns>
        public static void majAnneeDebutRuban(bool isBalTenYear)
        {
            Const.isTenYear = isBalTenYear;
            Const.anneeDebut = isBalTenYear ? DateTime.Now.AddYears(-10).Year : DateTime.Now.AddYears(-5).Year;
        }

        /// <summary>
        /// Renvoi le chemin complet du fichier flg PST
        /// </summary>
        /// <param name="mailbox"></param>
        /// <param name="annee"></param>
        /// <param name="archive"></param>
        /// <returns></returns>
        public static string getFullPathFlgPST(string mailbox)
        {
            try
            {
                string cheminFichierPST = Path.Combine(Const.path_Local_Flg_PST, Traitement.getNameFlgPST(mailbox));
                return cheminFichierPST;
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Renvoi si le fichier flag PST est présent localement
        /// </summary>
        /// <param name="mailbox"></param>
        /// <param name="annee"></param>
        /// <param name="archive"></param>
        /// <returns></returns>
        public static bool IsExistFlgPSTinLocal(string mailbox)
        {
            try
            {
                string cheminFichierPST = getFullPathFlgPST(mailbox);
                return File.Exists(cheminFichierPST);
            }
            catch 
            {
                return false;
            }
        }

        /// <summary>
        /// Norme de nommage du pst
        /// </summary>
        /// <param name="MailBoxName"></param>
        /// <param name="Annee"></param>
        /// <param name="Archive"></param>
        /// <returns></returns>
        public static string getNameFlgPST(string mailbox)
        {
            return string.Format("{0}_statut.flg", mailbox);
        }


        /// <summary>
        /// Renvoi le chemin complet du fichier
        /// </summary>
        /// <param name="mailbox"></param>
        /// <param name="annee"></param>
        /// <param name="archive"></param>
        /// <returns></returns>
        public static string getFullPathPST(string mailbox, string annee, bool archive)
        {
            try
            {
                string cheminFichierPST = Path.Combine(Const.path_Local_PST, Traitement.getNamePST(mailbox, annee, archive));
                return cheminFichierPST;
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Renvoi si le fichier PST est présent localement
        /// </summary>
        /// <param name="mailbox"></param>
        /// <param name="annee"></param>
        /// <param name="archive"></param>
        /// <returns></returns>
        public static bool IsExistPSTinLocal(string mailbox, string annee, bool archive)
        {
            try
            {
                string cheminFichierPST = getFullPathPST(mailbox, annee, archive);
                return File.Exists(cheminFichierPST);
            }
            catch (Exception)
            {
                return false;
            }
        }
        /// <summary>
        /// Norme de nommage du pst
        /// </summary>
        /// <param name="MailBoxName"></param>
        /// <param name="Annee"></param>
        /// <param name="Archive"></param>
        /// <returns></returns>
        public static string getNamePST(string mailbox, string annee, bool archive)
        {
            string nomFile = "";
            try
            {
                if (archive)
                {
                    nomFile = mailbox + "_" + annee +"_archive.pst";
                }
                else
                {
                    nomFile = mailbox + "_" + annee + ".pst";      
                }
            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur de récupération du nom du fichier .flg pour le PST. ", ex, false, MethodBase.GetCurrentMethod().Name);
                nomFile = "";
            }
            return nomFile;
        }

        /// <summary>
        /// Renvoi le chemin complet du fichier
        /// </summary>
        /// <param name="mailbox"></param>
        /// <returns></returns>
        public static string getFullPathFlg(string mailbox)
        {
            try
            {
                string cheminFichierFlg = Path.Combine(Const.path_Local_Flg_For_Session, getNameFlg(mailbox));
                return cheminFichierFlg;
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Renvoi si le fichier Flg est présent localement
        /// </summary>
        /// <param name="mailbox"></param>
        /// <returns></returns>
        public static bool IsExistFlginLocal(string mailbox)
        {
            try
            {
                string cheminFichierFlg = getFullPathFlg(mailbox);
                return File.Exists(cheminFichierFlg);
            }
            catch 
            {
                return false;
            }
        }

        /// <summary>
        /// Norme de nommage du flg ruban
        /// </summary>
        /// <param name="btDemande"></param>
        public static string getNameFlg(string btDemande)
        {
            return string.Format("{0}_statut.flg", btDemande);
        }

    }
}
