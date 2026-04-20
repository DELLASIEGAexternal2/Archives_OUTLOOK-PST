using Amazon.S3.IO;
using Amazon.S3.Model.Internal.MarshallTransformations;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Amazon.Util.Internal;
using System.IO.Compression;
using F = System.Windows.Forms;
using SevenZipExtractor;
using System.Windows.Forms;
using System.Configuration;
using System.Threading;
using OutlookPST.Model;
using Microsoft.Office.Interop.Outlook;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace OutlookPST
{


    public class Download
    {

        public Download() { }

        public static async Task<bool> MainDownload(object sender, ToastForm toast)
        {

            bool resultat;
            List<string> lstkeyName = new List<string>();
            StocObjInfoUser userStockobjInfo = new StocObjInfoUser();

            Cursor oldCursor = F.Cursor.Current;
            // Recuperation du BackgroundWorker remonté par l'événement.
            Const.worker = (System.ComponentModel.BackgroundWorker)sender;
            try
            {
                F.Cursor.Current = F.Cursors.WaitCursor;
                Const.cancellationTokenSourceForDownloads = new CancellationTokenSource();
                System.Threading.Thread.Sleep(Const.tempo);

                Tools.LogMessage(string.Format("Début de traitement : Boîte {0}_{1}.", Const.param.btDemande, Const.param.anneeDemande), null, false, MethodBase.GetCurrentMethod().Name);

                Const.worker.ReportProgress(1, MessageClass.ToastsMessage_Recup_Acces_Info);
                Const.ToastsMessage = MessageClass.ToastsMessage_Recup_Acces_Info;

                Const.param.userDemande = Const.userlogon;
                // Etape 1 : Est ce une délégation ?
                if (!(Const.param.btDemande_samAccountName.ToLower().Contains("_but")) && (Const.param.btDemande_samAccountName.ToLower()!= Const.userlogon))
                {
                    Const.param.userDemande = Const.param.btDemande_samAccountName.ToLower();
                    Tools.LogMessage("Accès en délégation.", null, false, MethodBase.GetCurrentMethod().Name);
                }

                // Etape 2 : Récupération du Fichier des clé depuis le NAS
                resultat = NASClass.GetNASInformation(userStockobjInfo, Const.param.userDemande);
                if (!resultat)
                {
                    // Si l'utilisateur n'a pas accés au nas ou n'a pas de clés actif
                    Const.param.archiveDemande = false;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Indisponible, MessageClass.message_NAS_PasDeCle);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Indisponible, MessageClass.message_NAS_PasDeCle);
                    Const.param.archiveDemande = true;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Indisponible, MessageClass.message_NAS_PasDeCle);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Indisponible, MessageClass.message_NAS_PasDeCle);
                    return false;
                }

                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;

                System.Threading.Thread.Sleep(Const.tempo);
                Tools.LogMessage("NAS : Clés d'authentifications récupérées.", null, false, MethodBase.GetCurrentMethod().Name);


                // Etape 3 : Donwload du S3 du PST
                Const.worker.ReportProgress(5, MessageClass.ToastsMessage_Recup_List_Data);
                Const.ToastsMessage = MessageClass.ToastsMessage_Recup_List_Data;

                // S3ListBucketContents : Identification des fichiers a récupérer depuis stockobj suivant 2 régles de nommage
                // -> Nomme (1) Bal Utilisateur : samAccountName\samAccountName_Année.7z ou samAccountName\samAccountName_Année_archive.7z
                // -> Nomme (2) Bal UT (Unité de travail) : samAccountName\NomBal_Année.7z ou samAccountName\NomBal_Année_archive.7z
                //      -> NomBal = Const.param.btDemande jusqu'a @

                // BAL User
                string nameForBALUser = string.IsNullOrEmpty(Const.param.btDemande_samAccountName) ? Environment.GetEnvironmentVariable(variable: "Username").ToLower() + "_" + Const.param.anneeDemande : Const.param.btDemande_samAccountName + "_" + Const.param.anneeDemande;               
                // BAL UT norme 1
                string nameForBUT = Const.param.btDemande.Contains("@") ? Const.param.btDemande.Substring(0, Const.param.btDemande.IndexOf("@")) + "_" + Const.param.anneeDemande : Const.param.btDemande + "_" + Const.param.anneeDemande;
                // BAL UT norme 2
                string nameForBUT2 = string.IsNullOrEmpty(Const.param.btDemande_mailNickName) ? string.Empty: Const.param.btDemande_mailNickName+"_" + Const.param.anneeDemande;

                // Folder stockobj racine des fichiers 
                string racineFolder = string.IsNullOrEmpty(Const.param.btDemande_samAccountName) ? Environment.GetEnvironmentVariable(variable: "Username").ToLower() : Const.param.btDemande_samAccountName;
                Task<bool> resultatS3L = S3Function.S3ListBucketContents(userStockobjInfo, racineFolder, nameForBALUser, nameForBUT, nameForBUT2, lstkeyName);
                resultatS3L.Wait(Const.cancellationTokenSourceForDownloads.Token);
                if (!resultatS3L.Result)
                {
                    Const.param.archiveDemande = false;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.statutRubanDemande, Const.param.commentaire);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.statutPSTDemande, Const.param.commentaire);
                    Const.param.archiveDemande = true;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.statutRubanDemande, Const.param.commentaire);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.statutPSTDemande, Const.param.commentaire);
                    return false;
                }

                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;

                if (lstkeyName.Count <= 0)
                {
                    // Message d'information
                    System.Windows.Forms.MessageBox.Show(MessageClass.Commentaire_PasDeFichier, "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    // Actualisation des flg
                    Const.param.archiveDemande = false;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Indisponible, MessageClass.Commentaire_PasDeFichier);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Indisponible, MessageClass.Commentaire_PasDeFichier);

                    Const.param.archiveDemande = true;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Indisponible, MessageClass.Commentaire_PasDeFichier);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Indisponible, MessageClass.Commentaire_PasDeFichier);
                    return false;
                }

                System.Threading.Thread.Sleep(Const.tempo);

                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;

                if (lstkeyName.Count >2)
                {
                    // Message d'information
                    System.Windows.Forms.MessageBox.Show(MessageClass.Commentaire_TropDeFichier, "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    // Actualisation des flg
                    Const.param.archiveDemande = false;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_TropDeFichier);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_TropDeFichier);

                    Const.param.archiveDemande = true;
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_TropDeFichier);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_TropDeFichier);
                    return false;
                }

                System.Threading.Thread.Sleep(Const.tempo);

                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;

                // Etape 4 : Donwload du S3 du PST
                Const.worker.ReportProgress(10, MessageClass.ToastsMessage_Prepa_DownLoad);
                Const.ToastsMessage = MessageClass.ToastsMessage_Prepa_DownLoad;

                // Definir depuis la liste des obj stockobj récupéré le fichier archive et le principal
                //-Le nom du fichier de la BAL Principale dans stocobj
                //-Le  nom du fichier de l’archive de la Bal Principale dans stocobj 
                //-prefix = Racine du fichier

                string KeyNameDataPrincipal = "";
                string KeyNameDataArchivePrincipal = "";
                string prefix = "";

                // Ctrl et identification 
                foreach (string itemKeyName in lstkeyName)
                {
                    prefix = itemKeyName.Substring(0, itemKeyName.LastIndexOf(@"/"));
                    if (itemKeyName.ToLower().Contains("_archive.7z"))
                    {
                        // Erreur si 2 fichiers identifié comme archive
                        if (!string.IsNullOrEmpty(KeyNameDataArchivePrincipal))
                        {
                            // Message d'information
                            System.Windows.Forms.MessageBox.Show(MessageClass.Commentaire_TropDeFichierBUT, "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                            // Actualisation des flg
                            Const.param.archiveDemande = false;
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_TropDeFichierBUT);
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_TropDeFichierBUT);

                            Const.param.archiveDemande = true;
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_TropDeFichierBUT);
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_TropDeFichierBUT);
                            return false;
                        }
                        //Nom du fichier Archive
                        KeyNameDataArchivePrincipal = itemKeyName;
                    }
                    else
                    {
                        // Erreur si 2 fichiers identifié comme Boite principale
                        if (!string.IsNullOrEmpty(KeyNameDataPrincipal))
                        {
                            // Message d'information
                            System.Windows.Forms.MessageBox.Show(MessageClass.Commentaire_TropDeFichierBT, "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                            // Actualisation des flg
                            Const.param.archiveDemande = false;
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_TropDeFichierBT);
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_TropDeFichierBT);

                            Const.param.archiveDemande = true;
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_TropDeFichierBT);
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_TropDeFichierBT);
                            return false;
                        }
                        //Nom du fichier pour la BT principal
                        KeyNameDataPrincipal = itemKeyName;
                    }

                }


                System.Threading.Thread.Sleep(Const.tempo);
                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;


                Tools.LogMessage(string.Format("Traitement : Boîte aux lettres principale. Boîte {0}_{1}.", Const.param.btDemande, Const.param.anneeDemande), null, false, MethodBase.GetCurrentMethod().Name);
                // Traitement de la boite Principal
                Const.param.archiveDemande = false;
                if (string.IsNullOrEmpty(KeyNameDataPrincipal))
                {
                    KeyNameDataPrincipal = string.Format("{0}/{1}", Const.param.btDemande_samAccountName, nameForBALUser+ ".7z"); //string.Format("{0}/{1}_{2}.7z", prefix, Const.param.btDemande_samAccountName, Const.param.anneeDemande);
                }

                var resultat_Bal = await SsDownload(Const.worker, userStockobjInfo, KeyNameDataPrincipal, toast, Const.param, 40);
                if (!resultat_Bal)
                {
                    if (Const.param.statutRubanDemande != Statut_ruban.Indisponible)
                    {
                        if (Const.cancellationTokenSourceForDownloads.IsCancellationRequested && Const.downloadsInUseThread)
                        {
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Annule, MessageClass.message_StocObj_Download_Cancel);
                        }
                        else
                        {
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_HS_Telechargement);
                        }
                    }
                    if (Const.param.statutPSTDemande != Statut_PST.Indisponible)
                    {
                        if (Const.cancellationTokenSourceForDownloads.IsCancellationRequested && Const.downloadsInUseThread)
                        {
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Annule, MessageClass.message_StocObj_Download_Cancel);
                        }
                        else
                        {
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_HS_Telechargement);
                        }
                    }
                }

                Tools.LogMessage(string.Format("Traitement : Boîte aux lettres principale. Boîte {0}_{1} statut {2}.", Const.param.btDemande, Const.param.anneeDemande,Const.param.statutPSTDemande.ToString()), null, false, MethodBase.GetCurrentMethod().Name);

                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;

                Tools.LogMessage(string.Format("Traitement : Boîte aux lettres Archive. Boîte {0}_{1}.", Const.param.btDemande, Const.param.anneeDemande), null, false, MethodBase.GetCurrentMethod().Name);
                /// ********* Traitement de la boite ARCHIVE ***************************
                Const.param.archiveDemande = true;
                if (string.IsNullOrEmpty(KeyNameDataArchivePrincipal))
                {
                    KeyNameDataArchivePrincipal = string.Format("{0}/{1}", Const.param.btDemande_samAccountName, nameForBALUser + "_archive.7z"); //string.Format("{0}/{1}_{2}_archive.7z", prefix, Const.param.btDemande_samAccountName, Const.param.anneeDemande);
                }

                Const.worker.ReportProgress(50, MessageClass.ToastsMessage_DownLoad_BAL_Archive);
                Const.ToastsMessage = MessageClass.ToastsMessage_DownLoad_BAL_Archive;
                System.Threading.Thread.Sleep(1000);

                var resultat_Archive = await SsDownload(Const.worker, userStockobjInfo, KeyNameDataArchivePrincipal, toast, Const.param, 70);
                if (!resultat_Archive)
                {
                    if (Const.param.statutRubanDemande != Statut_ruban.Indisponible)
                    {
                        if (Const.cancellationTokenSourceForDownloads.IsCancellationRequested && Const.downloadsInUseThread)
                        {
                            Const.param.statutRubanDemande = Statut_ruban.Annule;
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Annule, MessageClass.message_StocObj_Download_Cancel);
                        }
                        else
                        {
                            Const.param.statutRubanDemande = Statut_ruban.HS;
                            Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_HS_Telechargement_Archive);
                        }
                    }
                    if (Const.param.statutPSTDemande != Statut_PST.Indisponible)
                    {
                        if (Const.cancellationTokenSourceForDownloads.IsCancellationRequested && Const.downloadsInUseThread)
                        {
                            Const.param.statutPSTDemande = Statut_PST.Annule;
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Annule, MessageClass.message_StocObj_Download_Cancel);
                        }
                        else
                        {
                            Const.param.statutPSTDemande = Statut_PST.HS;
                            Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_HS_Telechargement_Archive);
                        }
                    }
                }

                System.Threading.Thread.Sleep(Const.tempo);
                Tools.LogMessage(string.Format("Traitement : Boîte aux lettres Archive. Boîte {0}_{1} statut {2}.", Const.param.btDemande, Const.param.anneeDemande, Const.param.statutPSTDemande.ToString()), null, false, MethodBase.GetCurrentMethod().Name);
            }
            catch (System.Exception ex)
            {
                // Etape 5 (si necessaire): Dans le Flag Ajout de l'erreur
            }
            finally
            {
                F.Cursor.Current = oldCursor;
                // Etape 6 : Mise a niveau du statue flag
                Const.worker.ReportProgress(100, MessageClass.ToastsMessage_Termine);
                Const.ToastsMessage = MessageClass.ToastsMessage_Termine;
            }

            return true;
        }


        public static async Task<bool> SsDownload(System.ComponentModel.BackgroundWorker worker, StocObjInfoUser userStockobjInfo, string KeyName, ToastForm toast, ParamDemande param, int processBar)
        {
            string fullfilename7Z;
            string fullpathSeptZipDest;
            string fullpathSeptZipToPST;
            long lengthFile;
            bool resultat;

            try
            {
                //// Etape 5 : Donwload du S3 du PST - Non Archive
                if (param.archiveDemande)
                {
                    Const.worker.ReportProgress(processBar, MessageClass.ToastsMessage_DownLoad_BAL_Archive);
                    Const.ToastsMessage = MessageClass.ToastsMessage_DownLoad_BAL_Archive;
                    System.Threading.Thread.Sleep(1000);
                }
                else
                {
                    Const.worker.ReportProgress(processBar, MessageClass.ToastsMessage_DownLoad_BAL_Principal);
                    Const.ToastsMessage = MessageClass.ToastsMessage_DownLoad_BAL_Principal;
                    System.Threading.Thread.Sleep(1000);
                }

                System.Threading.Thread.Sleep(Const.tempo);
                // Arret de l'utilisateur : Annulation 
                if (Const.stopThread)
                    return true;

                Const.downloadsInUseThread = true;

                Task<bool> reponse = S3Function.S3Download(userStockobjInfo, KeyName);
                reponse.Wait(Const.cancellationTokenSourceForDownloads.Token);
                //var resultatS3 = await S3Function.S3Download(userStockobjInfo, param, KeyName).Wait(Tools.sourceS3.Token);                
                if (!reponse.Result)
                {
                    string infoErreur = Const.param.statutRubanDemande == Statut_ruban.Indisponible ? "Non Disponible" : "Erreur dans le téléchargement depuis stocObj.";
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.statutRubanDemande, infoErreur);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Const.param.statutPSTDemande, infoErreur);
                    return false;
                }

                Const.downloadsInUseThread = false;
                System.Threading.Thread.Sleep(Const.tempo);

                if (param.archiveDemande)
                {
                    Const.worker.ReportProgress(processBar + 20, MessageClass.ToastsMessage_Decompress_BAL_Archive);
                    Const.ToastsMessage = MessageClass.ToastsMessage_Decompress_BAL_Archive;
                }
                else
                {
                    Const.worker.ReportProgress(processBar + 20, MessageClass.ToastsMessage_Decompress_BAL_Principal);
                    Const.ToastsMessage = MessageClass.ToastsMessage_Decompress_BAL_Principal;
                }

                // Verification de la presence du fichier après traitement
                // transforme le nom dans Stockobj=KeyName=> Chemin du fichier downloadé dans le ..\Temp7z\NomFichier_annee.7z ou ..\Temp7z\NomFichier_annee_archive.7z 
                fullfilename7Z = Path.Combine(Const.tmp7z, KeyName.Replace(@"/", @"\"));
                if (!File.Exists(fullfilename7Z))
                {
                    Tools.LogMessage(string.Format("Erreur : Fichier {0} téléchargé non trouvé.", fullfilename7Z), null, false, MethodBase.GetCurrentMethod().Name);
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_PasDeFichier_7z);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_PasDeFichier_7z);
                    return false;
                }

                // Chemin pour l"extract du 7z dans un sous rep au nom de la samAccountName
                fullpathSeptZipDest = Path.Combine(Const.tmp7z, param.btDemande_samAccountName);
                // Etape 6 : Dézippage du 7z en PST -> 
                resultat = _7zClass.GetSeptZip(fullfilename7Z, fullpathSeptZipDest);
                if (!resultat)
                {
                    Tools.LogMessage(string.Format("Erreur : lors de l'extract du fichier : {0}", fullfilename7Z), null, false, MethodBase.GetCurrentMethod().Name);
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_Erreur_7z);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_Erreur_7z);
                    return false;
                }

                System.Threading.Thread.Sleep(Const.tempo);

                if (param.archiveDemande)
                {
                    // Controle de la présence du fichier .pst et renommage
                    Const.worker.ReportProgress(processBar + 5, MessageClass.ToastsMessage_Mise_En_Place_BAL_Archive);
                    Const.ToastsMessage = MessageClass.ToastsMessage_Mise_En_Place_BAL_Archive;
                }
                else
                {
                    // Controle de la présence du fichier .pst et renommage
                    Const.worker.ReportProgress(processBar + 5, MessageClass.ToastsMessage_Mise_En_Place_BAL_Principal);
                    Const.ToastsMessage = MessageClass.ToastsMessage_Mise_En_Place_BAL_Principal;
                }


                // Ctr que le fichier exporter du 7z est celui attendu
                string folderExtract= Path.Combine(Const.tmp7z, param.btDemande_samAccountName);
                // Evol 10-2025: Nom BAL UT ou SamAccountName pour le nom du fichier pst
                // Nommage du fichier dans le 7z avec SamAccount
                // Nommage BAL User
                // le fichier basé sur samAccountName => Bal Utilistateur
                string nommageBALUser = string.IsNullOrEmpty(param.btDemande_samAccountName) ? Environment.GetEnvironmentVariable(variable: "Username") + "_" + param.anneeDemande : param.btDemande_samAccountName + "_" + param.anneeDemande;
                nommageBALUser = nommageBALUser.ToLower();
                // Nommage BAL UT
                // le fichier peut etre a base du nom de boite jusqu'a @
                string nommageBUT = param.btDemande.Contains("@") ? param.btDemande.Substring(0, param.btDemande.IndexOf("@")) + "_" + param.anneeDemande : param.btDemande + "_" + param.anneeDemande;
                nommageBUT= nommageBUT.ToLower();
                // Nommage BAL UT norme 2
                string nommageBUT2 = string.IsNullOrEmpty(Const.param.btDemande_mailNickName) ? string.Empty : Const.param.btDemande_mailNickName + "_" + Const.param.anneeDemande;
                if (!string.IsNullOrEmpty(nommageBUT2)) 
                    nommageBUT2 = nommageBUT2.ToLower();

                string extractNameFile = param.archiveDemande? nommageBALUser + "_archive.pst": nommageBALUser + ".pst";
                
                if (param.archiveDemande)
                {
                    if (File.Exists(Path.Combine(folderExtract, nommageBALUser + "_archive.pst")))
                    {
                        extractNameFile = nommageBALUser + "_archive.pst";
                    }

                    else if (File.Exists(Path.Combine(folderExtract, nommageBUT + "_archive.pst")))
                    {
                        extractNameFile = nommageBUT + "_archive.pst";
                    }

                    else if ((!string.IsNullOrEmpty(nommageBUT2)) && File.Exists(Path.Combine(folderExtract, nommageBUT2 + "_archive.pst")))
                    {
                        extractNameFile = nommageBUT2 + "_archive.pst";
                    }
                }
                else
                {
                    if (File.Exists(Path.Combine(folderExtract, nommageBALUser + ".pst")))
                    {
                        extractNameFile = nommageBALUser + ".pst";
                    }

                    else if (File.Exists(Path.Combine(folderExtract, nommageBUT + ".pst")))
                    {
                        extractNameFile = nommageBUT + ".pst";
                    }

                    else if ((!string.IsNullOrEmpty(nommageBUT2)) && File.Exists(Path.Combine(folderExtract, nommageBUT2 + ".pst")))
                    {
                        extractNameFile = nommageBUT2 + ".pst";
                    }
                }


                fullpathSeptZipToPST = Path.Combine(folderExtract, extractNameFile);
                //fullpathSeptZipToPST = fullpathSeptZip.ToLower().Replace(".7z", ".pst");

                if (!File.Exists(fullpathSeptZipToPST))
                {
                    Tools.LogMessage(string.Format("Erreur : Fichier {0} extrait non trouvé.", fullpathSeptZipToPST), null, false, MethodBase.GetCurrentMethod().Name);
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_Erreur_NonPresent);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_Erreur_NonPresent);
                    return false;
                }

                lengthFile = new System.IO.FileInfo(fullpathSeptZipToPST).Length;
                // Si mauvais mot de passe Taille == 0 : Controle le dezippage
                if (lengthFile == 0)
                {
                    Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_Erreur_MDP_7z);
                    Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_Erreur_MDP_7z);
                }
                else
                {

                    // Rename et deplace dans le repertoire des PST le fichier pst
                    string full7zsrcNew = string.Format(@"{0}\{1}", Const.path_Local_Flg_PST, Traitement.getNamePST(param.btDemande, param.anneeDemande, param.archiveDemande));
                    if (!MiseEnPlacePST(full7zsrcNew, fullpathSeptZipToPST, fullpathSeptZipDest))
                    {
                        //// Maj flag pour le PST de la boite au lettre
                        Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.HS, MessageClass.Commentaire_Erreur_MiseEnPlace);
                        Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.HS, MessageClass.Commentaire_Erreur_MiseEnPlace);
                        Const.stopThread = true;
                    }
                    else
                    {
                        //// Maj flag pour le PST de la boite au lettre
                        Traitement.createOrUpdateFlagRuban(Const.lstFlagRuban, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_ruban.Mount, MessageClass.Commentaire_Disponible);
                        Traitement.createOrUpdateFlagPST(Const.lstFlagPST, Const.param.btDemande, Const.param.anneeDemande, Const.param.archiveDemande, Statut_PST.Mount, MessageClass.Commentaire_Disponible);
                    }
                    
                }

                System.Threading.Thread.Sleep(Const.tempo);
                return true;
            }
            catch 
            {

                return false;
            }
        }


        private static bool MiseEnPlacePST(string full7zsrcNew, string fullpathSeptZipToPST,string fullpathSeptZipDest)
        {
            try
            {
            if (File.Exists(full7zsrcNew))
                    File.Delete(full7zsrcNew);
                File.Move(fullpathSeptZipToPST, full7zsrcNew);
                if (Directory.Exists(fullpathSeptZipDest))
                    Directory.Delete(fullpathSeptZipDest,true);
                return true;
            }
            catch (System.Exception ex) 
            {
                MessageBox.Show(MessageClass.message_DeplacementPSTImpossible, "Information :",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                Tools.LogMessage("Erreur de déplacement du fichier PST.", ex, false, MethodBase.GetCurrentMethod().Name);
                // Nettoyage si erreur
                //if (Directory.Exists(fullpathSeptZipDest))
                //    Directory.Delete(fullpathSeptZipDest, true);
                return false;
            }
        }



        
    }
}
