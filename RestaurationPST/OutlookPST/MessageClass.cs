using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPST
{
    public static class MessageClass
    {
        public static string ToastsMessage_Recup_Acces_Info = "Etape 1 : récupération des informations d'accés aux données.";
        public static string ToastsMessage_Recup_List_Data = "Etape 2 : recherche des éléments disponibles.";
        public static string ToastsMessage_Prepa_DownLoad = "Etape 3 : préparation des Téléchargements.";
        public static string ToastsMessage_DownLoad_BAL_Principal = "Etape 4 : téléchargement des données de la BAL.";
        public static string ToastsMessage_Decompress_BAL_Principal = "Etape 5 : décompression de l'archive.";
        public static string ToastsMessage_Mise_En_Place_BAL_Principal = "Etape 6 : Mise en place dans le répertoire.";
        public static string ToastsMessage_DownLoad_BAL_Archive = "Etape 7 : téléchargement des données de la BAL Archive.";
        public static string ToastsMessage_Decompress_BAL_Archive = "Etape 8 : décompression de l'archive.";
        public static string ToastsMessage_Mise_En_Place_BAL_Archive = "Etape 9 : Mise en place de l'archive dans le répertoire.";
        

        public static string ToastsMessage_Termine = "Etape 10 : terminée.";

        public static string ToastsMessage_BAL = "Transfert BAL : {0}/{1} Mo - Progression : {2}%";
        public static string ToastsMessage_BAL_Archive = "Transfert BAL ARCHIVE : {0}/{1} Mo - Progression : {2}%";


        public static string message_NAS_PasDeCle = "Pas de clé disponible dans le NAS.";
        public static string message_Commentaire_Start = "Téléchargement du {0}";

        public static string message_StocObj_Download_Cancel = "Opération Annulé par l'utilistateur. " + DateTime.Now.ToString();
        //public static string message_NAS_PasDeCle = "Pas de clé disponible dans le NAS.";

        public static string message_DeplacementPSTImpossible = "Impossible de déplacer le PST. \nDécharger le ou les archives de l'année que vous souhaitez installer \net relancer OUTLOOK. ";
        public static string message_PSTCharger = "Décharger le ou les archives de l'année que vous souhaitez installer \net relancer OUTLOOK. ";
        public static string message_ErrMise_En_Place = "Erreur lors de la mise en place.";

        public static string message_Choix = "Les données de votre boîte aux lettres sont déja disponibles en local, \n(Oui) -> Ouverture de l'archive. \n(Non) -> Re télécharger les archives.";
        public static string message_ChoixOffLigne = "Les données de votre boîte aux lettres sont déja disponibles en local, \n(OK) -> Ouverture de l'archive. \n(Annuler) -> Annuler l'opération.";

        public static string message_Traitement_En_Cours = "Traitement en cours. \nVeuillez attendre la fin du téléchargement ou interrompre le transfert.";
        public static string message_Traitement_Termine = "Installation {0} terminée.\n\n***** IMPORTANT *****\n\nNous vous recommandons de ne pas modifier vos archives, car aucune sauvegarde n’est effectuée sur ces données.";

        public static string Commentaite_Installe = "Installé le " + DateTime.Now.ToString();
        public static string Commentaite_PST_HS = "Problème de PST";
        public static string Commentaire_PasDeFichier = "Aucun fichier disponible.";
        public static string Commentaire_TropDeFichier = "Trop de fichier identifié pour l'année. max 2->(Boite + Archive)";
        public static string Commentaire_TropDeFichierBT = "Trop de fichier identifié pour la boîte principale.";
        public static string Commentaire_TropDeFichierBUT = "Trop de fichier identifié pour la boîte principale.";
        public static string Commentaire_HS_Telechargement = "Erreur dans le traitement de la boîte Archive.";
        public static string Commentaire_HS_Telechargement_Archive = "Erreur dans le traitement de la boîte Archive.";
        public static string Commentaire_PasDeFichier_7z = "Erreur : fichier.7z introuvable.";
        public static string Commentaire_Erreur_7z = "Erreur : problème de décompression 7z.";
        public static string Commentaire_Erreur_NonPresent = "Erreur : fichier introuvable.";
        public static string Commentaire_Erreur_MDP_7z = "Problème de mot de passe 7z.";
        public static string Commentaire_Erreur_MiseEnPlace = "Erreur : mise en place du PST impossible.";
        public static string Commentaire_Disponible = "Disponible.";
        public static string Commentaire_PST_ABS = "Le PST de la BAL n'est plus présent.\n";
        public static string Commentaire_PST_Archive_ABS = "Le PST de la BAL n'est plus présent.\n";
        public static string Commentaire_PST_ABS_Session = "Le PST de la BAL a été retiré de votre session OUTLOOK.\n";
        public static string Commentaire_PST_Archive_ABS_Session = "Le PST de la BAL Archive a été retiré de votre session OUTLOOK.\n";
        public static string Commentaire_PST_LimiteAtteinte = "Limite atteinte du nombre de PST.";
        public static string LimiteNombrePst = "Le nombre de PST est limité à {0}. \n - Veuillez décharger des PST de OUTLOOK. \n - Puis ré-ouvrir OUTLOOK, afin de pouvoir relancer le traitement.";
        public static string RechercheIndisponibleEnOffLigne = "\n* La recherche de la présence de vos PST dans la session n'est pas disponible en mode hors connexion.\nLe statut peut donc être incorrect.";
    }
}
