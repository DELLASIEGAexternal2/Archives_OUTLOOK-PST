using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPST
{
    public class ParamDemande
    {
        //**************************************
        // Informations PUBLIC pour le téléchargement  
        //**************************************
        // Nom de la boite demandé
        public string btDemande { get; set; }
        //// Année Demandé
        public string anneeDemande { get; set; }
        //// Ext Archive si < mi 2023 
        public bool archiveDemande { get; set; }
        //// samAccountName de la boite demande
        public string btDemande_samAccountName { get; set; }
        //// mailNickName de la boite demande
        public string btDemande_mailNickName { get; set; }
        //// statut pour le PST pour la demande
        public Statut_PST statutPSTDemande { get; set; }
        //// statut pour le ruban suite pour la demande
        public Statut_ruban statutRubanDemande { get; set; }
        public string commentaire { get; set; }
        //// Utilisateur qui demande l'accès a la boîte
        public string userDemande { get; set; }
    }
}
