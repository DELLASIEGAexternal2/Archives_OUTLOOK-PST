using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPST
{
    public enum Statut_ruban
    {
        None,
        Encours,
        Mount,
        OK,
        HS,
        Indisponible,
        Partial,
        Annule
    }
    /// <summary>
    /// Class de mise en place des flag dans le sous dossier PST pour le profil
    /// </summary>
    public class FlgRubanClass
    {
        public string MailBoxName { get; set; }

        public string Annee { get; set; }

        public bool Archive { get; set; }

        public Statut_ruban Statut { get; set; }

        public string Commentaire { get; set; }

    }
   
}
