using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookPST
{
    public enum Statut_PST
    {
        None,
        Encours,
        Mount,
        HS,
        Indisponible,
        Annule
    }
    public class FlgPSTClass
    {
        public string MailBoxName { get; set; }

        public string Annee { get; set; }

        public bool Archive { get; set; }

        public Statut_PST Statut { get; set; }

        public string Commentaire { get; set; }

    }
}
