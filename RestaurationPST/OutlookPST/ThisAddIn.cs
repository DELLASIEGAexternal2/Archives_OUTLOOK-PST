using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Reflection;

namespace OutlookPST
{
    public partial class ThisAddIn
    {
        public Outlook.Explorer currentExplorer = null;
        public string currentMailBox = "";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Ajout d'un événement sur le changement de séléction
            currentExplorer = this.Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(CurrentExplorer_Event_Mes_archives);            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Remarque : Outlook ne déclenche plus cet événement. Si du code
            //    doit s'exécuter à la fermeture d'Outlook (consultez https://go.microsoft.com/fwlink/?LinkId=506785)
        }

        public void CurrentExplorer_Event_Mes_archives()
        {
            try
            {
                
                // Actualisation des flag de quel Boite Mail 
                string crMailBox = Traitement.getCurrentMailBox();
                bool refresh = false;

                if (Const.currentMailBox != crMailBox)
                {                    
                    // Si la boite change => Mettre a jours la valeur de currentMailBox
                    Const.currentMailBox = crMailBox;
                    refresh = true;
                }

                // Recupération des PST dans la session si isConnected
                if (Tools.isConnected()!=Const.isConnected)
                {
                    Const.isConnected = Tools.isConnected();
                    Traitement.PSTInSession();
                    refresh = true;
                }

                if (refresh)
                {
                    // Est ce un Boite a archive 5 ou 10 ans
                    Traitement.isTenYear(crMailBox);

                    Traitement.refreshRuban(Globals.Ribbons.RibbonDownloadExplorer, Const.currentMailBox,Const.lstFlagRuban);
                }

            }
            catch (Exception ex)
            {
                Tools.LogMessage("Erreur lors de l'actualisation." + MethodBase.GetCurrentMethod().Name, ex, true);
            }
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
