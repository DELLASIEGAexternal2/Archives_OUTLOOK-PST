using Amazon.S3.Model;
using Amazon.S3;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Amazon.Runtime;
using Amazon;
using System.Runtime.CompilerServices;
using Amazon.S3.IO;
using System.Windows.Forms;
using OutlookPST.Model;
using System.Diagnostics;
using Amazon.S3.Model.Internal.MarshallTransformations;
using Amazon.Auth.AccessControlPolicy;
using System.Reflection;
using System.Drawing.Printing;
using Amazon.S3.Transfer;
using System.IO;
using System.Security.Cryptography;
using System.Data.SqlTypes;

namespace OutlookPST
{
    internal class S3Function
    {
        /// <summary>
        /// S3ListBucketContents
        /// </summary>
        /// <param name="userStockobjInfo"></param>
        /// <param name="param"></param>
        /// <param name="lstkeyName"></param>
        /// <returns></returns>
        public static async Task<bool> S3ListBucketContents(StocObjInfoUser userStockobjInfo, string racineFolder, string nameFileBAL_User , string nameFileBAL_UT, string nameFileBAL_UT2, List<string> lstkeyName)
        {
            try
            {
                // Create an Amazon S3 client object. 
                var config = new AmazonS3Config();
                config.ForcePathStyle = true;

                config.RegionEndpoint = RegionEndpoint.GetBySystemName(Const.stockObj_RegionEndpoint);

                config.ServiceURL = Const.stockObj_ServiceURL;
                config.UseHttp = false;

                var credentials = new BasicAWSCredentials(userStockobjInfo.StocObjAccessKey, userStockobjInfo.StocObjSecretKey);
                var s3client = new AmazonS3Client(credentials, config);

                string bucketName = Const.stockObj_BucketName;


                // Remplie la lst des fichiers disponible depuis la racinefolder avec un nom=elementRecherheNommageBALUser ou nom=elementRecherhNommageBUT
                bool successLstContent = S3Function.ListBucketContents(s3client, bucketName, racineFolder, nameFileBAL_User, nameFileBAL_UT, nameFileBAL_UT2, lstkeyName);
                if (!successLstContent)
                {

                    Tools.LogMessage(string.Format("Echec : Liste fichier(s) non disponible pour l'année {0}.", Const.param.anneeDemande), null, false, MethodBase.GetCurrentMethod().Name);
                    return false;
                }
                else
                {
                    Tools.LogMessage(string.Format("Success : Liste fichier(s) {0} disponible pour l'année {1}.", lstkeyName.Count,Const.param.anneeDemande), null, false, MethodBase.GetCurrentMethod().Name);
                }
                return true;
            }
            catch (Exception ex)
            {
                Const.param.statutRubanDemande = Statut_ruban.Indisponible;
                Const.param.statutPSTDemande= Statut_PST.Indisponible;
                //System.Windows.Forms.MessageBox.Show($"Echec.\n" + ex.Message, "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Tools.LogMessage(string.Format("Erreur dans la recherche des fichiers disponibles : {0} -SAM : {1}", ex.Message,racineFolder), ex, false, MethodBase.GetCurrentMethod().Name);
                return false;
            }

        }


        /// <summary>
        /// Shows how to list the objects in an Amazon S3 bucket.
        /// </summary>
        /// <param name="client">An initialized Amazon S3 client object.</param>
        /// <param name="bucketName">The name of the bucket for which to list
        /// <param name="racineFolder"> Racine de recherche
        /// <param name="nomBAL"> Nom du fichier pour le BAL Utilisateur
        /// <param name="nomBAL_UT"> Nom du fichier pour le BAL Unité de Travail
        /// <param name="nomBAL_UT2"> Nom du fichier pour le BAL Unité de Travail AD mailNickName -> norme du 15/10/2025
        /// the contents.</param>
        /// <returns>A boolean value indicating the success or failure of the
        /// copy operation.</returns>
        private static bool ListBucketContents(IAmazonS3 client, string bucketName, string racineFolder, string nameFileBAL_User, string nameFileBAL_UT, string nameFileBAL_UT2, List<string> lstkeyName)
        {
            string prefix = racineFolder + "/";
            try
            {
                var request = new ListObjectsV2Request
                {
                    BucketName = bucketName,
                    MaxKeys = 40,
                    Prefix = racineFolder + "/"
                };                

                ListObjectsV2Response response;

                bool jePrend=false;

                do
                {
                    if (string.IsNullOrEmpty(nameFileBAL_UT2))
                        Tools.LogMessage(string.Format("Nom fichier à trouver : {0} ou {1}", nameFileBAL_User, nameFileBAL_UT), null, false, MethodBase.GetCurrentMethod().Name);
                    else
                        Tools.LogMessage(string.Format("Nom fichier à trouver : {0} ou {1} ou {2}", nameFileBAL_User, nameFileBAL_UT, nameFileBAL_UT2 ), null, false, MethodBase.GetCurrentMethod().Name);
                    
                    response = client.ListObjectsV2(request);

                    response.S3Objects
                        .ForEach(obj =>
                        {
                            jePrend = false;
                            if (obj.Key.ToLower().Contains(nameFileBAL_User.ToLower() + ".7z"))
                                jePrend=true;

                            if (obj.Key.ToLower().Contains(nameFileBAL_User.ToLower() + "_archive.7z"))
                                jePrend = true;

                            if (obj.Key.ToLower().Contains(nameFileBAL_UT.ToLower() + ".7z"))
                                jePrend = true;

                            if (obj.Key.ToLower().Contains(nameFileBAL_UT.ToLower() + "_archive.7z"))
                                jePrend = true;

                            if (!string.IsNullOrEmpty(nameFileBAL_UT2) && obj.Key.ToLower().Contains(nameFileBAL_UT2.ToLower() + ".7z"))
                                jePrend = true;

                            if (!string.IsNullOrEmpty(nameFileBAL_UT2) && obj.Key.ToLower().Contains(nameFileBAL_UT2.ToLower() + "_archive.7z"))
                                jePrend = true;

                            //if (obj.Key.Contains("_" + Const.param.anneeDemande))
                            // Evo : regle de nommage
                            //if (obj.Key.ToLower().Contains(nameFileBAL_User.ToLower() + ".7z") 
                            //    || obj.Key.ToLower().Contains(nameFileBAL_User.ToLower() + "_archive.7z")
                            //    || obj.Key.ToLower().Contains(nameFileBAL_UT.ToLower() + ".7z") 
                            //    || obj.Key.ToLower().Contains(nameFileBAL_UT.ToLower() + "_archive.7z")
                            //    || obj.Key.ToLower().Contains(nameFileBAL_UT2.ToLower() + ".7z")
                            //    || obj.Key.ToLower().Contains(nameFileBAL_UT2.ToLower() + "_archive.7z")
                            //    )
                            if (jePrend)
                            {
                                //System.Windows.Forms.MessageBox.Show(obj.Key);
                                lstkeyName.Add(obj.Key);
                                Tools.LogMessage(string.Format("Fichier : prise en compte : {0} - Taille : {1}", obj.Key, obj.Size.HasValue ? obj.Size.Value.ToString() : "0"), null, false, MethodBase.GetCurrentMethod().Name);
                            }
                            else
                            {
                                Tools.LogMessage(string.Format("Fichier : non prise en compte : {0} - Taille : {1}", obj.Key, obj.Size.HasValue?obj.Size.Value.ToString():"0"), null, false, MethodBase.GetCurrentMethod().Name);
                            }
                        }

                        //System.Windows.Forms.MessageBox.Show($"{obj.Key,-35}{obj.LastModified.ToShortDateString(),10}{obj.Size,10}")
                        );

                    // If the response is truncated, set the request ContinuationToken
                    // from the NextContinuationToken property of the response.
                    request.ContinuationToken = response.NextContinuationToken;
                }
                while (response.IsTruncated.HasValue && response.IsTruncated.Value);

                return true;
            }
            catch (AmazonS3Exception ex)
            {
                // Acces refusé
                if (ex.HResult==-2146233088)
                {
                    Const.param.statutRubanDemande = Statut_ruban.Indisponible;
                    Const.param.statutPSTDemande = Statut_PST.Indisponible;
                    Const.param.commentaire = "Non disponible";
                    Tools.LogMessage(string.Format("Non disponible dans le serveur stocobj lors de la recherche des éléménts de messageries. Prefix :{0}", prefix), ex, false, MethodBase.GetCurrentMethod().Name);

                }
                else
                {
                    Const.param.statutRubanDemande = Statut_ruban.HS;
                    Const.param.statutPSTDemande = Statut_PST.HS;
                    Const.param.commentaire = ex.Message;
                    Tools.LogMessage(string.Format("Erreur retourné par le serveur stocobj lors de la recherche des éléménts de messageries. Prefix :{0}",prefix), ex, false, MethodBase.GetCurrentMethod().Name);
                }
                //System.Windows.Forms.MessageBox.Show($"Erreur retourné par le serveur. Message:'{ex.Message}' .", "Erreur :", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }


        /// <summary>
        /// Shows how to download an object from an Amazon S3 bucket to the
        /// local computer.
        /// </summary>
        /// <param name="client">An initialized Amazon S3 client object.</param>
        /// <param name="bucketName">The name of the bucket where the object is
        /// currently stored.</param>
        /// <param name="objectName">The name of the object to download.</param>
        /// <param name="filePath">The path, including filename, where the
        /// downloaded object will be stored.</param>
        /// <returns>A boolean value indicating the success or failure of the
        /// download process.</returns>
        public static async Task<bool> DownloadObjectFromBucketAsync(
            IAmazonS3 client,
            string bucketName,
            string objectName,
            string filePath)
        {
            // Create a GetObject request
            GetObjectRequest request = new GetObjectRequest
            {
                BucketName = bucketName,
                Key = objectName,
            };

            

            try
            {                
                using (GetObjectResponse response = client.GetObject(request))
                {
                    response.WriteObjectProgressEvent += Response_WriteObjectProgressEvent;
                    // Save object to local file
                    await response.WriteResponseStreamToFileAsync($"{filePath}\\{objectName}", false, Const.cancellationTokenSourceForDownloads.Token);
                    return response.HttpStatusCode == System.Net.HttpStatusCode.OK;                    
                }
            }
            catch 
            {
                //System.Windows.Forms.MessageBox.Show($"Error saving {objectName}: {ex.Message}", "Erreur :", System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Error);
                throw;
            }
        }

        public static async Task<bool> DownloadObjectFromBucketAsync2(
        IAmazonS3 client,
        string bucketName,
        string objectName,
        string filePath)
            {

                var fileTransferUtility = new TransferUtility(client);
            

                try
                {
                    var request = new TransferUtilityDownloadRequest
                    {
                        BucketName = bucketName,
                        FilePath = Path.Combine(filePath,objectName),
                        Key = objectName                        
                };
                

                await fileTransferUtility.DownloadAsync(request, Const.cancellationTokenSourceForDownloads.Token);
                    return true;
                }
                catch (AmazonS3Exception s3Exception)
                {
                    MessageBox.Show(s3Exception.Message, "Error 102",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error 103",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
        }


        /// <summary>
        /// Shows how to download an object from an Amazon S3 bucket to the
        /// local computer.
        /// </summary>
        /// <param name="client">An initialized Amazon S3 client object.</param>
        /// <param name="bucketName">The name of the bucket where the object is
        /// currently stored.</param>
        /// <param name="objectName">The name of the object to download.</param>
        /// <param name="filePath">The path, including filename, where the
        /// downloaded object will be stored.</param>
        /// <returns>A boolean value indicating the success or failure of the
        /// download process.</returns>
        public static bool DownloadObjectFromBucket(
            IAmazonS3 client,
            string bucketName,
            string objectName,
            string filePath)
        {
            // Create a GetObject request
            var request = new GetObjectRequest
            {
                BucketName = bucketName,
                Key = objectName,
            };

            try
            {                

                using (var response = client.GetObject(request))
                {                    
                    response.WriteObjectProgressEvent += Response_WriteObjectProgressEvent;
                    // Save object to local file
                    response.WriteResponseStreamToFile($"{filePath}\\{objectName}");                   
                    return response.HttpStatusCode == System.Net.HttpStatusCode.OK;
                }
            }
            catch (AmazonS3Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error saving {objectName}: {ex.Message}", "Erreur :", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        private static void Response_WriteObjectProgressEvent(object sender, WriteObjectProgressArgs e)
        {
            double transferredBytesInMega= e.TransferredBytes/1048576;
            double totalBytesInMega= e.TotalBytes/1048576;
            Const.worker.ReportProgress(e.PercentDone, "Téléchargement");
            //Const.ToastsMessage = $"Transfert : {e.TransferredBytes}/{e.TotalBytes} - Progression : {e.PercentDone}%";            
            if (Const.param.archiveDemande)
            {
                Const.ToastsMessage = string.Format(MessageClass.ToastsMessage_BAL_Archive, transferredBytesInMega.ToString("0"), totalBytesInMega.ToString("0"), e.PercentDone.ToString());
            }
            else
            {
                Const.ToastsMessage = string.Format(MessageClass.ToastsMessage_BAL, transferredBytesInMega.ToString("0"), totalBytesInMega.ToString("0"),e.PercentDone.ToString());
            }
            

        }

        public static async Task<bool> S3Download(StocObjInfoUser userStockobjInfo,  string keyNameDemande)
        {
            string filePath = Const.tmp7z;
            try
            {
                // Create an Amazon S3 client object. The constructor uses the
                // default user installed on the system. To work with Amazon S3
                // features in a different AWS Region, pass the AWS Region as a
                // parameter to the client constructor.

                var config = new AmazonS3Config();                
                config.ForcePathStyle = true;
                

                config.RegionEndpoint = RegionEndpoint.GetBySystemName(Const.stockObj_RegionEndpoint);

                config.ServiceURL = Const.stockObj_ServiceURL;
                config.UseHttp = false;
                
                var credentials = new BasicAWSCredentials(userStockobjInfo.StocObjAccessKey, userStockobjInfo.StocObjSecretKey);
                var s3client = new AmazonS3Client(credentials,config);
                
                string bucketName = Const.stockObj_BucketName;




                //// Download an object from a bucket.
                //bool success =  S3Function.DownloadObjectFromBucket(s3client, bucketName, keyNameDemande, filePath);
                var success = await S3Function.DownloadObjectFromBucketAsync(s3client, bucketName, keyNameDemande, filePath);
                if (success)
                {
                    //System.Windows.Forms.MessageBox.Show($"Success download {keyName}.\n","Information",System.Windows.Forms.MessageBoxButtons.OK,System.Windows.Forms.MessageBoxIcon.Information);
                    Const.param.statutPSTDemande = Statut_PST.Mount;
                    //System.Windows.Forms.MessageBox.Show($"Success.\n", "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                }
                else
                {
                    //Console.WriteLine($"Sorry, could not download {keyName}.\n");
                    //System.Windows.Forms.MessageBox.Show($"Echec du download {keyName}.\n", "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    Const.param.statutPSTDemande = Statut_PST.HS;
                    System.Windows.Forms.MessageBox.Show($"Echec.\n", "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
                return true;
            }
            catch (Exception ex) {
                Const.param.statutPSTDemande = Statut_PST.HS;
                if (ex.Message== @"The specified key does not exist.")
                    {
                    Const.param.statutPSTDemande = Statut_PST.Indisponible;
                    Const.param.statutRubanDemande = Statut_ruban.Indisponible;                    
                }
                //System.Windows.Forms.MessageBox.Show($"Echec.\n"+ex.Message, "Information", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

    }
}
