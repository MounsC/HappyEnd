using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace HappyEnd
{
    [SupportedOSPlatform("windows")]
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("Démarrage de l'extraction des emails Outlook...");
            
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string savePath = Path.Combine(desktopPath, "OutlookExport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss"));
                Directory.CreateDirectory(savePath);
                Console.WriteLine($"Dossier de sauvegarde créé : {savePath}");

                Console.WriteLine("Connexion à Outlook...");
                Type outlookAppType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookAppType == null)
                {
                    throw new System.Exception("Impossible de trouver Outlook sur cet ordinateur.");
                }
                
                dynamic outlookApp = Activator.CreateInstance(outlookAppType);
                dynamic outlookNamespace = outlookApp.GetNamespace("MAPI");
                
                Console.WriteLine("Connexion à Outlook réussie.");
                
                List<dynamic> allFolders = new List<dynamic>();
                
                int[] importantFolders = new int[]
                {
                    6,    // olFolderInbox - Boîte de réception
                    5,    // olFolderSentMail - Éléments envoyés
                    16,   // olFolderDrafts - Brouillons
                    3,    // olFolderDeletedItems - Éléments supprimés
                    4,    // olFolderOutbox - Boîte d'envoi
                    23    // olFolderJunk - Courrier indésirable
                };
                
                bool modeHorsLigne = true;
                Console.WriteLine("Mode hors ligne activé - extraction uniquement des dossiers principaux.");
                
                foreach (int folderType in importantFolders)
                {
                    try
                    {
                        dynamic folder = outlookNamespace.GetDefaultFolder(folderType);
                        Console.WriteLine($"Traitement du dossier principal : {folder.Name}");
                        
                        if (modeHorsLigne)
                        {
                            allFolders.Add(folder);
                        }
                        else
                        {
                            GetAllFolders(folder, allFolders);
                        }
                        
                        Thread.Sleep(100);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erreur lors de l'accès au dossier : {ex.Message}");
                    }
                }
                
                if (!modeHorsLigne)
                {
                    try
                    {
                        for (int i = 1; i <= outlookNamespace.Stores.Count; i++)
                        {
                            try
                            {
                                dynamic store = outlookNamespace.Stores[i];
                                if (store.IsDataFileStore)
                                {
                                    Console.WriteLine($"Traitement du magasin de données : {store.DisplayName}");
                                    
                                    try
                                    {
                                        dynamic rootFolder = store.GetRootFolder();
                                        GetAllFolders(rootFolder, allFolders);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Erreur lors de l'accès au magasin de données : {ex.Message}");
                                    }
                                }
                                
                                Thread.Sleep(100);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Erreur lors de l'accès au magasin de données à l'index {i} : {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erreur lors de l'accès aux magasins de données : {ex.Message}");
                    }
                }
                
                int totalEmails = 0;
                int savedEmails = 0;
                int convertedEmails = 0;
                
                foreach (dynamic folder in allFolders)
                {
                    try
                    {
                        string folderName = "Dossier inconnu";
                        try { folderName = folder.Name; } catch { }
                        
                        Console.WriteLine($"Extraction des emails du dossier : {folderName}");
                        
                        string folderPath = Path.Combine(savePath, SanitizeFolderName(folderName));
                        Directory.CreateDirectory(folderPath);
                        
                        dynamic items = null;
                        try
                        {
                            items = folder.Items;
                            totalEmails += items.Count;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Erreur lors de l'accès aux éléments du dossier : {ex.Message}");
                            continue;
                        }
                        
                        foreach (dynamic item in items)
                        {
                            try
                            {
                                string itemClass = "0";
                                try { itemClass = item.Class.ToString(); } catch { }
                                
                                if (itemClass == "43") // olMail
                                {
                                    try
                                    {
                                        string dateStr = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                                        try
                                        {
                                            DateTime receivedTime = item.ReceivedTime;
                                            if (receivedTime != DateTime.MinValue)
                                            {
                                                dateStr = receivedTime.ToString("yyyyMMdd_HHmmss");
                                            }
                                        }
                                        catch { }
                                        
                                        string subject = "Sans_titre";
                                        try { subject = item.Subject; } catch { }
                                        
                                        string fileNameMsg = $"{SanitizeFileName(subject)}_{dateStr}.msg";
                                        string filePathMsg = Path.Combine(folderPath, fileNameMsg);
                                        
                                        item.SaveAs(filePathMsg, 3); // olMSG = 3
                                        savedEmails++;
                                        
                                        try
                                        {
                                            string fileNameEml = $"{SanitizeFileName(subject)}_{dateStr}.eml";
                                            string filePathEml = Path.Combine(folderPath, fileNameEml);
                                            
                                            try
                                            {
                                                item.SaveAs(filePathEml, 10); // olTXT = 0, olRTF = 1, olTemplate = 2, olMSG = 3, olDoc = 4, olHTML = 5, olVCard = 6, olVCal = 7, olICal = 8
                                                convertedEmails++;
                                            }
                                            catch
                                            {
                                                string from = "";
                                                string to = "";
                                                string cc = "";
                                                string subject2 = "";
                                                string body = "";
                                                string htmlBody = "";
                                                DateTime sentTime = DateTime.Now;
                                                
                                                try { from = item.SenderEmailAddress; } catch { }
                                                try { to = item.To; } catch { }
                                                try { cc = item.CC; } catch { }
                                                try { subject2 = item.Subject; } catch { }
                                                try { body = item.Body; } catch { }
                                                try { htmlBody = item.HTMLBody; } catch { }
                                                try { sentTime = item.SentOn; } catch { }
                                                
                                                string emlContent = $@"From: {from}
To: {to}
Cc: {cc}
Subject: {subject2}
Date: {sentTime.ToString("r")}
MIME-Version: 1.0
Content-Type: multipart/alternative;
    boundary=""--boundary_text_alternative""

----boundary_text_alternative
Content-Type: text/plain; charset=""utf-8""

{body}

----boundary_text_alternative
Content-Type: text/html; charset=""utf-8""

{htmlBody}

----boundary_text_alternative--
";
                                                File.WriteAllText(filePathEml, emlContent, Encoding.UTF8);
                                                convertedEmails++;
                                            }
                                            
                                            string fileNameTxt = $"{SanitizeFileName(subject)}_{dateStr}.txt";
                                            string filePathTxt = Path.Combine(folderPath, fileNameTxt);
                                            
                                            try
                                            {
                                                string textContent = "";
                                                try { textContent = item.Body; } catch { }
                                                
                                                if (!string.IsNullOrEmpty(textContent))
                                                {
                                                    string metadata = "";
                                                    try { metadata += $"De: {item.SenderName} <{item.SenderEmailAddress}>\r\n"; } catch { }
                                                    try { metadata += $"À: {item.To}\r\n"; } catch { }
                                                    try { metadata += $"Cc: {item.CC}\r\n"; } catch { }
                                                    try { metadata += $"Objet: {item.Subject}\r\n"; } catch { }
                                                    try { metadata += $"Date: {item.ReceivedTime}\r\n"; } catch { }
                                                    metadata += "\r\n-------------------------------------------------\r\n\r\n";
                                                    
                                                    File.WriteAllText(filePathTxt, metadata + textContent, Encoding.UTF8);
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Erreur lors de la création du fichier texte : {ex.Message}");
                                            }
                                            
                                            try
                                            {
                                                string htmlContent = "";
                                                try { htmlContent = item.HTMLBody; } catch { }
                                                
                                                if (!string.IsNullOrEmpty(htmlContent))
                                                {
                                                    string fileNameHtml = $"{SanitizeFileName(subject)}_{dateStr}.html";
                                                    string filePathHtml = Path.Combine(folderPath, fileNameHtml);
                                                    
                                                    File.WriteAllText(filePathHtml, htmlContent, Encoding.UTF8);
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Erreur lors de la création du fichier HTML : {ex.Message}");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la conversion en format EML : {ex.Message}");
                                        }
                                        
                                        try
                                        {
                                            int attachmentCount = 0;
                                            try { attachmentCount = item.Attachments.Count; } catch { }
                                            
                                            if (attachmentCount > 0)
                                            {
                                                string attachmentFolder = Path.Combine(folderPath, "Attachments_" + Path.GetFileNameWithoutExtension(fileNameMsg));
                                                
                                                try
                                                {
                                                    Directory.CreateDirectory(attachmentFolder);
                                                    
                                                    foreach (dynamic attachment in item.Attachments)
                                                    {
                                                        try
                                                        {
                                                            string attachmentName = "Attachment";
                                                            try { attachmentName = attachment.FileName; } catch { }
                                                            
                                                            string attachmentPath = Path.Combine(attachmentFolder, SanitizeFileName(attachmentName));
                                                            attachment.SaveAsFile(attachmentPath);
                                                            
                                                            Thread.Sleep(50);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            Console.WriteLine($"Erreur lors de la sauvegarde d'une pièce jointe : {ex.Message}");
                                                        }
                                                        finally
                                                        {
                                                            try { Marshal.ReleaseComObject(attachment); } catch { }
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine($"Erreur lors de la création du dossier de pièces jointes : {ex.Message}");
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la sauvegarde des pièces jointes : {ex.Message}");
                                        }
                                        
                                        Thread.Sleep(50);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Erreur lors de la sauvegarde d'un email : {ex.Message}");
                                    }
                                }
                                else if (itemClass == "26") // olAppointment
                                {
                                    try
                                    {
                                        string dateStr = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                                        try
                                        {
                                            DateTime startTime = item.Start;
                                            if (startTime != DateTime.MinValue)
                                            {
                                                dateStr = startTime.ToString("yyyyMMdd_HHmmss");
                                            }
                                        }
                                        catch { }
                                        
                                        string subject = "Sans_titre";
                                        try { subject = item.Subject; } catch { }
                                        
                                        string fileName = $"RDV_{SanitizeFileName(subject)}_{dateStr}.msg";
                                        string filePath = Path.Combine(folderPath, fileName);
                                        
                                        item.SaveAs(filePath, 3); // olMSG = 3
                                        savedEmails++;
                                        
                                        try
                                        {
                                            string fileNameTxt = $"RDV_{SanitizeFileName(subject)}_{dateStr}.txt";
                                            string filePathTxt = Path.Combine(folderPath, fileNameTxt);
                                            
                                            string textContent = "";
                                            textContent += $"Sujet: {subject}\r\n";
                                            try { textContent += $"Début: {item.Start}\r\n"; } catch { }
                                            try { textContent += $"Fin: {item.End}\r\n"; } catch { }
                                            try { textContent += $"Lieu: {item.Location}\r\n"; } catch { }
                                            try { textContent += $"Organisateur: {item.Organizer}\r\n"; } catch { }
                                            try { textContent += $"Participants: {item.RequiredAttendees}\r\n"; } catch { }
                                            try { textContent += $"Participants facultatifs: {item.OptionalAttendees}\r\n"; } catch { }
                                            textContent += "\r\n-------------------------------------------------\r\n\r\n";
                                            try { textContent += item.Body; } catch { }
                                            
                                            File.WriteAllText(filePathTxt, textContent, Encoding.UTF8);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la création du fichier texte pour un rendez-vous : {ex.Message}");
                                        }
                                        
                                        Thread.Sleep(50);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Erreur lors de la sauvegarde d'un rendez-vous : {ex.Message}");
                                    }
                                }
                                else if (itemClass == "40") // olContact
                                {
                                    try
                                    {
                                        string fullName = "Contact";
                                        try { fullName = item.FullName; } catch { }
                                        
                                        string fileName = $"Contact_{SanitizeFileName(fullName)}.msg";
                                        string filePath = Path.Combine(folderPath, fileName);
                                        
                                        item.SaveAs(filePath, 3); // olMSG = 3
                                        savedEmails++;
                                        
                                        try
                                        {
                                            string fileNameVcf = $"Contact_{SanitizeFileName(fullName)}.vcf";
                                            string filePathVcf = Path.Combine(folderPath, fileNameVcf);
                                            
                                            try
                                            {
                                                item.SaveAs(filePathVcf, 6); // olVCard = 6
                                            }
                                            catch
                                            {
                                                string vCardContent = "BEGIN:VCARD\r\nVERSION:3.0\r\n";
                                                try { vCardContent += $"FN:{fullName}\r\n"; } catch { }
                                                try { vCardContent += $"N:{item.LastName};{item.FirstName};;;\r\n"; } catch { }
                                                try { vCardContent += $"EMAIL:{item.Email1Address}\r\n"; } catch { }
                                                try { vCardContent += $"TEL;TYPE=WORK:{item.BusinessTelephoneNumber}\r\n"; } catch { }
                                                try { vCardContent += $"TEL;TYPE=CELL:{item.MobileTelephoneNumber}\r\n"; } catch { }
                                                try { vCardContent += $"ADR;TYPE=WORK:;;{item.BusinessAddress};;;;\r\n"; } catch { }
                                                try { vCardContent += $"ORG:{item.CompanyName}\r\n"; } catch { }
                                                try { vCardContent += $"TITLE:{item.JobTitle}\r\n"; } catch { }
                                                vCardContent += "END:VCARD";
                                                
                                                File.WriteAllText(filePathVcf, vCardContent, Encoding.UTF8);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la création du fichier vCard : {ex.Message}");
                                        }
                                        
                                        try
                                        {
                                            string fileNameTxt = $"Contact_{SanitizeFileName(fullName)}.txt";
                                            string filePathTxt = Path.Combine(folderPath, fileNameTxt);
                                            
                                            string textContent = "";
                                            textContent += $"Nom complet: {fullName}\r\n";
                                            try { textContent += $"E-mail: {item.Email1Address}\r\n"; } catch { }
                                            try { textContent += $"E-mail 2: {item.Email2Address}\r\n"; } catch { }
                                            try { textContent += $"E-mail 3: {item.Email3Address}\r\n"; } catch { }
                                            try { textContent += $"Téléphone bureau: {item.BusinessTelephoneNumber}\r\n"; } catch { }
                                            try { textContent += $"Téléphone mobile: {item.MobileTelephoneNumber}\r\n"; } catch { }
                                            try { textContent += $"Téléphone domicile: {item.HomeTelephoneNumber}\r\n"; } catch { }
                                            try { textContent += $"Entreprise: {item.CompanyName}\r\n"; } catch { }
                                            try { textContent += $"Fonction: {item.JobTitle}\r\n"; } catch { }
                                            try { textContent += $"Adresse professionnelle:\r\n{item.BusinessAddress}\r\n"; } catch { }
                                            try { textContent += $"Adresse personnelle:\r\n{item.HomeAddress}\r\n"; } catch { }
                                            
                                            File.WriteAllText(filePathTxt, textContent, Encoding.UTF8);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la création du fichier texte pour un contact : {ex.Message}");
                                        }
                                        
                                        Thread.Sleep(50);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Erreur lors de la sauvegarde d'un contact : {ex.Message}");
                                    }
                                }
                                else if (itemClass == "48") // olTask
                                {
                                    try
                                    {
                                        string subject = "Sans_titre";
                                        try { subject = item.Subject; } catch { }
                                        
                                        string fileName = $"Tache_{SanitizeFileName(subject)}.msg";
                                        string filePath = Path.Combine(folderPath, fileName);
                                        
                                        item.SaveAs(filePath, 3); // olMSG = 3
                                        savedEmails++;
                                        
                                        try
                                        {
                                            string fileNameTxt = $"Tache_{SanitizeFileName(subject)}.txt";
                                            string filePathTxt = Path.Combine(folderPath, fileNameTxt);
                                            
                                            string textContent = "";
                                            textContent += $"Sujet: {subject}\r\n";
                                            try { textContent += $"État: {item.Status}\r\n"; } catch { }
                                            try { textContent += $"Priorité: {item.Importance}\r\n"; } catch { }
                                            try { textContent += $"Date d'échéance: {item.DueDate}\r\n"; } catch { }
                                            try { textContent += $"Date de début: {item.StartDate}\r\n"; } catch { }
                                            try { textContent += $"Terminé: {item.Complete}\r\n"; } catch { }
                                            try { textContent += $"Propriétaire: {item.Owner}\r\n"; } catch { }
                                            textContent += "\r\n-------------------------------------------------\r\n\r\n";
                                            try { textContent += item.Body; } catch { }
                                            
                                            File.WriteAllText(filePathTxt, textContent, Encoding.UTF8);
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la création du fichier texte pour une tâche : {ex.Message}");
                                        }
                                        
                                        Thread.Sleep(50);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Erreur lors de la sauvegarde d'une tâche : {ex.Message}");
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        string subject = "Sans_titre";
                                        try { subject = item.Subject; } catch { }
                                        
                                        string fileName = $"Element_{SanitizeFileName(subject)}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.msg";
                                        string filePath = Path.Combine(folderPath, fileName);
                                        
                                        item.SaveAs(filePath, 3); // olMSG = 3
                                        savedEmails++;
                                        
                                        try
                                        {
                                            string fileNameTxt = $"Element_{SanitizeFileName(subject)}_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.txt";
                                            string filePathTxt = Path.Combine(folderPath, fileNameTxt);
                                            
                                            string textContent = "";
                                            try { textContent = item.Body; } catch { }
                                            
                                            if (!string.IsNullOrEmpty(textContent))
                                            {
                                                string metadata = "";
                                                try { metadata += $"Sujet: {item.Subject}\r\n"; } catch { }
                                                try { metadata += $"Type: {item.MessageClass}\r\n"; } catch { }
                                                metadata += "\r\n-------------------------------------------------\r\n\r\n";
                                                
                                                File.WriteAllText(filePathTxt, metadata + textContent, Encoding.UTF8);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Erreur lors de la création du fichier texte : {ex.Message}");
                                        }
                                        
                                        Thread.Sleep(50);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Erreur lors de la sauvegarde d'un élément : {ex.Message}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Erreur lors du traitement d'un élément : {ex.Message}");
                            }
                            finally
                            {
                                try { Marshal.ReleaseComObject(item); } catch { }
                            }
                        }
                        
                        try { Marshal.ReleaseComObject(items); } catch { }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erreur lors du traitement du dossier : {ex.Message}");
                    }
                    finally
                    {
                        try { Marshal.ReleaseComObject(folder); } catch { }
                    }
                    
                    Thread.Sleep(100);
                }
                
                string summaryPath = Path.Combine(savePath, "resume_extraction.txt");
                File.WriteAllText(summaryPath, 
                    $"Extraction terminée le {DateTime.Now}\n" +
                    $"Nombre total d'éléments traités : {totalEmails}\n" +
                    $"Nombre d'éléments sauvegardés : {savedEmails}\n" +
                    $"Nombre d'éléments convertis en formats universels : {convertedEmails}\n" +
                    $"Dossiers traités : {allFolders.Count}\n" +
                    $"Mode hors ligne : {modeHorsLigne}");
                
                Console.WriteLine($"Extraction terminée. {savedEmails} éléments sauvegardés sur {totalEmails} éléments trouvés.");
                Console.WriteLine($"Les données ont été sauvegardées dans : {savePath}");
                Console.WriteLine($"{convertedEmails} éléments ont été convertis en formats universels (.eml, .txt, .html, .vcf).");
                
                try { Marshal.ReleaseComObject(outlookNamespace); } catch { }
                try { Marshal.ReleaseComObject(outlookApp); } catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur lors de l'extraction : {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            
            Console.WriteLine("Appuyez sur une touche pour quitter...");
            Console.ReadKey();
        }
        
        static void GetAllFolders(dynamic folder, List<dynamic> allFolders)
        {
            try
            {
                string entryID = "";
                try { entryID = folder.EntryID; } catch { }
                
                if (!string.IsNullOrEmpty(entryID) && !allFolders.Any(f => 
                {
                    try { return f.EntryID == entryID; } 
                    catch { return false; }
                }))
                {
                    allFolders.Add(folder);
                    
                    dynamic subFolders = null;
                    try { subFolders = folder.Folders; } catch { return; }
                    
                    int subFolderCount = 0;
                    try { subFolderCount = subFolders.Count; } catch { return; }
                    
                    if (subFolderCount > 0)
                    {
                        foreach (dynamic subFolder in subFolders)
                        {
                            try
                            {
                                GetAllFolders(subFolder, allFolders);
                                
                                Thread.Sleep(50);
                            }
                            catch (Exception ex)
                            {
                                string subFolderName = "Sous-dossier inconnu";
                                try { subFolderName = subFolder.Name; } catch { }
                                
                                Console.WriteLine($"Erreur lors de l'accès au sous-dossier {subFolderName} : {ex.Message}");
                            }
                            finally
                            {
                                try { Marshal.ReleaseComObject(subFolder); } catch { }
                            }
                        }
                    }
                    
                    try { Marshal.ReleaseComObject(subFolders); } catch { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erreur dans GetAllFolders : {ex.Message}");
            }
        }
        
        static string SanitizeFolderName(string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
                return "Dossier_Sans_Nom";
                
            foreach (char c in Path.GetInvalidPathChars())
            {
                folderName = folderName.Replace(c, '_');
            }
            
            if (folderName.Length > 50)
                folderName = folderName.Substring(0, 50);
                
            return folderName;
        }
        
        static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Sans_titre";
                
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            
            if (fileName.Length > 100)
                fileName = fileName.Substring(0, 100);
                
            return fileName;
        }
    }
}
