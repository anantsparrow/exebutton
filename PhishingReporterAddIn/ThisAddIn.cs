using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PhishingReporterAddIn
{
    [global::Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute()]
    public partial class ThisAddIn
    {
        public static ThisAddIn Instance { get; private set; }
        private System.Threading.SynchronizationContext _uiSyncContext;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;
            _uiSyncContext = System.Threading.SynchronizationContext.Current;
            // Pre-load or initialize configuration on startup to ensure folder structure exists
            ConfigManager.Load();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Shutdown logic if needed
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new PhishingRibbon();
        }

        public void ReportPhishingEmail()
        {
            Outlook.MailItem mailItem = null;
            try
            {
                // Obtain current active window (can be Explorer or Inspector reading pane)
                object activeWindow = this.Application.ActiveWindow();
                if (activeWindow is Outlook.Inspector inspector)
                {
                    mailItem = inspector.CurrentItem as Outlook.MailItem;
                }
                else if (activeWindow is Outlook.Explorer explorer)
                {
                    if (explorer.Selection != null && explorer.Selection.Count > 0)
                    {
                        mailItem = explorer.Selection[1] as Outlook.MailItem;
                    }
                }

                if (mailItem == null)
                {
                    var warningForm = new NotificationForm(false, "No Email Selected", "Please select or open an email before reporting phishing.");
                    warningForm.ShowDialog();
                    return;
                }

                // 1. Extract and copy COM properties on UI thread to avoid wrong-thread COM access exceptions
                string subject = mailItem.Subject ?? "";
                string senderEmail = mailItem.SenderEmailAddress ?? "";
                
                string fromAddress = senderEmail;
                try
                {
                    if (mailItem.Sender != null)
                    {
                        fromAddress = mailItem.Sender.Address ?? senderEmail;
                    }
                }
                catch {}

                var toList = new List<string>();
                if (mailItem.Recipients != null)
                {
                    foreach (Outlook.Recipient recipient in mailItem.Recipients)
                    {
                        if (recipient.Type == (int)Outlook.OlMailRecipientType.olTo)
                        {
                            string email = recipient.Address;
                            if (string.IsNullOrEmpty(email) || !email.Contains("@"))
                            {
                                try
                                {
                                    var exchangeUser = recipient.AddressEntry.GetExchangeUser();
                                    if (exchangeUser != null)
                                    {
                                        email = exchangeUser.PrimarySmtpAddress;
                                    }
                                }
                                catch {}
                            }
                            if (string.IsNullOrEmpty(email))
                            {
                                email = recipient.Name;
                            }
                            toList.Add(email);
                        }
                    }
                }

                string dateStr = "";
                try
                {
                    dateStr = mailItem.ReceivedTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                }
                catch
                {
                    dateStr = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                }

                string headers = "";
                try
                {
                    headers = mailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F") as string;
                }
                catch {}
                if (string.IsNullOrEmpty(headers))
                {
                    headers = $"From: {senderEmail}\r\nTo: {string.Join(", ", toList)}\r\nSubject: {subject}\r\nDate: {dateStr}";
                }

                string body = mailItem.HTMLBody;
                if (string.IsNullOrEmpty(body))
                {
                    body = mailItem.Body ?? "";
                }

                // Create unique temp directory to extract attachment files and save the MSG file
                string tempDir = Path.Combine(Path.GetTempPath(), "PhishingReportTemp_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(tempDir);

                var attachmentPaths = new List<string>();
                if (mailItem.Attachments != null && mailItem.Attachments.Count > 0)
                {
                    int attIndex = 1;
                    foreach (Outlook.Attachment attachment in mailItem.Attachments)
                    {
                        string originalName = attachment.FileName;
                        if (string.IsNullOrEmpty(originalName))
                        {
                            originalName = "attachment_" + attIndex;
                        }
                        
                        string safeFileName = MakeValidFileName(originalName);
                        if (string.IsNullOrEmpty(safeFileName))
                        {
                            safeFileName = "attachment_" + attIndex;
                        }
                        
                        string fullPath = Path.Combine(tempDir, safeFileName);
                        
                        // Handle potential name collisions (e.g., duplicate attachment names)
                        int collisionCounter = 1;
                        while (File.Exists(fullPath))
                        {
                            string filenameNoExt = Path.GetFileNameWithoutExtension(safeFileName);
                            string ext = Path.GetExtension(safeFileName);
                            fullPath = Path.Combine(tempDir, $"{filenameNoExt}_{collisionCounter}{ext}");
                            collisionCounter++;
                        }
                        
                        attachment.SaveAsFile(fullPath);
                        attachmentPaths.Add(fullPath);
                        attIndex++;
                    }
                }

                // Save email as MSG file
                string safeSubject = MakeValidFileName(subject);
                if (string.IsNullOrEmpty(safeSubject))
                {
                    safeSubject = "ReportedEmail";
                }
                string msgPath = Path.Combine(tempDir, safeSubject + ".msg");
                mailItem.SaveAs(msgPath, Outlook.OlSaveAsType.olMSG);

                // Load REST API endpoint details
                var config = ConfigManager.Load();
                if (string.IsNullOrEmpty(config.ApiUrl))
                {
                    var errForm = new NotificationForm(false, "Configuration Error", "The REST API endpoint URL is not configured. Please check config.json.");
                    errForm.ShowDialog();
                    return;
                }

                // 2. Perform HTTP file transfer inside a worker pool thread
                Task.Run(async () =>
                {
                    bool success = false;
                    string errorMessage = "";

                    try
                    {
                        success = await UploadReportAsync(config.ApiUrl, subject, senderEmail, fromAddress, toList, dateStr, headers, body, msgPath, attachmentPaths, config.TimeoutSeconds);
                    }
                    catch (Exception ex)
                    {
                        success = false;
                        errorMessage = ex.Message;
                    }
                    finally
                    {
                        // Clean up temporary disk files in all cases
                        try
                        {
                            if (Directory.Exists(tempDir))
                            {
                                Directory.Delete(tempDir, true);
                            }
                        }
                        catch {}
                    }

                    // 3. Dispatch completion UI back onto the main STA thread safely
                    if (_uiSyncContext != null)
                    {
                        _uiSyncContext.Post(_ =>
                        {
                            ShowResultNotification(success, errorMessage);
                        }, null);
                    }
                    else
                    {
                        // Fallback UI in case SynchronizationContext is unassigned
                        ShowResultNotification(success, errorMessage);
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to process the email: " + ex.Message, "Phishing Triage Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowResultNotification(bool success, string errorMessage)
        {
            if (success)
            {
                var successForm = new NotificationForm(true, "Report Submitted", "Thank you! The phishing report was submitted successfully for analysis.");
                successForm.ShowDialog();
            }
            else
            {
                string detail = "We could not connect to the security triage service. Please check your connection.";
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    detail = $"Upload failed: {errorMessage}";
                }
                var failureForm = new NotificationForm(false, "Submission Failed", detail);
                failureForm.ShowDialog();
            }
        }

        private async Task<bool> UploadReportAsync(string apiUrl, string subject, string sender, string fromAddress, List<string> toList, string dateStr, string headers, string body, string msgPath, List<string> attachmentPaths, int timeoutSeconds)
        {
            using (var client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(timeoutSeconds);
                using (var content = new MultipartFormDataContent())
                {
                    // Add text fields
                    content.Add(new StringContent(subject), "subject");
                    content.Add(new StringContent(sender), "sender");
                    content.Add(new StringContent(fromAddress), "from");
                    
                    foreach (var to in toList)
                    {
                        content.Add(new StringContent(to), "to[]");
                    }
                    content.Add(new StringContent(dateStr), "date");
                    content.Add(new StringContent(headers), "headers");
                    content.Add(new StringContent(body), "body");

                    var openStreams = new List<Stream>();
                    try
                    {
                        // Add MSG file stream
                        if (File.Exists(msgPath))
                        {
                            var msgStream = File.OpenRead(msgPath);
                            openStreams.Add(msgStream);
                            var fileContent = new StreamContent(msgStream);
                            fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                            content.Add(fileContent, "email_msg", Path.GetFileName(msgPath));
                        }

                        // Add attachment streams
                        foreach (var attPath in attachmentPaths)
                        {
                            if (File.Exists(attPath))
                            {
                                var attStream = File.OpenRead(attPath);
                                openStreams.Add(attStream);
                                var attContent = new StreamContent(attStream);
                                attContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");
                                content.Add(attContent, "attachments[]", Path.GetFileName(attPath));
                            }
                        }

                        var response = await client.PostAsync(apiUrl, content);
                        return response.IsSuccessStatusCode;
                    }
                    finally
                    {
                        // Clean up file handle locks
                        foreach (var stream in openStreams)
                        {
                            try { stream.Dispose(); } catch {}
                        }
                    }
                }
            }
        }

        private string MakeValidFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }
            // Truncate to safe length
            if (name.Length > 100)
            {
                string ext = Path.GetExtension(name);
                name = name.Substring(0, 90) + ext;
            }
            return name;
        }
    }
}
