using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web;
using MsgReader.Exceptions;
using MsgReader.Helpers;
using MsgReader.Localization;
using MsgReader.Mime.Header;
using MsgReader.Outlook;

/*
  Some extension methods for reading Mail/Message file
*/

namespace MsgReader
{
    public partial class Reader
    {
        public string ExtractToHTMLString(string inputFile, bool hyperlinks = false)
        {
            _errorMessage = string.Empty;

            var extension = CheckFileName(inputFile);

            switch (extension)
            {
                case ".EML":
                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                    {
                        var message = Mime.Message.Load(stream);
                        return WriteEmlEmailToString(message, hyperlinks);
                    }

                case ".MSG":
                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                    using (var message = new Storage.Message(stream))
                    {
                        switch (message.Type)
                        {
                            case Storage.Message.MessageType.Email:
                            case Storage.Message.MessageType.SignedEmail:
                                return WriteMsgEmailToString(message, hyperlinks);

                            //case Storage.Message.MessageType.AppointmentRequest:
                            //case Storage.Message.MessageType.Appointment:
                            //case Storage.Message.MessageType.AppointmentResponse:
                            //case Storage.Message.MessageType.AppointmentResponsePositive:
                            //case Storage.Message.MessageType.AppointmentResponseNegative:
                            //    return WriteMsgAppointment(message, outputFolder, hyperlinks).ToArray();

                            //case Storage.Message.MessageType.Task:
                            //case Storage.Message.MessageType.TaskRequestAccept:
                            //    return WriteMsgTask(message, outputFolder, hyperlinks).ToArray();

                            //case Storage.Message.MessageType.Contact:
                            //    return WriteMsgContact(message, outputFolder, hyperlinks).ToArray();

                            //case Storage.Message.MessageType.StickyNote:
                            //    return WriteMsgStickyNote(message, outputFolder).ToArray();

                            case Storage.Message.MessageType.Unknown:
                                throw new MRFileTypeNotSupported("Unknown message type");
                        }
                    }

                    break;
            }

            return "";
        }

        /// <summary>
        /// Extracts plain bodytext from msg or eml file
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="hyperlinks"></param>
        /// <returns></returns>
        public string GetPlainTextFromEmail(string inputFile)
        {
            _errorMessage = string.Empty;

            var extension = CheckFileName(inputFile);

            switch (extension)
            {
                case ".EML":
                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                    {
                        var message = Mime.Message.Load(stream);
                        return message.TextBody.GetBodyAsText();
                    }

                case ".MSG":
                    using (var stream = File.Open(inputFile, FileMode.Open, FileAccess.Read))
                    using (var message = new Storage.Message(stream))
                    {
                        switch (message.Type)
                        {
                            case Storage.Message.MessageType.Email:
                            case Storage.Message.MessageType.SignedEmail:
                                return message.BodyText;

                            case Storage.Message.MessageType.AppointmentRequest:
                            case Storage.Message.MessageType.Appointment:
                            case Storage.Message.MessageType.AppointmentResponse:
                            case Storage.Message.MessageType.AppointmentResponsePositive:
                            case Storage.Message.MessageType.AppointmentResponseNegative:
                                return message.BodyText;

                            case Storage.Message.MessageType.Task:
                            case Storage.Message.MessageType.TaskRequestAccept:
                                return message.BodyText;

                            case Storage.Message.MessageType.Contact:
                                return message.BodyText;

                            case Storage.Message.MessageType.StickyNote:
                                return message.BodyText;

                            case Storage.Message.MessageType.Unknown:
                                throw new MRFileTypeNotSupported("Unknown message type");
                        }
                    }

                    break;
            }

            return "";
        }

        /// <summary>
        /// Checks if the <paramref name="inputFile"/> and <paramref name="outputFolder"/> is valid
        /// </summary>
        /// <param name="inputFile"></param>
        /// <param name="outputFolder"></param>
        /// <exception cref="ArgumentNullException">Raised when the <paramref name="inputFile"/> or <paramref name="outputFolder"/> is null or empty</exception>
        /// <exception cref="FileNotFoundException">Raised when the <paramref name="inputFile"/> does not exists</exception>
        /// <exception cref="MRFileTypeNotSupported">Raised when the extension is not .msg or .eml</exception>
        private static string CheckFileName(string inputFile)
        {
            var extension = Path.GetExtension(inputFile);
            if (string.IsNullOrEmpty(extension))
                throw new MRFileTypeNotSupported("Expected .msg or .eml extension on the inputfile");

            extension = extension.ToUpperInvariant();

            switch (extension)
            {
                case ".MSG":
                case ".EML":
                    return extension;

                default:
                    throw new MRFileTypeNotSupported("Wrong file extension, expected .msg or .eml");
            }
        }

        #region WriteMsgEmail
        /// <summary>
        /// Writes the body of the MSG E-mail to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message"><see cref="Storage.Message"/></param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private string WriteMsgEmailToString(Storage.Message message, bool hyperlinks)
        {
            var fileName = "email";
            bool htmlBody;
            string body;
            string dummy;
            List<string> attachmentList;
            List<string> files;

            PreProcessMsgFileEx(message,
                hyperlinks,
                ref fileName,
                out htmlBody,
                out body,
                out dummy,
                out attachmentList,
                out files);

            if (!htmlBody)
                hyperlinks = false;

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    LanguageConsts.EmailFollowUpFlag,
                    LanguageConsts.EmailFollowUpLabel,
                    LanguageConsts.EmailFollowUpStatusLabel,
                    LanguageConsts.EmailFollowUpCompletedText,
                    LanguageConsts.TaskStartDateLabel,
                    LanguageConsts.TaskDueDateLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                    #endregion
                };

                if (message.Type == Storage.Message.MessageType.SignedEmail)
                    languageConsts.Add(LanguageConsts.EmailSignedBy);

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var emailHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel,
                message.GetEmailSender(htmlBody, hyperlinks));

            // Sent on
            if (message.SentOn != null)
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                    ((DateTime)message.SentOn).ToString(LanguageConsts.DataFormatWithTime));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                message.GetEmailRecipients(Storage.Recipient.RecipientType.To, htmlBody, hyperlinks));

            // CC
            var cc = message.GetEmailRecipients(Storage.Recipient.RecipientType.Cc, htmlBody, hyperlinks);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = message.GetEmailRecipients(Storage.Recipient.RecipientType.Bcc, htmlBody, hyperlinks);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

            if (message.Type == Storage.Message.MessageType.SignedEmail)
            {
                var signerInfo = message.SignedBy;
                if (message.SignedOn != null)
                    signerInfo += " " + LanguageConsts.EmailSignedByOn + " " +
                                  ((DateTime)message.SignedOn).ToString(LanguageConsts.DataFormatWithTime);

                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSignedBy, signerInfo);
            }

            // Subject
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSubjectLabel, message.Subject);

            // Urgent
            if (!string.IsNullOrEmpty(message.ImportanceText))
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, message.ImportanceText);

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Attachments
            if (attachmentList.Count != 0)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,
                    string.Join(", ", attachmentList));

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // Follow up
            if (message.Flag != null)
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpLabel,
                    message.Flag.Request);

                // When complete
                if (message.Task.Complete != null && (bool)message.Task.Complete)
                {
                    WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFollowUpStatusLabel,
                        LanguageConsts.EmailFollowUpCompletedText);

                    // Task completed date
                    if (message.Task.CompleteTime != null)
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDateCompleted,
                            ((DateTime)message.Task.CompleteTime).ToString(LanguageConsts.DataFormatWithTime));
                }
                else
                {
                    // Task startdate
                    if (message.Task.StartDate != null)
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskStartDateLabel,
                            ((DateTime)message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime));

                    // Task duedate
                    if (message.Task.DueDate != null)
                        WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.TaskDueDateLabel,
                            ((DateTime)message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime));
                }

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Categories
            var categories = message.Categories;
            if (categories != null)
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    String.Join("; ", categories));

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);

            body = InjectHeader(body, emailHeader.ToString());

            // Write the body to a file
            //File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }
        #endregion

        #region WriteEmlEmail
        /// <summary>
        /// Writes the body of the EML E-mail to html or text and extracts all the attachments. The
        /// result is returned as a List of strings
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="outputFolder">The folder where we need to write the output</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and attachments</param>
        /// <returns></returns>
        private string WriteEmlEmailToString(Mime.Message message, bool hyperlinks)
        {
            var fileName = "email";
            bool htmlBody;
            string body;
            List<string> attachmentList;
            List<string> files;

            PreProcessEmlFileEx(message,
                hyperlinks,
                ref fileName,
                out htmlBody,
                out body,
                out attachmentList,
                out files);

            if (!htmlBody)
                hyperlinks = false;

            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var emailHeader = new StringBuilder();

            var headers = message.Headers;

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            var from = string.Empty;
            if (headers.From != null)
                from = message.GetEmailAddresses(new List<RfcMailAddress> { headers.From }, hyperlinks, htmlBody);

            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel, from);

            // Sent on
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel,
                (message.Headers.DateSent.ToLocalTime()).ToString(LanguageConsts.DataFormatWithTime));

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel,
                message.GetEmailAddresses(headers.To, hyperlinks, htmlBody));

            // CC
            var cc = message.GetEmailAddresses(headers.Cc, hyperlinks, htmlBody);
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, cc);

            // BCC
            var bcc = message.GetEmailAddresses(headers.Bcc, hyperlinks, htmlBody);
            if (!string.IsNullOrEmpty(bcc))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, bcc);

            // Subject
            var subject = message.Headers.Subject ?? string.Empty;
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSubjectLabel, subject);

            // Urgent
            var importanceText = string.Empty;
            switch (message.Headers.Importance)
            {
                case MailPriority.Low:
                    importanceText = LanguageConsts.ImportanceLowText;
                    break;

                case MailPriority.Normal:
                    importanceText = LanguageConsts.ImportanceNormalText;
                    break;

                case MailPriority.High:
                    importanceText = LanguageConsts.ImportanceHighText;
                    break;
            }

            if (!string.IsNullOrEmpty(importanceText))
            {
                WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importanceText);

                // Empty line
                WriteHeaderEmptyLine(emailHeader, htmlBody);
            }

            // Attachments
            if (attachmentList.Count != 0)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,
                    string.Join(", ", attachmentList));

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);

            body = InjectHeader(body, emailHeader.ToString());

            // Write the body to a file
           // File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }
        #endregion

        public static string MergeMailHeaderAndRtfBody(Dictionary<string,string> mailMeta, string rtfText,bool htmlBody)
        {
            RtfToHtmlConverter cnv = new RtfToHtmlConverter();
            var body = cnv.ConvertRtfToHtml(rtfText);
            return MergeMailHeaderAndBody(mailMeta,body, htmlBody);
        }
        public static string MergeMailHeaderAndBody(Dictionary<string,string> mailMeta, string body,bool htmlBody)
        {
            var maxLength = 0;
            var emailHeader = new StringBuilder();

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.EmailFromLabel,
                    LanguageConsts.EmailSentOnLabel,
                    LanguageConsts.EmailToLabel,
                    LanguageConsts.EmailCcLabel,
                    LanguageConsts.EmailBccLabel,
                    LanguageConsts.EmailSubjectLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.EmailAttachmentsLabel,
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            // Start of table
            WriteHeaderStart(emailHeader, htmlBody);

            // From
            var from = mailMeta["from"]; 

            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailFromLabel, from);

            // Sent on
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSentOnLabel, mailMeta["DateSent"]);
                

            // To
            WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailToLabel, mailMeta["To"]);
              
            // CC
            if (!string.IsNullOrEmpty(mailMeta["CC"]))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailCcLabel, mailMeta["CC"]);

            // BCC
            if (!string.IsNullOrEmpty(mailMeta["BCC"]))
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailBccLabel, mailMeta["BCC"]);

            // Subject
            var subject = mailMeta["Subject"];
            WriteHeaderLine(emailHeader, htmlBody, maxLength, LanguageConsts.EmailSubjectLabel, subject);

            
            // Attachments
            if (mailMeta["attachment"]!=null)
                WriteHeaderLineNoEncoding(emailHeader, htmlBody, maxLength, LanguageConsts.EmailAttachmentsLabel,mailMeta["attachment"]);

            // Empty line
            WriteHeaderEmptyLine(emailHeader, htmlBody);

            // End of table + empty line
            WriteHeaderEnd(emailHeader, htmlBody);

            body = InjectHeader(body, emailHeader.ToString());

            // Write the body to a file
            // File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }

        public static string MergeAppointmentHeaderAndBody(Dictionary<string,string> appMeta,string rtfText)
        {

            RtfToHtmlConverter cnv = new RtfToHtmlConverter();
            var body=cnv.ConvertRtfToHtml(rtfText);

            bool htmlBody=true;
            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.AppointmentSubjectLabel,
                    LanguageConsts.AppointmentLocationLabel,
                    LanguageConsts.AppointmentStartDateLabel,
                    LanguageConsts.AppointmentEndDateLabel,
                    LanguageConsts.AppointmentRecurrenceTypeLabel,
                    LanguageConsts.AppointmentClientIntentLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentRecurrencePaternLabel,
                    LanguageConsts.AppointmentOrganizerLabel,
                    LanguageConsts.AppointmentMandatoryParticipantsLabel,
                    LanguageConsts.AppointmentOptionalParticipantsLabel,
                    LanguageConsts.AppointmentCategoriesLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskDateCompleted,
                    LanguageConsts.EmailCategoriesLabel
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var appointmentHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(appointmentHeader, htmlBody);

            // Subject
            WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentSubjectLabel, appMeta["Subject"]);

            // Location
            WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentLocationLabel,appMeta["Location"]);

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Start
            if (appMeta["Start"] != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentStartDateLabel,appMeta["Start"]);

            // End
            if (appMeta["End"] != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentEndDateLabel, appMeta["End"]);

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Recurrence type
            if (appMeta["RecurrenceTypeText"] != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentRecurrenceTypeLabel,
                    appMeta["RecurrenceTypeText"]);

            // Recurrence patern
            if (appMeta["RecurrencePatern"]!=null)
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentRecurrencePaternLabel,
                    appMeta["RecurrencePatern"]);

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Status
            if (appMeta["ClientIntentText"] != null)
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentClientIntentLabel,
                    appMeta["ClientIntentText"]);

            // Appointment organizer (FROM)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentOrganizerLabel,
                appMeta["EmailSender"]);

            // Mandatory participants (TO)
            WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                LanguageConsts.AppointmentMandatoryParticipantsLabel,
                appMeta["EmailRecipients"]);

            // Optional participants (CC)
            var cc = appMeta["EmailRecipientsCC"];
            if (!string.IsNullOrEmpty(cc))
                WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength,
                    LanguageConsts.AppointmentOptionalParticipantsLabel, cc);

            // Empty line
            WriteHeaderEmptyLine(appointmentHeader, htmlBody);

            // Categories
            var categories = appMeta["Categories"];
            if (categories != null)
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    String.Join("; ", categories));

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Urgent
            var importance = appMeta["ImportanceText"];
            if (!string.IsNullOrEmpty(importance))
            {
                WriteHeaderLine(appointmentHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importance);

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // Attachments
            if (appMeta["Attachments"] != null)
            {
                WriteHeaderLineNoEncoding(appointmentHeader, htmlBody, maxLength, LanguageConsts.AppointmentAttachmentsLabel,
                    appMeta["Attachments"]);

                // Empty line
                WriteHeaderEmptyLine(appointmentHeader, htmlBody);
            }

            // End of table + empty line
            WriteHeaderEnd(appointmentHeader, htmlBody);

            body = InjectHeader(body, appointmentHeader.ToString());

            // Write the body to a file
           // File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }
        public static string MergeGinContactAndBody(Dictionary<string, string> cntcMeta, string txt)
        {

            var body = txt;
            var contactHeader = new StringBuilder();
            bool htmlBody = true;
            var maxLength = 0;
            // Start of table
            WriteHeaderStart(contactHeader, htmlBody);

            foreach(var d in cntcMeta)
            {
                if(d.Key!=null && d.Value!=null)
                    WriteHeaderLine(contactHeader, htmlBody, maxLength, d.Key, d.Value);
            }
           

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            WriteHeaderEnd(contactHeader, htmlBody);

            body = InjectHeader(body, contactHeader.ToString());

            // Write the body to a file
            // File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }

        public static string MergeContactAndBody(Dictionary<string,string> cntcMeta, string rtfText)
        {
            RtfToHtmlConverter cnv = new RtfToHtmlConverter();
            var body = cnv.ConvertRtfToHtml(rtfText);

            bool htmlBody = true;
            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                #region Language consts
                var languageConsts = new List<string>
                {
                    LanguageConsts.DisplayNameLabel,
                    LanguageConsts.SurNameLabel,
                    LanguageConsts.GivenNameLabel,
                    LanguageConsts.FunctionLabel,
                    LanguageConsts.DepartmentLabel,
                    LanguageConsts.CompanyLabel,
                    LanguageConsts.WorkAddressLabel,
                    LanguageConsts.BusinessTelephoneNumberLabel,
                    LanguageConsts.BusinessTelephoneNumber2Label,
                    LanguageConsts.BusinessFaxNumberLabel,
                    LanguageConsts.HomeAddressLabel,
                    LanguageConsts.HomeTelephoneNumberLabel,
                    LanguageConsts.HomeTelephoneNumber2Label,
                    LanguageConsts.HomeFaxNumberLabel,
                    LanguageConsts.OtherAddressLabel,
                    LanguageConsts.OtherFaxLabel,
                    LanguageConsts.PrimaryTelephoneNumberLabel,
                    LanguageConsts.PrimaryFaxNumberLabel,
                    LanguageConsts.AssistantTelephoneNumberLabel,
                    LanguageConsts.InstantMessagingAddressLabel,
                    LanguageConsts.CompanyMainTelephoneNumberLabel,
                    LanguageConsts.CellularTelephoneNumberLabel,
                    LanguageConsts.CarTelephoneNumberLabel,
                    LanguageConsts.RadioTelephoneNumberLabel,
                    LanguageConsts.BeeperTelephoneNumberLabel,
                    LanguageConsts.CallbackTelephoneNumberLabel,
                    LanguageConsts.TextTelephoneLabel,
                    LanguageConsts.ISDNNumberLabel,
                    LanguageConsts.TelexNumberLabel,
                    LanguageConsts.Email1EmailAddressLabel,
                    LanguageConsts.Email1DisplayNameLabel,
                    LanguageConsts.Email2EmailAddressLabel,
                    LanguageConsts.Email2DisplayNameLabel,
                    LanguageConsts.Email3EmailAddressLabel,
                    LanguageConsts.Email3DisplayNameLabel,
                    LanguageConsts.BirthdayLabel,
                    LanguageConsts.WeddingAnniversaryLabel,
                    LanguageConsts.SpouseNameLabel,
                    LanguageConsts.ProfessionLabel,
                    LanguageConsts.HtmlLabel
                };
                #endregion

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var contactHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(contactHeader, htmlBody);


            // Full name
            if (!string.IsNullOrEmpty(cntcMeta["DisplayName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.DisplayNameLabel,
                    cntcMeta["DisplayName"]);

            // Last name
            if (!string.IsNullOrEmpty(cntcMeta["SurName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.SurNameLabel, cntcMeta["SurName"]);

            // First name
            if (!string.IsNullOrEmpty(cntcMeta["GivenName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.GivenNameLabel, cntcMeta["GivenName"]);

            // Job title
            if (!string.IsNullOrEmpty(cntcMeta["Function"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.FunctionLabel, cntcMeta["Function"]);

            // Department
            if (!string.IsNullOrEmpty(cntcMeta["Department"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.DepartmentLabel,
                    cntcMeta["Department"]);

            // Company
            if (!string.IsNullOrEmpty(cntcMeta["Company"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CompanyLabel, cntcMeta["Company"]);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Business address
            if (!string.IsNullOrEmpty(cntcMeta["WorkAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.WorkAddressLabel,
                    cntcMeta["WorkAddress"]);

            // Home address
            if (!string.IsNullOrEmpty(cntcMeta["HomeAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeAddressLabel,
                    cntcMeta["HomeAddress"]);

            // Other address
            if (!string.IsNullOrEmpty(cntcMeta["OtherAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.OtherAddressLabel,
                    cntcMeta["OtherAddress"]);

            // Instant messaging
            if (!string.IsNullOrEmpty(cntcMeta["InstantMessagingAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.InstantMessagingAddressLabel,
                    cntcMeta["InstantMessagingAddress"]);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Business telephone number
            if (!string.IsNullOrEmpty(cntcMeta["BusinessTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BusinessTelephoneNumberLabel,
                    cntcMeta["BusinessTelephoneNumber"]);

            // Business telephone number 2
            if (!string.IsNullOrEmpty(cntcMeta["BusinessTelephoneNumber2"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BusinessTelephoneNumber2Label,
                    cntcMeta["BusinessTelephoneNumber2"]);

            // Assistant's telephone number
            if (!string.IsNullOrEmpty(cntcMeta["AssistantTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.AssistantTelephoneNumberLabel,
                    cntcMeta["AssistantTelephoneNumber"]);

            // Company main phone
            if (!string.IsNullOrEmpty(cntcMeta["CompanyMainTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CompanyMainTelephoneNumberLabel,
                    cntcMeta["CompanyMainTelephoneNumber"]);

            // Home telephone number
            if (!string.IsNullOrEmpty(cntcMeta["HomeTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeTelephoneNumberLabel,
                    cntcMeta["HomeTelephoneNumber"]);

            // Home telephone number 2
            if (!string.IsNullOrEmpty(cntcMeta["HomeTelephoneNumber2"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeTelephoneNumber2Label,
                   cntcMeta["HomeTelephoneNumber2"]);

            // Mobile phone
            if (!string.IsNullOrEmpty(cntcMeta["CellularTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CellularTelephoneNumberLabel,
                    cntcMeta["CellularTelephoneNumber"]);

            // Car phone
            if (!string.IsNullOrEmpty(cntcMeta["CarTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CarTelephoneNumberLabel,
                    cntcMeta["CarTelephoneNumber"]);

            // Radio
            if (!string.IsNullOrEmpty(cntcMeta["RadioTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.RadioTelephoneNumberLabel,
                    cntcMeta["RadioTelephoneNumber"]);

            // Beeper
            if (!string.IsNullOrEmpty(cntcMeta["BeeperTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BeeperTelephoneNumberLabel,
                    cntcMeta["BeeperTelephoneNumber"]);

            // Callback
            if (!string.IsNullOrEmpty(cntcMeta["CallbackTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.CallbackTelephoneNumberLabel,
                    cntcMeta["CallbackTelephoneNumber"]);

            // Other
            if (!string.IsNullOrEmpty(cntcMeta["OtherTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.OtherTelephoneNumberLabel,
                    cntcMeta["OtherTelephoneNumber"]);

            // Primary telephone number
            if (!string.IsNullOrEmpty(cntcMeta["PrimaryTelephoneNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.PrimaryTelephoneNumberLabel,
                    cntcMeta["PrimaryTelephoneNumber"]);

            // Telex
            if (!string.IsNullOrEmpty(cntcMeta["TelexNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.TelexNumberLabel,
                    cntcMeta["TelexNumber"]);

            // TTY/TDD phone
            if (!string.IsNullOrEmpty(cntcMeta["TextTelephone"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.TextTelephoneLabel,
                    cntcMeta["TextTelephone"]);

            // ISDN
            if (!string.IsNullOrEmpty(cntcMeta["ISDNNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.ISDNNumberLabel,
                   cntcMeta["ISDNNumber"]);

            // Other fax (primary fax, weird that they call it like this in Outlook)
            if (!string.IsNullOrEmpty(cntcMeta["PrimaryFaxNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.PrimaryFaxNumberLabel,
                    cntcMeta["PrimaryFaxNumber"]);

            // Business fax
            if (!string.IsNullOrEmpty(cntcMeta["BusinessFaxNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BusinessFaxNumberLabel,
                    cntcMeta["BusinessFaxNumber"]);

            // Home fax
            if (!string.IsNullOrEmpty(cntcMeta["HomeFaxNumber"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HomeFaxNumberLabel,
                     cntcMeta["HomeFaxNumber"]);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // E-mail
            if (!string.IsNullOrEmpty(cntcMeta["Email1EmailAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email1EmailAddressLabel,
                     cntcMeta["Email1EmailAddress"]);

            // E-mail display as
            if (!string.IsNullOrEmpty(cntcMeta["Email1DisplayName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email1DisplayNameLabel,
                    cntcMeta["Email1DisplayName"]);

            // E-mail 2
            if (!string.IsNullOrEmpty(cntcMeta["Email2EmailAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email2EmailAddressLabel,
                    cntcMeta["Email2EmailAddress"]);

            // E-mail display as 2
            if (!string.IsNullOrEmpty(cntcMeta["Email2DisplayName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email2DisplayNameLabel,
                    cntcMeta["Email2DisplayName"]);

            // E-mail 3
            if (!string.IsNullOrEmpty(cntcMeta["Email3EmailAddress"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email3EmailAddressLabel,
                    cntcMeta["Email3EmailAddress"]);

            // E-mail display as 3
            if (!string.IsNullOrEmpty(cntcMeta["Email3DisplayName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.Email3DisplayNameLabel,
                    cntcMeta["Email3DisplayName"]);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            // Birthday
            if (cntcMeta["Birthday"] != null)
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.BirthdayLabel,
                    cntcMeta["Birthday"]);

            // Anniversary
            if (cntcMeta["WeddingAnniversary"] != null)
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.WeddingAnniversaryLabel,
                    cntcMeta["WeddingAnniversary"]);

            // Spouse/Partner
            if (!string.IsNullOrEmpty(cntcMeta["SpouseName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.SpouseNameLabel,
                    cntcMeta["SpouseName"]);

            // Profession
            if (!string.IsNullOrEmpty(cntcMeta["Profession"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.ProfessionLabel,
                    cntcMeta["Profession"]);

            // Assistant
            if (!string.IsNullOrEmpty(cntcMeta["AssistantName"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.AssistantTelephoneNumberLabel,
                    cntcMeta["AssistantName"]);

            // Web page
            if (!string.IsNullOrEmpty(cntcMeta["Html"]))
                WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.HtmlLabel, cntcMeta["Html"]);

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            //// Categories
            //var categories = message.Categories;
            //if (categories != null)
            //    WriteHeaderLine(contactHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
            //        String.Join("; ", categories));

            // Empty line
            WriteHeaderEmptyLine(contactHeader, htmlBody);

            WriteHeaderEnd(contactHeader, htmlBody);

            body = InjectHeader(body, contactHeader.ToString());

            // Write the body to a file
           // File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }

        public static string MergeTaskAndBody(Storage.Message message, string rtfText)
        {
            RtfToHtmlConverter cnv = new RtfToHtmlConverter();
            var body = cnv.ConvertRtfToHtml(rtfText);

            bool htmlBody = true;
            var maxLength = 0;

            // Calculate padding width when we are going to write a text file
            if (!htmlBody)
            {
                var languageConsts = new List<string>
                {
                    #region LanguageConsts
                    LanguageConsts.TaskSubjectLabel,
                    LanguageConsts.TaskStartDateLabel,
                    LanguageConsts.TaskDueDateLabel,
                    LanguageConsts.ImportanceLabel,
                    LanguageConsts.TaskStatusLabel,
                    LanguageConsts.TaskPercentageCompleteLabel,
                    LanguageConsts.TaskEstimatedEffortLabel,
                    LanguageConsts.TaskActualEffortLabel,
                    LanguageConsts.TaskOwnerLabel,
                    LanguageConsts.TaskContactsLabel,
                    LanguageConsts.EmailCategoriesLabel,
                    LanguageConsts.TaskCompanyLabel,
                    LanguageConsts.TaskBillingInformationLabel,
                    LanguageConsts.TaskMileageLabel
                    #endregion
                };

                maxLength = languageConsts.Select(languageConst => languageConst.Length).Concat(new[] { 0 }).Max() + 2;
            }

            var taskHeader = new StringBuilder();

            // Start of table
            WriteHeaderStart(taskHeader, htmlBody);

            // Subject
            WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskSubjectLabel, message.Subject);

            // Task startdate
            if (message.Task.StartDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskStartDateLabel,
                    ((DateTime)message.Task.StartDate).ToString(LanguageConsts.DataFormatWithTime));

            // Task duedate
            if (message.Task.DueDate != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength,
                    LanguageConsts.TaskDueDateLabel,
                    ((DateTime)message.Task.DueDate).ToString(LanguageConsts.DataFormatWithTime));

            // Urgent
            var importance = message.ImportanceText;
            if (!string.IsNullOrEmpty(importance))
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.ImportanceLabel, importance);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // Status
            if (message.Task.StatusText != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskStatusLabel, message.Task.StatusText);

            // Percentage complete
            if (message.Task.PercentageComplete != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskPercentageCompleteLabel,
                    (message.Task.PercentageComplete * 100) + "%");

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // Estimated effort
            if (message.Task.EstimatedEffortText != null)
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskEstimatedEffortLabel,
                    message.Task.EstimatedEffortText);

                // Actual effort
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskActualEffortLabel,
                    message.Task.ActualEffortText);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Owner
            if (message.Task.Owner != null)
            {
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskOwnerLabel, message.Task.Owner);

                // Empty line
                WriteHeaderEmptyLine(taskHeader, htmlBody);
            }

            // Contacts
            if (message.Task.Contacts != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskContactsLabel,
                    string.Join("; ", message.Task.Contacts.ToArray()));

            // Categories
            var categories = message.Categories;
            if (categories != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.EmailCategoriesLabel,
                    String.Join("; ", categories));

            // Companies
            if (message.Task.Companies != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskCompanyLabel,
                    string.Join("; ", message.Task.Companies.ToArray()));

            // Billing information
            if (message.Task.BillingInformation != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskBillingInformationLabel,
                    message.Task.BillingInformation);

            // Mileage
            if (message.Task.Mileage != null)
                WriteHeaderLine(taskHeader, htmlBody, maxLength, LanguageConsts.TaskMileageLabel, message.Task.Mileage);

            // Attachments
            //if (attachmentList.Count != 0)
            //{
            //    WriteHeaderLineNoEncoding(taskHeader, htmlBody, maxLength, LanguageConsts.AppointmentAttachmentsLabel,
            //        string.Join(", ", attachmentList));

            //    // Empty line
            //    WriteHeaderEmptyLine(taskHeader, htmlBody);
            //}

            // Empty line
            WriteHeaderEmptyLine(taskHeader, htmlBody);

            // End of table
            WriteHeaderEnd(taskHeader, htmlBody);

            body = InjectHeader(body, taskHeader.ToString());

            // Write the body to a file
           // File.WriteAllText(fileName, body, Encoding.UTF8);

            return body;
        }

        #region PreProcessEmlFile
        /// <summary>
        /// This function pre processes the EML <see cref="Mime.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Mime.MessagePart">attachment</see> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Mime.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and 
        /// attachments (when there is an html body)</param>
        /// <param name="outputFolder">The outputfolder where alle extracted files need to be written</param>
        /// <param name="fileName">Returns the filename for the html or text body</param>
        /// <param name="htmlBody">Returns true when the <see cref="Mime.Message"/> object did contain 
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after pre processing the <see cref="Mime.Message"/> object</param>
        private void PreProcessEmlFileEx(Mime.Message message,
            bool hyperlinks,
            ref string fileName,
            out bool htmlBody,
            out string body,
            out List<string> attachments,
            out List<string> files)
        {
            attachments = new List<string>();
            files = new List<string>();

            var bodyMessagePart = message.HtmlBody;

            if (bodyMessagePart != null)
            {
                body = bodyMessagePart.GetBodyAsText();
                htmlBody = true;
            }
            else
            {
                bodyMessagePart = message.TextBody;

                // When there is no body at all we just make an empty html document
                if (bodyMessagePart != null)
                {
                    body = bodyMessagePart.GetBodyAsText();
                    htmlBody = false;
                }
                else
                {
                    htmlBody = true;
                    body = "<html><head></head><body></body></html>";
                }
            }
        }
        #endregion

        #region PreProcessMsgFile
        /// <summary>
        /// This function pre processes the Outlook MSG <see cref="Storage.Message"/> object, it tries to find the html (or text) body
        /// and reads all the available <see cref="Storage.Attachment"/> objects. When an attachment is inline it tries to
        /// map this attachment to the html body part when this is available
        /// </summary>
        /// <param name="message">The <see cref="Storage.Message"/> object</param>
        /// <param name="hyperlinks">When true then hyperlinks are generated for the To, CC, BCC and 
        /// attachments (when there is an html body)</param>
        /// <param name="outputFolder">The outputfolder where alle extracted files need to be written</param>
        /// <param name="fileName">Returns the filename for the html or text body</param>
        /// <param name="htmlBody">Returns true when the <see cref="Storage.Message"/> object did contain 
        /// an HTML body</param>
        /// <param name="body">Returns the html or text body</param>
        /// <param name="contactPhotoFileName">Returns the filename of the contact photo. This field will only
        /// return a value when the <see cref="Storage.Message"/> object is a <see cref="Storage.Message.MessageType.Contact"/> 
        /// type and the <see cref="Storage.Message.Attachments"/> contains an object that has the 
        /// <param ref="Storage.Message.Attachment.IsContactPhoto"/> set to true, otherwise this field will always be null</param>
        /// <param name="attachments">Returns a list of names with the found attachment</param>
        /// <param name="files">Returns all the files that are generated after pre processing the <see cref="Storage.Message"/> object</param>
        private void PreProcessMsgFileEx(Storage.Message message,
            bool hyperlinks,
            ref string fileName,
            out bool htmlBody,
            out string body,
            out string contactPhotoFileName,
            out List<string> attachments,
            out List<string> files)
        {
            const string rtfInlineObject = "[*[RTFINLINEOBJECT]*]";

            htmlBody = true;
            attachments = new List<string>();
            files = new List<string>();
            contactPhotoFileName = null;
            body = message.BodyHtml;

            if (string.IsNullOrEmpty(body))
            {
                htmlBody = false;
                body = message.BodyRtf;

                // If the body is not null then we convert it to HTML
                if (body != null)
                {
                    // The RtfToHtmlConverter doesn't support the RTF \objattph tag. So we need to 
                    // replace the tag with some text that does survive the conversion. Later on we 
                    // will replace these tags with the correct inline image tags
                    body = body.Replace("\\objattph", rtfInlineObject);
                    var converter = new RtfToHtmlConverter();
                    body = converter.ConvertRtfToHtml(body);
                    htmlBody = true;
                }
                else
                {
                    body = message.BodyText;

                    // When there is no body at all we just make an empty html document
                    if (body == null)
                    {
                        htmlBody = true;
                        body = "<html><head></head><body></body></html>";
                    }
                }
            }
        }
        #endregion
    }
}