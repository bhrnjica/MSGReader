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
            var htmlConvertedFromRtf = false;
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
                    htmlConvertedFromRtf = true;
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

            //fileName = outputFolder +
            //           (!string.IsNullOrEmpty(message.Subject)
            //               ? FileManager.RemoveInvalidFileNameChars(message.Subject)
            //               : fileName) + (htmlBody ? ".htm" : ".txt");

            //files.Add(fileName);

            //var inlineAttachments = new SortedDictionary<int, string>();

            //foreach (var attachment in message.Attachments)
            //{
            //    FileInfo fileInfo = null;
            //    var attachmentFileName = string.Empty;
            //    var renderingPosition = -1;
            //    var isInline = false;

            //    // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
            //    if (attachment is Storage.Attachment)
            //    {
            //        var attach = (Storage.Attachment)attachment;
            //        attachmentFileName = attach.FileName;
            //        renderingPosition = attach.RenderingPosition;
            //        fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attachmentFileName));
            //        File.WriteAllBytes(fileInfo.FullName, attach.Data);
            //        isInline = attach.IsInline;

            //        if (attach.IsContactPhoto && htmlBody)
            //        {
            //            contactPhotoFileName = fileInfo.FullName;
            //            continue;
            //        }

            //        if (!htmlConvertedFromRtf)
            //        {
            //            // When we find an inline attachment we have to replace the CID tag inside the html body
            //            // with the name of the inline attachment. But before we do this we check if the CID exists.
            //            // When the CID does not exists we treat the inline attachment as a normal attachment
            //            if (htmlBody && !string.IsNullOrEmpty(attach.ContentId) && body.Contains(attach.ContentId))
            //                body = body.Replace("cid:" + attach.ContentId, fileInfo.FullName);
            //            else
            //                // If we didn't find the cid tag we treat the inline attachment as a normal one 
            //                isInline = false;
            //        }
            //    }
            //    // ReSharper disable CanBeReplacedWithTryCastAndCheckForNull
            //    else if (attachment is Storage.Message)
            //    // ReSharper restore CanBeReplacedWithTryCastAndCheckForNull
            //    {
            //        var msg = (Storage.Message)attachment;
            //        attachmentFileName = msg.FileName;
            //        renderingPosition = msg.RenderingPosition;

            //        fileInfo = new FileInfo(FileManager.FileExistsMakeNew(outputFolder + attachmentFileName));
            //        msg.Save(fileInfo.FullName);
            //    }

            //    if (fileInfo == null) continue;

            //    if (!isInline)
            //        files.Add(fileInfo.FullName);

            //    // Check if the attachment has a render position. This property is only filled when the
            //    // body is RTF and the attachment is made inline
            //    if (htmlBody && renderingPosition != -1)
            //    {
            //        if (!isInline)
            //            using (var icon = FileIcon.GetFileIcon(fileInfo.FullName))
            //            {
            //                var iconFileName = outputFolder + Guid.NewGuid() + ".png";
            //                icon.Save(iconFileName, ImageFormat.Png);
            //                inlineAttachments.Add(renderingPosition,
            //                    iconFileName + "|" + attachmentFileName + "|" + fileInfo.FullName);
            //            }
            //        else
            //            inlineAttachments.Add(renderingPosition, attachmentFileName);
            //    }
            //    else
            //        renderingPosition = -1;

            //    if (!isInline && renderingPosition == -1)
            //    {
            //        if (htmlBody)
            //        {
            //            if (hyperlinks)
            //                attachments.Add("<a href=\"" + fileInfo.Name + "\">" +
            //                                HttpUtility.HtmlEncode(attachmentFileName) + "</a> (" +
            //                                FileManager.GetFileSizeString(fileInfo.Length) + ")");
            //            else
            //                attachments.Add(HttpUtility.HtmlEncode(attachmentFileName) + " (" +
            //                                FileManager.GetFileSizeString(fileInfo.Length) + ")");
            //        }
            //        else
            //            attachments.Add(attachmentFileName + " (" + FileManager.GetFileSizeString(fileInfo.Length) + ")");
            //    }
            //}

            //if (htmlBody)
            //    foreach (var inlineAttachment in inlineAttachments)
            //    {
            //        var names = inlineAttachment.Value.Split('|');

            //        if (names.Length == 3)
            //            body = ReplaceFirstOccurence(body, rtfInlineObject,
            //                "<table style=\"width: 70px; display: inline; text-align: center; font-family: Times New Roman; font-size: 12pt;\"><tr><td>" +
            //                (hyperlinks ? "<a href=\"" + names[2] + "\">" : string.Empty) + "<img alt=\"\" src=\"" +
            //                names[0] + "\">" + (hyperlinks ? "</a>" : string.Empty) + "</td></tr><tr><td>" +
            //                HttpUtility.HtmlEncode(names[1]) +
            //                "</td></tr></table>");
            //        else
            //            body = ReplaceFirstOccurence(body, rtfInlineObject, "<img alt=\"\" src=\"" + names[0] + "\">");
            //    }
        }
        #endregion
    }
}