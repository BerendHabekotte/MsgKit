using MimeKit;
using MsgKit.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MsgKit
{
    internal class MsgToMimeConverter
    {

        #region constructors

        internal MsgToMimeConverter(MsgReader.Outlook.Storage.Message msgMessage, MimeMessage mimeMessage)
        {
            if (msgMessage == null) { throw new ArgumentNullException(nameof(msgMessage)); }
            if (mimeMessage == null) { throw new ArgumentNullException(nameof(mimeMessage)); }
            this.msgMessage = msgMessage;
            this.mimeMessage = mimeMessage;
        }

        #endregion

        internal void Convert()
        {
            ConvertSenderToFrom();
            ConvertRepresentingToResentSender();
            mimeMessage.Subject = msgMessage.Subject;
            ConvertSentOnToDate();
            mimeMessage.MessageId = msgMessage.Id;
            ConvertPriority();
            ConvertImportance();
            ConvertRecipients();
            ConvertBody();
            ConvertContentLanguageHeader();
            ConvertHeaders();
        }


        #region private methods

        private void ConvertSenderToFrom()
        {
            if (msgMessage.Sender == null)
            {
                return;
            }
            var from = new MailboxAddress(
                    Encoding.ASCII,
                    msgMessage.Sender.DisplayName,
                    msgMessage.Sender.Email);
            mimeMessage.From.Add(from);
        }

        private void ConvertRepresentingToResentSender()
        {
            if (msgMessage.SenderRepresenting == null)
            {
                return;
            }
            var resentSender = new MailboxAddress(
                    Encoding.ASCII,
                    msgMessage.SenderRepresenting.DisplayName,
                    msgMessage.SenderRepresenting.Email);
            mimeMessage.ResentSender = resentSender;
        }

        private void ConvertSentOnToDate()
        {
            if (msgMessage.SentOn != null)
            {
                mimeMessage.Date = new DateTimeOffset((DateTime)msgMessage.SentOn);
            }
        }

        private void ConvertPriority()
        {
            var priority = msgMessage.Headers.UnknownHeaders["X-Priority"];
            mimeMessage.Priority = priority.MapToPriority();
        }

        private void ConvertImportance()
        {
            mimeMessage.Importance = msgMessage.Importance.Map();
        }


        private void ConvertRecipients()
        {
            if (msgMessage.Recipients == null)
            {
                return;
            }
            if (msgMessage.Recipients.Count == 0)
            {
                return;
            }
            foreach (var recipient in msgMessage.Recipients)
            {
                if (recipient.Type == null)
                {
                    continue;
                }
                var displayName = string.IsNullOrEmpty(recipient.DisplayName) ? recipient.Email : recipient.DisplayName;
                var mailAddress = new MailboxAddress(displayName, recipient.Email);
                switch (recipient.Type)
                {
                    case MsgReader.Outlook.RecipientType.Bcc:
                        mimeMessage.Bcc.Add(mailAddress);
                        break;
                    case MsgReader.Outlook.RecipientType.Cc:
                        mimeMessage.Cc.Add(mailAddress);
                        break;
                    case MsgReader.Outlook.RecipientType.Resource:
                    case MsgReader.Outlook.RecipientType.Room:
                        break;
                    case MsgReader.Outlook.RecipientType.To:
                        mimeMessage.To.Add(mailAddress);
                        break;
                    case MsgReader.Outlook.RecipientType.Unknown:
                        break;
                }
            }
        }

        private void ConvertBody()
        {
            ConvertBodyByBodyBuilder();

            // This loops through the top-level parts (i.e. it doesn't open up attachments and continue to traverse).
            // As such, any included messages are just attachments here.

            //foreach (var bodyPart in eml.BodyParts)
            //{
            //    var handled = false;

            //    // If the part hasn't previously been handled by "body" part handling
            //    if (!handled)
            //    {
            //        var attachmentStream = new MemoryStream();
            //        var fileName = bodyPart.ContentType.Name;
            //        var extension = string.Empty;

            //        if (bodyPart is MessagePart messagePart)
            //        {
            //            messagePart.Message.WriteTo(attachmentStream);
            //            if (messagePart.Message != null)
            //                fileName = messagePart.Message.Subject;

            //            extension = ".eml";
            //        }
            //        else if (bodyPart is MessageDispositionNotification)
            //        {
            //            var part = (MessageDispositionNotification)bodyPart;
            //            fileName = part.FileName;
            //        }
            //        else if (bodyPart is MessageDeliveryStatus)
            //        {
            //            var part = (MessageDeliveryStatus)bodyPart;
            //            fileName = "details";
            //            extension = ".txt";
            //            part.WriteTo(FormatOptions.Default, attachmentStream, true);
            //        }
            //        else
            //        {
            //            var part = (MimePart)bodyPart;
            //            part.Content.DecodeTo(attachmentStream);
            //            fileName = part.FileName;
            //            bodyPart.WriteTo(attachmentStream);
            //        }

            //        fileName = string.IsNullOrWhiteSpace(fileName)
            //            ? $"part_{++namelessCount:00}"
            //            : FileManager.RemoveInvalidFileNameChars(fileName);

            //        if (!string.IsNullOrEmpty(extension))
            //            fileName += extension;

            //        var inline = bodyPart.ContentDisposition != null &&
            //            bodyPart.ContentDisposition.Disposition.Equals("inline",
            //                StringComparison.InvariantCultureIgnoreCase);

            //        attachmentStream.Position = 0;
            //        msg.Attachments.Add(attachmentStream, fileName, -1, inline, bodyPart.ContentId);
            //    }
            //}

            //mimeMessage.Body = multiPart;
        }

        private void ConvertBodyByBodyBuilder()
        {
            var builder = new BodyBuilder
            {
                TextBody = msgMessage.BodyText,
                HtmlBody = msgMessage.BodyHtml
            };
            mimeMessage.Body = builder.ToMessageBody();
            mimeMessage.Body.ContentType.Boundary = msgMessage.Headers.ContentType.Boundary;
        }

        private void ConvertContentLanguageHeader()
        {
            var language = msgMessage.Headers.UnknownHeaders["Content-Language"];
            if (string.IsNullOrEmpty(language))
            {
                return;
            }
            mimeMessage.Body.Headers.Add(HeaderId.ContentLanguage, language);
        }

        /// <summary>
        /// Since we use the BodyBuilder approach to generate the body, and the BodyBuilder class 
        /// misses the RftBody property, we do not convert the <see cref="MsgReader.Outlook.Storage.Message.BodyRtf"/> property.
        /// </summary>
        /// <param name="body">The MimeMessage body target.</param>
        private void ConvertMsgToEmlBodyRtf(MultipartAlternative body)
        {
            if (msgMessage.BodyRtf == null)
            {
                return;
            }
            var rtfPart = new TextPart(MimeKit.Text.TextFormat.RichText)
            {
                Text = msgMessage.BodyRtf
            };
            body.Add(rtfPart);
        }

        private void ConvertHeaders(
            )
        {
            var msgHeaders = msgMessage.Headers;
            if (msgHeaders == null)
            {
                return;
            }
            ConvertHeaderAddressList(HeaderId.Bcc, msgHeaders.Bcc);
            ConvertHeaderAddressList(HeaderId.Cc, msgHeaders.Cc);
            ConvertHeaderString(HeaderId.ContentDescription, msgHeaders.ContentDescription);
            ConvertHeaderObject(HeaderId.ContentDisposition, msgHeaders.ContentDisposition);
            ConvertHeaderString(HeaderId.ContentId, msgHeaders.ContentId);
            ConvertHeaderObject(HeaderId.ContentTransferEncoding, msgHeaders.ContentTransferEncoding);
            ConvertHeaderObject(HeaderId.ContentType, msgHeaders.ContentType);
            ConvertHeaderString(HeaderId.Date, msgHeaders.Date);
            ConvertHeaderAddressList(HeaderId.DispositionNotificationTo, msgHeaders.DispositionNotificationTo);
            ConvertHeaderAddress(HeaderId.From, msgHeaders.From);
            ConvertHeaderObject(HeaderId.Importance, msgHeaders.Importance);
            ConvertHeaderStringList(HeaderId.InReplyTo, msgHeaders.InReplyTo);
            ConvertHeaderStringList(HeaderId.Keywords, msgHeaders.Keywords);
            ConvertHeaderString(HeaderId.MessageId, msgHeaders.MessageId);
            ConvertHeaderString(HeaderId.MimeVersion, msgHeaders.MimeVersion);
            ConvertHeaderReceived();
            ConvertHeaderStringList(HeaderId.References, msgHeaders.References);
            ConvertHeaderAddress(HeaderId.ReplyTo, msgHeaders.ReplyTo);
            ConvertHeaderAddress(HeaderId.ReturnPath, msgHeaders.ReturnPath);
            ConvertHeaderAddress(HeaderId.Sender, msgHeaders.Sender);
            ConvertHeaderString(HeaderId.Subject, msgHeaders.Subject);
            ConvertHeaderAddressList(HeaderId.To, msgHeaders.To);
            ConvertHeadersUnknown();
        }

        private void ConvertHeadersUnknown()
        {
            foreach (var key in msgMessage.Headers.UnknownHeaders.AllKeys)
            {
                var value = msgMessage.Headers.UnknownHeaders[key];
                HeaderId id;
                var keyString = key.Replace("-", string.Empty);
                if (Enum.TryParse(keyString, true, out id))
                {
                    ConvertHeaderString(id, value);
                }
                else
                {
                    if (key != "Content-Language")
                    {
                        mimeMessage.Headers.Add(key, Encoding.ASCII, value);
                    }
                }
            }
        }

        private void ConvertHeaderReceived()
        {
            if (msgMessage.Headers.Received == null)
            {
                return;
            }
            foreach (var received in msgMessage.Headers.Received)
            {
                ConvertHeaderString(HeaderId.Received, received.Raw);
            }
        }

        private void ConvertHeaderString(HeaderId headerId, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }
            if (mimeMessage.Headers.Any(h => h.Id == headerId))
            {
                return;
            }
            var header = new Header(headerId, value);
            mimeMessage.Headers.Add(header);
        }

        private void ConvertHeaderString(string field, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }
            if (mimeMessage.Headers.Any(h => h.Field == field))
            {
                return;
            }
            mimeMessage.Headers.Add(field, Encoding.ASCII, value);
        }



        private void ConvertHeaderObject(HeaderId headerId, object value)
        {
            if (value == null)
            {
                return;
            }
            ConvertHeaderString(headerId, value.ToString());
        }


        private void ConvertHeaderAddressList(
            HeaderId headerId,
            List<MsgReader.Mime.Header.RfcMailAddress> addresses)
        {
            if (mimeMessage.Headers.Any(h => h.Id == headerId))
            {
                return;
            }
            foreach (var address in addresses)
            {
                ConvertHeaderAddress(headerId, address);
            }
        }

        private void ConvertHeaderAddress(
            HeaderId headerId,
            MsgReader.Mime.Header.RfcMailAddress address)
        {
            if (address == null)
            {
                return;
            }
            if (string.IsNullOrEmpty(address.Raw))
            {
                return;
            }
            ConvertHeaderString(headerId, address.Raw);
        }

        private void ConvertHeaderStringList(
            HeaderId headerId,
            List<string> values)
        {
            if (mimeMessage.Headers.Any(h => h.Id == headerId))
            {
                return;
            }
            foreach (var value in values)
            {
                ConvertHeaderString(headerId, value);
            }
        }

        #endregion

        #region private fields

        MsgReader.Outlook.Storage.Message msgMessage;
        MimeMessage mimeMessage;

        #endregion



    }
}
