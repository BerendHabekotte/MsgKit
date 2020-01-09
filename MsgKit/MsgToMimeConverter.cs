using MimeKit;
using MsgKit.Enums;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace MsgKit
{
    internal class MsgToMimeConverter
    {

        #region constructors

        internal MsgToMimeConverter(MsgReader.Outlook.Storage.Message msgMessage, MimeMessage mimeMessage)
        {
            this.msgMessage = msgMessage ?? throw new ArgumentNullException(nameof(msgMessage));
            this.mimeMessage = mimeMessage ?? throw new ArgumentNullException(nameof(mimeMessage));
            bodyBuilder = null;
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
            if (msgMessage.SenderRepresenting.AddressType != "SMTP")
            {
                return;
            }
            var address = msgMessage.SenderRepresenting.Email;
            var resentSender = new MailboxAddress(
                    Encoding.ASCII,
                    msgMessage.SenderRepresenting.DisplayName,
                    address);
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
                    case MsgReader.Outlook.RecipientType.Unknown:
                        break;
                }
            }
            if (msgMessage.Headers.To == null)
            {
                return;
            }
            foreach(var address in msgMessage.Headers.To)
            {
                var displayName = string.IsNullOrEmpty(address.DisplayName) ? address.Address : address.DisplayName;
                var mailAddress = new MailboxAddress(displayName, address.Address);
                mimeMessage.To.Add(mailAddress);
            }
        }

        private void ConvertBody()
        {
            if (msgMessage.Headers.ContentType.MediaType == "multipart/report")
            {
                if (ConvertReport())
                {
                    return;
                }
            }
            if (msgMessage.Headers.ContentType.MediaType == "application/ms-tnef")
            {
                if (ConvertMsTnef())
                {
                    return;
                }
            }
            ConvertBodyByBodyBuilder();
        }

        private bool ConvertReport()
        {
            if (msgMessage.Headers.ContentType.MediaType == "multipart/report")
            {
                return ConvertMultipartReport();
            }
            if (msgMessage.SubjectPrefix != null && msgMessage.Headers.ContentType.MediaType == "application/ms-tnef")
            {
                return ConvertMsTnef();
            }
            return false;
        }

        private bool ConvertMultipartReport()
        {
            if (!msgMessage.Headers.ContentType.Parameters.ContainsKey("report-type"))
            {
                return false;
            }
            var bodyPart = new MultipartReport(reportType: msgMessage.Headers.ContentType.Parameters["report-type"]);
            ConvertAttachments(bodyPart);
            if (msgMessage.Headers.ContentType.Parameters["report-type"] == "delivery-status")
            {
                ConvertDeliveryStatusReport(bodyPart);
                return true;
            }
            if (msgMessage.Headers.ContentType.Parameters["report-type"] == "disposition-notification")
            {
                ConvertDispositionNotificationReport(bodyPart);
                return true;
            }
            return false;
        }

        private static bool ConvertMsTnef()
        {
            // TODO BHA
            return false;
        }

        private void ConvertAttachments(Multipart body)
        {
            foreach (var @object in msgMessage.Attachments)
            {
                var attachment = (@object as MsgReader.Outlook.Storage.Attachment);
                var message = (@object as MsgReader.Outlook.Storage.Message);
                if (attachment == null && message == null)
                {
                    continue;
                }
                if (message != null)
                {
                    ConvertMessageAttachment(body, message);
                    continue;
                }
                if (attachment == null)
                {
                    continue;
                }
                if (attachment.IsInline)
                {
                    ConvertInlineAttachment(body, attachment);
                    continue;
                }
                ConvertMimeAttachment(body, attachment);
            }
        }

        private static void ConvertMessageAttachment(Multipart body, MsgReader.Outlook.Storage.Message messageAttachment)
        {
            var dataStream = new MemoryStream();
            messageAttachment.Save(dataStream);
            dataStream.Seek(0, SeekOrigin.Begin);
            var mimePart = new MimePart
            {
                Content = new MimeContent(dataStream),
                FileName = messageAttachment.Subject + ".msg",
                ContentId = messageAttachment.Headers.ContentId,
                IsAttachment = true
            };
            body.Add(mimePart);
        }

        private void ConvertInlineAttachment(Multipart body, MsgReader.Outlook.Storage.Attachment attachment)
        {
            throw new NotImplementedException("ConvertInlineAttachment");
        }

        private static void ConvertMimeAttachment(Multipart body, MsgReader.Outlook.Storage.Attachment attachment)
        {
            var dataStream = new MemoryStream(attachment.Data);
            var mimePart = new MimePart
            {
                Content = new MimeContent(dataStream),
                FileName = attachment.FileName,
                ContentId = attachment.ContentId,
                IsAttachment = true
            };
            body.Add(mimePart);
        }

        private void ConvertDeliveryStatusReport(MultipartReport bodyPart)
        {
            // TODO BHA Add MessageDeliveryStatus.Headers, Make Email attachment readable
            var deliveryStatus = new MessageDeliveryStatus();
            bodyPart.Add(deliveryStatus);
            ConvertBodyPartAlternatives(bodyPart);
            bodyPart.ContentType.Boundary = msgMessage.Headers.ContentType.Boundary;
            mimeMessage.Body = bodyPart as MimeEntity;
        }

        private void ConvertDispositionNotificationReport(MultipartReport bodyPart)
        {
            // TODO BHA Add MessageDispositionNotification.Headers
            var dispositionNotification = new MessageDispositionNotification();
            bodyPart.Add(dispositionNotification);
            ConvertBodyPartAlternatives(bodyPart);
            bodyPart.ContentType.Boundary = msgMessage.Headers.ContentType.Boundary;
            mimeMessage.Body = bodyPart as MimeEntity;
        }

        private void ConvertBodyPartAlternatives(Multipart body)
        {
            var bodypartAlternative = new MultipartAlternative();
            ConvertBodyText(bodypartAlternative);
            ConvertBodyRtf(bodypartAlternative);
            ConvertBodyHtml(bodypartAlternative);
            body.Add(bodypartAlternative);
        }

        private void ConvertBodyText(MultipartAlternative body)
        {
            if (msgMessage.BodyText == null)
            {
                return;
            }
            var textPart = new TextPart(MimeKit.Text.TextFormat.Plain)
            {
                Text = msgMessage.BodyText
            };
            body.Add(textPart);
        }

        private void ConvertBodyHtml(MultipartAlternative body)
        {
            if (msgMessage.BodyHtml == null)
            {
                return;
            }
            var textPart = new TextPart(MimeKit.Text.TextFormat.Html)
            {
                Text = msgMessage.BodyHtml
            };
            body.Add(textPart);
        }

        private void ConvertBodyRtf(MultipartAlternative body)
        {
            if (msgMessage.BodyRtf == null)
            {
                return;
            }
            var textPart = new TextPart(MimeKit.Text.TextFormat.RichText)
            {
                Text = msgMessage.BodyRtf
            };
            body.Add(textPart);
        }

        private void ConvertBodyByBodyBuilder()
        {
            bodyBuilder = new BodyBuilder
            {
                TextBody = msgMessage.BodyText,
                HtmlBody = msgMessage.BodyHtml
            };
            ConvertAttachments();
            mimeMessage.Body = bodyBuilder.ToMessageBody();
            if (msgMessage.Headers.ContentType.MediaType == mimeMessage.Body.ContentType.MimeType &&
                msgMessage.Headers.ContentType.Boundary != null)
            {
                mimeMessage.Body.ContentType.Boundary = msgMessage.Headers.ContentType.Boundary;
            }
            ConvertBodyRtf();
        }

        private void ConvertAttachments()
        {
            foreach (var @object in msgMessage.Attachments)
            {
                var attachment = (@object as MsgReader.Outlook.Storage.Attachment);
                var message = (@object as MsgReader.Outlook.Storage.Message);
                if (message != null)
                {
                    ConvertMessageAttachment(message);
                    continue;
                }
                if (attachment == null)
                {
                    continue;
                }
                if (attachment.IsInline)
                {
                    ConvertInlineAttachment(attachment);
                    continue;
                }
                ConvertMimeAttachment(attachment);
            }
        }

        private void ConvertMessageAttachment(MsgReader.Outlook.Storage.Message messageAttachment)
        {
            var dataStream = new MemoryStream();
            messageAttachment.Save(dataStream);
            var mimePart = new MimePart
            {
                Content = new MimeContent(dataStream),
                FileName = messageAttachment.Subject + ".msg"
            };
            bodyBuilder.Attachments.Add(mimePart);
        }

        private void ConvertMimeAttachment(MsgReader.Outlook.Storage.Attachment attachment)
        {
            var dataStream = new MemoryStream(attachment.Data);
            var mimePart = new MimePart
            {
                Content = new MimeContent(dataStream),
                FileName = attachment.FileName,
                ContentId = attachment.ContentId
            };
            bodyBuilder.Attachments.Add(mimePart);
        }

        private void ConvertInlineAttachment(MsgReader.Outlook.Storage.Attachment attachment)
        {
            var linkedResource = bodyBuilder.LinkedResources.Add(fileName: attachment.FileName, data: attachment.Data);
            linkedResource.ContentId = attachment.ContentId;
        }

        private void ConvertBodyRtf()
        {
            if (string.IsNullOrEmpty(msgMessage.BodyRtf))
            {
                return;
            }
            var rtfBody = new TextPart(MimeKit.Text.TextFormat.RichText)
            {
                Text = msgMessage.BodyRtf,
                IsAttachment = false
            };
            if (mimeMessage.Body is MultipartAlternative)
            {
                ((MultipartAlternative)mimeMessage.Body).Add(rtfBody);
                return;
            }
            var multiPartAlternative =
                (mimeMessage.Body as Multipart)
                .FirstOrDefault(p => p.GetType() == typeof(MultipartAlternative));
            if (multiPartAlternative == null)
            {
                return;
            }
            ((MultipartAlternative)multiPartAlternative).Add(rtfBody);
        }

        private void ConvertHeaders()
        {
            var converter = new MsgToMimeHeaderConverter(msgMessage.Headers, mimeMessage.Headers);
            converter.Convert();
        }

        #endregion

        #region private fields

        readonly MsgReader.Outlook.Storage.Message msgMessage;
        readonly MimeMessage mimeMessage;
        BodyBuilder bodyBuilder;

        #endregion
    }
}
