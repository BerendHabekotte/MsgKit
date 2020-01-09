using Microsoft.VisualStudio.TestTools.UnitTesting;
using MimeKit;
using MsgKit;
using MsgKit.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MsgKitUnitTests
{
    [TestClass]
    public class ConverterTests
    {
        [TestMethod]
        public void ConvertMsgToEmlBasics()
        {
            ConvertMsgMessageTest("TEST");
        }

        [TestMethod]
        public void ConvertMsgToEmlBasicsIMap()
        {
            ConvertMsgMessageTest("TEST_IMap");
        }

        [TestMethod]
        public void ConvertMsgWithAttachmentToEml()
        {
            ConvertMsgMessageTest("Email_with_attachment");
        }

        [TestMethod]
        public void ConvertMsgWithEmailAttachmentToEml()
        {
            ConvertMsgMessageTest("Email_with_eml_email_attachment");
        }

        [TestMethod]
        public void ConvertMsgWithInlineAttachmentToEml()
        {
            ConvertMsgMessageTest("Linked_resource");
        }

        [TestMethod]
        public void ConvertMsgWithNotifacations()
        {
            ConvertMsgMessageTest("Message_with_received_and_read_notification_and_vote_buttons");
        }

        [TestMethod]
        public void ConvertMsgDeliveredNotification()
        {
            ConvertMsgMessageTest("Delivered_Report");
        }

        [TestMethod]
        public void ConvertMsgReadNotification()
        {
            ConvertMsgMessageTest("Read_Report");
        }

        [TestMethod]
        public void ConvertEmlToMsgDeliveredNotification()
        {
            ConvertEmlMessageTest("Delivered_Report_AsEml");
        }

        [TestMethod]
        public void ConvertMsgVote()
        {
            ConvertMsgMessageTest("Vote");
        }


        [TestMethod]
        public void LoadEmlMessage()
        {
            // Arrange
            var emlStream = new MemoryStream(Properties.Resources.TEST_AsEml);
            emlStream.Seek(0, SeekOrigin.Begin);
            var expectedbodyText = "TEST\r\n";

            // Act
            var actual = MimeMessage.Load(emlStream);

            // Assert
            Assert.AreEqual(expectedbodyText, actual.TextBody);
            foreach (var header in actual.Headers)
            {
                var key = header.Id;
                if (key != HeaderId.Unknown)
                {
                    Assert.IsTrue(actual.Headers.IndexOf(key) != -1);
                }
            }
        }

        private void ConvertEmlMessageTest(string resourceName)
        {
            // Arrange
            var emlResource = (byte[])Properties.Resources.ResourceManager.GetObject(resourceName);
            using (var emlStream = new MemoryStream(emlResource))
            using (var msgStream = new MemoryStream())
            {
                var mimeMessage = MimeMessage.Load(emlStream);
                var expectedId = mimeMessage.MessageId;


                emlStream.Seek(0, SeekOrigin.Begin);

                // Act
                MsgReader.Outlook.Storage.Message actual;
                Converter.ConvertEmlToMsg(emlStream, msgStream);
                actual = new MsgReader.Outlook.Storage.Message(msgStream);
                msgStream.Position = 0;
                using (var fileStream = new FileStream($"{resourceName}_Converted.msg", FileMode.Create))
                {
                    msgStream.WriteTo(fileStream);
                }

                // Assert
                Assert.AreEqual(expectedId, RemoveTags(actual.Id));
            }
        }

        private void ConvertMsgMessageTest(string resourceName)
        {
            // Arrange
            var msgResource = (byte[])Properties.Resources.ResourceManager.GetObject(resourceName);
            using (var msgStream = new MemoryStream(msgResource))
            using (var msgMessage = new MsgReader.Outlook.Storage.Message(msgStream))
            using (var emlStream = new MemoryStream())
            {
                var expectedId = RemoveTags(msgMessage.Id);
                var expectedDate = new DateTimeOffset((DateTime)msgMessage.SentOn);
                var expectedFromName = msgMessage.Sender.DisplayName;
                var expectedFromAddress = msgMessage.Sender.Email;
                var expectedSubject = msgMessage.Subject;
                var expectedPriority = msgMessage.Headers.UnknownHeaders["X-Priority"].MapToPriority();
                var expectedImportance = msgMessage.Importance.Map();
                var expectedTo = msgMessage.Headers.To.First();
                var expectedBodyText = msgMessage.BodyText;
                var expectedBodyHtml = msgMessage.BodyHtml;
                var expectedBodyRtf = msgMessage.BodyRtf;

                var expectedDateHeader = new DateTimeOffset(DateTime.Parse(msgMessage.Headers.Date));
                var expectedFromHeader = msgMessage.Headers.From;
                var expectedImportanceHeader = msgMessage.Headers.Importance.ToString();
                var expectedMessageIdHeader = msgMessage.Headers.MessageId;
                var expectedReceivedHeader = msgMessage.Headers.Received.First().Raw;
                var expectedMimeVersionHeader = msgMessage.Headers.MimeVersion;
                var expectedSubjectHeader = msgMessage.Headers.Subject;
                var expectedDispositionHeader = msgMessage.Headers.DispositionNotificationTo;
                var expectedInReplyToHeader = msgMessage.Headers.InReplyTo;
                var expectedKeywordsHeader = msgMessage.Headers.Keywords;

                var attachments = 
                    msgMessage.Attachments
                        .Where(a => a.GetType() == typeof(MsgReader.Outlook.Storage.Attachment) &&
                          ((MsgReader.Outlook.Storage.Attachment)a).IsInline == false);
                var expectedAttachmemtCount = attachments.Count();
                MsgReader.Outlook.Storage.Attachment expectedAttachment = null;
                var expectedAttachmentFileName = string.Empty;
                if (expectedAttachmemtCount > 0)
                {
                    expectedAttachment =
                        (MsgReader.Outlook.Storage.Attachment)attachments
                        .Where(a => ((MsgReader.Outlook.Storage.Attachment)a).IsInline == false)
                        .FirstOrDefault();
                    if (expectedAttachment != null)
                    {
                        expectedAttachmentFileName = expectedAttachment.FileName;
                    }
                }

                var messageAttachments =
                    msgMessage.Attachments
                        .Where(a => a.GetType() == typeof(MsgReader.Outlook.Storage.Message));
                MsgReader.Outlook.Storage.Message expectedMessageAttachment = null;
                var expectedAttachedMessageName = string.Empty;
                if (messageAttachments.Count() > 0)
                {
                    expectedMessageAttachment =
                            (MsgReader.Outlook.Storage.Message)messageAttachments
                            .FirstOrDefault();
                    if (expectedMessageAttachment != null)
                    {
                        expectedAttachedMessageName = expectedMessageAttachment.FileName;
                    }
                }
                expectedAttachmemtCount += messageAttachments.Count();

                if (expectedAttachment != null)
                {
                    expectedAttachmentFileName = expectedAttachment.FileName;
                }
                if (expectedMessageAttachment != null)
                {
                }

                var linkedResources = msgMessage.Attachments
                    .Where(a => a.GetType() == typeof(MsgReader.Outlook.Storage.Attachment) &&
                        ((MsgReader.Outlook.Storage.Attachment)a).IsInline == true);
                var expectedLinkedResourcesCount = linkedResources.Count();
                MsgReader.Outlook.Storage.Attachment expectedLinkedResource = null;
                var expectedLinkedResourceFilename = string.Empty;
                if (expectedLinkedResourcesCount > 0)
                {
                    expectedLinkedResource = (MsgReader.Outlook.Storage.Attachment)linkedResources.FirstOrDefault();
                    if (expectedLinkedResource != null)
                    {
                        expectedLinkedResourceFilename = expectedLinkedResource.FileName;
                    }
                }

                // Act
                MimeMessage actual;
                Converter.ConvertMsgToEml(msgStream, emlStream);
                actual = MimeMessage.Load(stream: emlStream, persistent: true);
                emlStream.Position = 0;
                using (var fileStream = new FileStream($"{resourceName}.eml", FileMode.Create))
                {
                    emlStream.WriteTo(fileStream);
                }

                // Assert
                Assert.AreEqual(expectedId, actual.MessageId);
                Assert.AreEqual(expectedDate, actual.Date);
                var actualFrom = actual.From.Mailboxes.FirstOrDefault();
                Assert.IsNotNull(actualFrom);
                Assert.AreEqual(expectedFromName, actualFrom.Name);
                Assert.AreEqual(expectedFromAddress, actualFrom.Address);
                Assert.AreEqual(expectedSubject, actual.Subject);
                Assert.AreEqual(expectedPriority, actual.Priority);
                Assert.AreEqual(expectedImportance, actual.Importance);
                if (string.IsNullOrEmpty(expectedTo.DisplayName))
                {
                    Assert.AreEqual(expectedTo.Address, actual.To.First().Name);
                }
                else
                {
                    Assert.AreEqual(expectedTo.DisplayName, actual.To.First().Name);
                }
                Assert.AreEqual(expectedTo.Address, (actual.To.First() as MimeKit.MailboxAddress).Address);
                Assert.AreEqual(NullIfNoContent(expectedBodyText), NullIfNoContent(actual.TextBody));
                Assert.AreEqual(expectedBodyHtml, actual.HtmlBody);
                if (expectedBodyRtf != null)
                {
                    MultipartAlternative multipartAlternative;
                    if (actual.Body as MultipartAlternative != null)
                    {
                        multipartAlternative = (MultipartAlternative)actual.Body;
                    }
                    else
                    {
                        multipartAlternative = (MultipartAlternative)(actual.Body as Multipart)
                            .First(p => p.GetType() == typeof(MultipartAlternative));
                    }
                    var rtfPart = multipartAlternative.First(p => p.GetType() == typeof(TextPart) && ((TextPart)p).IsRichText);
                    var actualBodyRtf = ((TextPart)rtfPart).Text.Replace("\r\n\r", "\n\r");
                    Assert.AreEqual(expectedBodyRtf, actualBodyRtf);
                }
                var actualDateHeaderValue = GetHeaderValue(actual, HeaderId.Date);
                var actualDateHeader = new DateTimeOffset(DateTime.Parse(actualDateHeaderValue));
                Assert.AreEqual(expectedDateHeader, actualDateHeader);
                AssertHeader(expectedDispositionHeader, actual, HeaderId.DispositionNotificationTo);
                var expectedFromHeaderValue = RemoveQuotes(expectedFromHeader.MailAddress.ToString());
                var actualFromHeader = GetHeaderValue(actual, HeaderId.From); ;
                Assert.AreEqual(expectedFromHeaderValue, ReplaceTabs(actualFromHeader));
                var actualImportanceHeader = GetHeaderValue(actual, HeaderId.Importance);
                Assert.AreEqual(expectedImportanceHeader.ToUpper(), actualImportanceHeader.ToUpper());
                AssertHeader(expectedInReplyToHeader, actual, HeaderId.InReplyTo);
                AssertHeader(expectedKeywordsHeader, actual, HeaderId.Keywords);
                var actualMessageIdHeader = RemoveTags(GetHeaderValue(actual, HeaderId.MessageId));
                Assert.AreEqual(expectedMessageIdHeader, actualMessageIdHeader);
                Assert.AreEqual(expectedMimeVersionHeader, actual.Headers.First(h => h.Id == HeaderId.MimeVersion).Value);
                var actualReceived = ReplaceTabs(RemoveQuotes(actual.Headers.First(h => h.Id == HeaderId.Received).Value));
                Assert.AreEqual(expectedReceivedHeader, actualReceived);
                var actualSubjectHeader = GetHeaderValue(actual, HeaderId.Subject);
                Assert.AreEqual(expectedSubjectHeader, actualSubjectHeader);
                var keys = msgMessage.Headers.UnknownHeaders.AllKeys;
                foreach (var key in msgMessage.Headers.UnknownHeaders.AllKeys)
                {
                    var keyFound = Enum.TryParse(RemoveSeparators(key), true, out HeaderId id);
                    if (!keyFound)
                    {
                        continue;
                    }
                    if (id == HeaderId.ContentLanguage)
                    {
                        // Content-Language header disappears mysteriously
                        continue;
                    }
                    Assert.IsTrue(keyFound);
                    Assert.IsTrue(actual.Headers.IndexOf(id) != -1);
                    var expectedHeader = RemoveSpacesAndTabsAndQuotes(msgMessage.Headers.UnknownHeaders[key]);
                    var actualHeader = RemoveSpacesAndTabsAndQuotes(actual.Headers[id]);
                    Assert.AreEqual(expectedHeader, actualHeader);
                }

                var actualAttachmentCount = actual.Attachments.Count();
                Assert.AreEqual(expectedAttachmemtCount, actualAttachmentCount);
                MimeEntity actualAttachment = null;
                if (actualAttachmentCount > 0)
                {
                    actualAttachment = actual.Attachments.First();
                }
                Assert.AreEqual(expectedAttachment == null && expectedMessageAttachment == null, actualAttachment == null);
                if (expectedMessageAttachment != null)
                {
                    actualAttachment = actual.Attachments.First(a => ((MimePart)a).FileName.StartsWith(expectedMessageAttachment.Subject));
                    Assert.IsNotNull(actualAttachment);
                }
                if (expectedAttachment != null)
                {
                    actualAttachment = actual.Attachments.First(a => ((MimePart)a).FileName == expectedAttachment.FileName);
                    Assert.IsNotNull(actualAttachment);
                }
                var actualLinkedResources = GetInlineImages(actual);
                var actualLinkedResourcesCount = actualLinkedResources.Count();
                MimePart actualLinkedResource = null;
                Assert.AreEqual(expectedLinkedResourcesCount, actualLinkedResourcesCount);
                if (actualLinkedResourcesCount > 0)
                {
                    actualLinkedResource = (MimePart)actualLinkedResources.First();
                }
                Assert.AreEqual(expectedLinkedResource == null, actualLinkedResource == null);
                if (expectedLinkedResource != null)
                {
                    Assert.AreEqual(expectedLinkedResource.FileName, actualLinkedResource.FileName);
                }
            }
        }

        private string GetHeaderValue(MimeMessage message, HeaderId id)
        {
            return message.Headers.First(h => h.Id == id).Value;
        }

        private string RemoveTags(string value)
        {
            return value.Trim('<', '>', '/');
        }

        private string RemoveQuotes(string value)
        {
            return value.Replace("\"", string.Empty);
        }

        private string ReplaceTabs(string value)
        {
            return value.Replace("\t", " ");
        }

        private string RemoveSpacesAndTabsAndQuotes(string value)
        {
            var result = value.Replace("\t", string.Empty);
            result = result.Replace("\"", string.Empty);
            result = result.Replace(" ", string.Empty);
            return result.Trim(';');
        }

        private string RemoveSeparators(string value)
        {
            return value.Replace("-", string.Empty);
        }

        private string NullIfNoContent(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return null;
            }
            var result = value.Replace("\n", "").Replace("\r", "").Trim();
            if (string.IsNullOrEmpty(result))
            {
                return null;
            }
            return value;
        }

        private void AssertHeader(
            System.Collections.Generic.List<string> expectedValues,
            MimeMessage actual,
            HeaderId actualId)
        {
            if (expectedValues == null || expectedValues.Count == 0)
            {
                Assert.IsTrue(actual.Headers.IndexOf(actualId) == -1);
                return;
            }
            Assert.IsTrue(actual.Headers.IndexOf(actualId) >= 0);
            var actualValues = actual.Headers[actualId];
            foreach (var expectedValue in expectedValues)
            {
                Assert.IsTrue(actualValues.Contains(expectedValue));
            }
        }

        private void AssertHeader(
            System.Collections.Generic.List<MsgReader.Mime.Header.RfcMailAddress> expectedValues,
            MimeMessage actual,
            HeaderId actualId)
        {
            if (expectedValues == null || expectedValues.Count == 0)
            {
                Assert.IsTrue(actual.Headers.IndexOf(actualId) == -1);
                return;
            }
            Assert.IsTrue(actual.Headers.IndexOf(actualId) >= 0);
            var actualValues = actual.Headers[actualId].Split(';');
            foreach (var expectedValue in expectedValues)
            {
                Assert.IsNotNull(actualValues.FirstOrDefault(a => a.Contains(expectedValue.Address)));
            }
        }

        private IEnumerable<MimeEntity> GetInlineImages(MimeMessage message)
        {
            var mimeParts = message.BodyParts
                .Where(m => m.ContentId != null && ((MimePart)m).Content != null && m.ContentType.MediaType == "image" && (message.HtmlBody.IndexOf("cid:" + m.ContentId) > -1))
                .ToList();
            return mimeParts;
        }
    }
}
