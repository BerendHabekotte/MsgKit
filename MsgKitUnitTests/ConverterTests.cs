using Microsoft.VisualStudio.TestTools.UnitTesting;
using MimeKit;
using MsgKit;
using MsgKit.Enums;
using System;
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
            // Arrange
            var msgStream = new MemoryStream(Properties.Resources.TEST);
            var emlStream = new MemoryStream();
            using (var msg = new MsgReader.Outlook.Storage.Message(msgStream))
            {
                var expectedId = RemoveTags(msg.Id);
                var expectedDate = new DateTimeOffset((DateTime)msg.SentOn);
                var expectedFromName = msg.Sender.DisplayName;
                var expectedFromAddress = msg.Sender.Email;
                var expectedSubject = msg.Subject;
                var expectedPriority = msg.Headers.UnknownHeaders["X-Priority"].MapToPriority();
                var expectedImportance = msg.Importance.Map();
                var expectedTo = msg.Recipients.First(r => r.Type == MsgReader.Outlook.RecipientType.To);
                var expectedBodyText = msg.BodyText;

                var expectedDateHeader = new DateTimeOffset(DateTime.Parse(msg.Headers.Date));
                var expectedFromHeader = msg.Headers.From;
                var expectedImportanceHeader = msg.Headers.Importance.ToString();
                var expectedMessageIdHeader = msg.Headers.MessageId;
                var expectedReceivedHeader = msg.Headers.Received.First().Raw;
                var expectedMimeVersionHeader = msg.Headers.MimeVersion;
                var expectedSubjectHeader = msg.Headers.Subject;

                // Act
                Converter.ConvertMsgToEml(msgStream, emlStream);
                var actual = MimeMessage.Load(stream: emlStream, persistent: true) ;
                emlStream.Position = 0;
                using (var fileStream = new System.IO.FileStream("TestEmlMessage.eml", FileMode.Create))
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
                    Assert.AreEqual(expectedTo.Email, actual.To.First().Name);
                }
                else
                {
                    Assert.AreEqual(expectedTo.DisplayName, actual.To.First().Name);
                }
                Assert.AreEqual(expectedTo.Email, (actual.To.First() as MimeKit.MailboxAddress).Address);
                Assert.AreEqual(expectedBodyText, actual.TextBody);
                var actualDateHeaderValue = GetHeaderValue(actual, HeaderId.Date);
                var actualDateHeader = new DateTimeOffset(DateTime.Parse(actualDateHeaderValue));
                Assert.AreEqual(expectedDateHeader, actualDateHeader);
                Assert.IsTrue(actual.Headers.IndexOf(HeaderId.DispositionNotificationTo) == -1);
                var expectedFromHeaderValue = RemoveQuotes(expectedFromHeader.MailAddress.ToString());
                var actualFromHeader = GetHeaderValue(actual, HeaderId.From); ;
                Assert.AreEqual(expectedFromHeaderValue, actualFromHeader);
                var actualImportanceHeader = GetHeaderValue(actual, HeaderId.Importance);
                Assert.AreEqual(expectedImportanceHeader.ToUpper(), actualImportanceHeader.ToUpper());
                Assert.IsTrue(actual.Headers.IndexOf(HeaderId.InReplyTo) == -1);
                Assert.IsTrue(actual.Headers.IndexOf(HeaderId.Keywords) == -1);
                var actualMessageIdHeader = RemoveTags(GetHeaderValue(actual, HeaderId.MessageId));
                Assert.AreEqual(expectedMessageIdHeader, actualMessageIdHeader);
                Assert.AreEqual(expectedMimeVersionHeader, actual.Headers.First(h => h.Id == HeaderId.MimeVersion).Value);
                var actualReceived = ReplaceTabs(RemoveQuotes(actual.Headers.First(h => h.Id == HeaderId.Received).Value));
                Assert.AreEqual(expectedReceivedHeader, actualReceived);
                var actualSubjectHeader = GetHeaderValue(actual, HeaderId.Subject);
                Assert.AreEqual(expectedSubjectHeader, actualSubjectHeader);
                var keys = msg.Headers.UnknownHeaders.AllKeys;
                foreach (var key in msg.Headers.UnknownHeaders.AllKeys)
                {
                    var keyFound = Enum.TryParse(RemoveSeparators(key), true, out HeaderId id);
                    if (!keyFound)
                    {
                        continue;
                    }
                    if (id == HeaderId.ContentLanguage)
                    {
                        // TODO BHA: Convert the Content-Language header
                        continue;
                    }
                    Assert.IsTrue(keyFound);
                    Assert.IsTrue(actual.Headers.IndexOf(id) != -1);
                    var expectedHeader = RemoveSpacesAndTabsAndQuotes(msg.Headers.UnknownHeaders[key]);
                    var actualHeader = RemoveSpacesAndTabsAndQuotes(actual.Headers[id]);
                    Assert.AreEqual(expectedHeader, actualHeader);
                }
            }
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
            foreach(var header in actual.Headers)
            {
                var key = header.Id;
                if (key != HeaderId.Unknown)
                {
                    Assert.IsTrue(actual.Headers.IndexOf(key) != -1);
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
    }
}
