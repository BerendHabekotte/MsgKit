using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsgKit
{
    /// <summary>
    /// Converts the headers of an msg message to the headers of a mime message.
    /// </summary>
    public class MsgToMimeHeaderConverter
    {
        #region private fields

        private MsgReader.Mime.Header.MessageHeader msgHeaders;
        private HeaderList mimeHeaders;

        #endregion

        #region constructors

        /// <summary>
        /// Constructor that takes msg headers as a conversion source and mime headers as the conversion target.
        /// </summary>
        /// <param name="source">Msg headers as the conversion source.</param>
        /// <param name="target">Mime headers as the conversion target.</param>
        public MsgToMimeHeaderConverter(MsgReader.Mime.Header.MessageHeader source, HeaderList target)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (target == null)
            {
                throw new ArgumentNullException(nameof(target));
            }
            this.msgHeaders = source;
            this.mimeHeaders = target;
        }

        #endregion

        #region public methods

        public void Convert()
        {
            ConvertHeaderAddressList(HeaderId.Bcc, msgHeaders.Bcc);
            ConvertHeaderAddressList(HeaderId.Cc, msgHeaders.Cc);
            ConvertHeaderString(HeaderId.ContentDescription, msgHeaders.ContentDescription);
            ConvertHeaderObject(HeaderId.ContentDisposition, msgHeaders.ContentDisposition);
            ConvertHeaderString(HeaderId.ContentId, msgHeaders.ContentId);
            ConvertHeaderObject(HeaderId.ContentTransferEncoding, msgHeaders.ContentTransferEncoding);
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

        #endregion

        #region private methods

        private void ConvertHeaderAddressList(
            HeaderId headerId,
            List<MsgReader.Mime.Header.RfcMailAddress> addresses)
        {
            if (mimeHeaders.Any(h => h.Id == headerId))
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

        private void ConvertHeaderString(HeaderId headerId, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }
            if (mimeHeaders.Any(h => h.Id == headerId))
            {
                return;
            }
            var header = new Header(headerId, value);
            mimeHeaders.Add(header);
        }

        private void ConvertHeaderObject(HeaderId headerId, object value)
        {
            if (value == null)
            {
                return;
            }
            ConvertHeaderString(headerId, value.ToString());
        }

        private void ConvertHeaderStringList(
            HeaderId headerId,
            List<string> values)
        {
            if (mimeHeaders.Any(h => h.Id == headerId))
            {
                return;
            }
            foreach (var value in values)
            {
                ConvertHeaderString(headerId, value);
            }
        }

        private void ConvertHeaderReceived()
        {
            if (msgHeaders.Received == null)
            {
                return;
            }
            foreach (var received in msgHeaders.Received)
            {
                ConvertHeaderString(HeaderId.Received, received.Raw);
            }
        }

        private void ConvertHeadersUnknown()
        {
            foreach (var key in msgHeaders.UnknownHeaders.AllKeys)
            {
                var value = msgHeaders.UnknownHeaders[key];
                HeaderId id;
                var keyString = key.Replace("-", string.Empty);
                if (Enum.TryParse(keyString, true, out id))
                {
                    ConvertHeaderString(id, value);
                }
                else
                {
                    mimeHeaders.Add(key, Encoding.ASCII, value);
                }
            }
        }

        #endregion
    }
}
