using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsgKit.Enums
{
    /// <summary>
    /// Extension conversion methods for enums.
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Converts <see cref="MsgReader.Outlook.MessageImportance"/> to <see cref="MimeKit.MessageImportance"/>./>
        /// </summary>
        /// <param name="importance">The MSG importance.</param>
        /// <returns>The mime importance.</returns>
        public static MimeKit.MessageImportance Map(this MsgReader.Outlook.MessageImportance? importance)
        {
            if (importance == null)
            {
                return MimeKit.MessageImportance.Normal; 
            }
            switch (importance)
            {
                case MsgReader.Outlook.MessageImportance.Low: return MimeKit.MessageImportance.Low;
                case MsgReader.Outlook.MessageImportance.Normal: return MimeKit.MessageImportance.Normal;
                case MsgReader.Outlook.MessageImportance.High: return MimeKit.MessageImportance.High;
                default: return MimeKit.MessageImportance.Normal; 
            }
        }

        /// <summary>
        /// Converts <see cref="MsgReader.Outlook.Storage.Message"/> to <see cref="MimeKit.MessageImportance"/>./>
        /// </summary>
        /// <param name="priority">The MSG importance.</param>
        /// <returns>The mime importance.</returns>
        public static MimeKit.MessagePriority MapToPriority(this string priority)
        {
            if (priority == null)
            {
                return MimeKit.MessagePriority.Normal;
            }
            switch (priority)
            {
                case "0": return MimeKit.MessagePriority.NonUrgent;
                case "1": return MimeKit.MessagePriority.Normal;
                case "2": return MimeKit.MessagePriority.Urgent;
                default: return MimeKit.MessagePriority.Normal;
            }
        }
    }
}
