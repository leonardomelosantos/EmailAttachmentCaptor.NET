using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EmailAttachmentCaptor.NET
{
    public static class EmailAttachmentCaptorHelper
    {
        public static List<MessagePart> GetEmailAttachments(EmailAttachmentCaptorSettings settings)
        {
            if (settings == null)
                throw new ArgumentNullException(nameof(settings));

            List<MessagePart> attachmentsResult = new List<MessagePart>();

            using (Pop3Client client = new Pop3Client())
            {
                ConnectToEmailServer(settings, client);
                ProcessInboxEmails(attachmentsResult, client, settings);
            }
            return attachmentsResult;
        }

        private static void ConnectToEmailServer(EmailAttachmentCaptorSettings settings, Pop3Client client)
        {
            client.Connect(settings.ConfigEmailHostname, settings.ConfigEmailPort, settings.UseSsl);
            client.Authenticate(settings.ConfigEmailUsername, settings.ConfigEmailPassword);
        }

        private static void ProcessInboxEmails(List<MessagePart> attachmentsResult, Pop3Client client, EmailAttachmentCaptorSettings settings)
        {
            List<string> allowedExtensions = settings.GetAllowedExtensions();
            List<string> emailsUids = client.GetMessageUids(); // Fetch all the current uids seen
            for (int i = 0; i < emailsUids.Count; i++)
            {
                Message unseenMessage = client.GetMessage(i + 1);
                List<MessagePart> emailAttachments = ProcessEmail(allowedExtensions, unseenMessage);
                attachmentsResult.AddRange(emailAttachments);
            }
        }

        private static List<MessagePart> ProcessEmail(List<string> allowedExtensions, Message unseenMessage)
        {
            List<MessagePart> attachmentsResult = new List<MessagePart>();
            List<MessagePart> attachments = unseenMessage.FindAllAttachments();
            foreach (MessagePart attachment in attachments)
            {
                if (attachment != null && IsExtensionAllowed(allowedExtensions, attachment))
                {
                    attachmentsResult.Add(attachment);
                }
            }
            return attachmentsResult;
        }

        private static bool IsExtensionAllowed(List<string> allowedExtensions, MessagePart attachment)
        {
            if (allowedExtensions == null || !allowedExtensions.Any())
                return true;

            return allowedExtensions.Contains(Path.GetExtension(attachment.FileName.ToLower()).ToLower());
        }
    }
}
