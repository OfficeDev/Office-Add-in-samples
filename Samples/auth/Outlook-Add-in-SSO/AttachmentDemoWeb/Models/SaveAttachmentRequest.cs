using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AttachmentDemoWeb.Models
{
    public class SaveAttachmentRequest
    {
        public string[] attachmentIds { get; set; }
        public string messageId { get; set; }
        public string outlookToken { get; set; }
        public string outlookRestUrl { get; set; }
        public string oneDriveToken { get; set; }

        public bool IsValid()
        {
            return attachmentIds != null && attachmentIds.Length > 0 &&
                !string.IsNullOrEmpty(messageId) &&
                !string.IsNullOrEmpty(outlookToken) &&
                !string.IsNullOrEmpty(outlookRestUrl) &&
                !string.IsNullOrEmpty(oneDriveToken);
        }
    }
}