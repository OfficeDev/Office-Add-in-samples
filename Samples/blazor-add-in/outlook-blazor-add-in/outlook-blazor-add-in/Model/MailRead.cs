/* Copyright (c) Maarten van Stam and Eric Legault. All rights reserved. Licensed under the MIT License. */

﻿using System.Text;

namespace BlazorAddIn.Model
{
    public class MailRead
    {
        public MailRead()
        {

        }

        public MailRead(string subject)
        {
            Subject = subject ?? "{UNKNOWN FROM CONSTRUCTOR}";
        }

        public string? AttachmentId { get; set; }
        public string? AttachmentName { get; set; }
        public string? Subject { get; set; }
        public string? AttachmentBase64Data { get; set; }

        public string? AttachmentType { get; set; }
        public bool Inline { get; set; } = false;

        public string? AttachmentContent { get; set; }
        public string? AttachmentImageURL { get; set; }

        public void DecodeBase64()
        {

            if (string.IsNullOrEmpty(AttachmentBase64Data) == false)
            {
                if (AttachmentType?.ToUpper() == "FILE" && Inline)
                {
                    AttachmentImageURL = string.Format("data:image/jpeg;base64,{0}", AttachmentBase64Data);
                }
                else
                {
                    AttachmentBase64Data = Encoding.UTF8.GetString(System.Convert.FromBase64String(AttachmentBase64Data));
                }
            }
        }
    }
}
