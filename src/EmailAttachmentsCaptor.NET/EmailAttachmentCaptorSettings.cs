using System;
using System.Collections.Generic;
using System.Text;

namespace EmailAttachmentCaptor.NET
{
    public class EmailAttachmentCaptorSettings
    {
        public List<string> Extensions { get; set; } = new List<string>();
        public string ConfigEmailHostname { get; set; } = string.Empty;
        public int ConfigEmailPort { get; set; }
        public bool UseSsl { get; set; } = false;
        public string ConfigEmailUsername { get; set; } = string.Empty;
        public string ConfigEmailPassword { get; set; } = string.Empty;

        public List<string> GetAllowedExtensions()
        {
            if (Extensions == null)
            {
                return new List<string>();
            }
            return Extensions;
        }
    }
}
