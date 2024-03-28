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
        public static int CapturarArquivosAnexosEmails()
        {
            int qtdNovosAnexosBaixados = 0;

            string[] extensions = { ".pdf", ".txt" };

            using (Pop3Client client = new Pop3Client())
            {
                var configEmailHostname = "";
                var configEmailPort = 0;
                bool useSsl = false;
                string configEmailUsername = "";
                string configEmailPassword = "";

                client.Connect(configEmailHostname, configEmailPort, useSsl); // Connect to the server
                client.Authenticate(configEmailUsername, configEmailPassword); // Authenticate ourselves towards the server

                List<string> uids = client.GetMessageUids(); // Fetch all the current uids seen

                for (int i = 0; i < uids.Count; i++)
                {
                    long currentMessageUidOnServer = long.Parse(uids[i]);

                    Message unseenMessage = client.GetMessage(i + 1);
                    DateTime dataEmail = unseenMessage.Headers.DateSent;

                    var anexos = unseenMessage.FindAllAttachments();
                    foreach (var anexo in anexos)
                    {
                        if (extensions.Contains(Path.GetExtension(anexo.FileName).ToLower()))
                        {
                            string nomeArquivo = anexo.FileName;
                            byte[] conteudoArquivo = anexo.Body;
                            string mineType = anexo.ContentType.ToString();

                            qtdNovosAnexosBaixados++;
                        }
                    }
                }
            }

            return qtdNovosAnexosBaixados;
        }
    }
}
