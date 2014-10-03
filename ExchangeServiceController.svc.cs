using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace ExchangeWCFService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class ExchangeServiceController : IExchangeServiceController
    {

        public string SendEmail(CustomEmailMessage Message)
        {

            var Content = "";
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

                ExchangeService Exservice = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                Exservice.Credentials = new WebCredentials("username", "password", "domain");

              
                Exservice.Url = new System.Uri("https://mail.domainname/EWS/Exchange.asmx");



                EmailMessage email = new EmailMessage(Exservice);



                if (Message.TOList.Count() > 0)
                {

                    foreach (var item in Message.TOList)
                    {
                        email.ToRecipients.Add(item.ToString());
                    }

                }


                email.Subject = Message.Subject;
                email.Body = new MessageBody(BodyType.HTML, Message.HtmlMessage);
                if (Message.CCList.Count() > 0)
                {

                    foreach (var item in Message.CCList)
                    {
                        email.CcRecipients.Add(item.ToString());
                    }

                }

                email.Send();


                Content = "SUCCESS | EMAIL SENDED";
            }
            catch (Exception ex)
            {

                Content = "ERROR | " + ex.Message;
            }



            return Content;

        }


        private static bool CertificateValidationCallBack(
object sender,
System.Security.Cryptography.X509Certificates.X509Certificate certificate,
System.Security.Cryptography.X509Certificates.X509Chain chain,
System.Net.Security.SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (System.Security.Cryptography.X509Certificates.X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                            continue;
                        }
                        else
                        {
                            if (status.Status != System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }
            else
            {
                // In all other cases, return false.
                return false;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

    }
}
