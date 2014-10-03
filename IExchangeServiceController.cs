using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
namespace ExchangeWCFService
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IExchangeServiceController
    {


        [OperationContract]
        string SendEmail(CustomEmailMessage Message);



  

        // TODO: Add your service operations here
    }


    [DataContract]
    public class CustomEmailMessage {
        [DataMember]
        public string HtmlMessage { get; set; }
        [DataMember]
        public List<EmailAddress> TOList { get; set; }
        [DataMember]
        public string Subject { get; set; }
        [DataMember]
        public List<EmailAddress> CCList { get; set; }
    }



}
