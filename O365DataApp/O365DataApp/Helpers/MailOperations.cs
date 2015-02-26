using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365DataApp.Helpers
{
    static class MailOperations
    {
        public static async Task sendMail(OutlookServicesClient mailClient, List<Message> newMails)
        {
            foreach (var newMail in newMails)
            {
                await mailClient.Me.SendMailAsync(newMail, true);
            }
        }

        public static async Task<List<Message>> getMails(OutlookServicesClient mailClient)
        {
            List<Message> myMails = new List<Message>();
            var mailsResult = await mailClient.Me.Messages.ExecuteAsync();
            do
            {
                var mails = mailsResult.CurrentPage;
                foreach (var mail in mails)
                {
                    myMails.Add((Message)mail);
                }
                mailsResult = await mailsResult.GetNextPageAsync();
            } while (mailsResult != null);
            return myMails;
        }
    }
}
