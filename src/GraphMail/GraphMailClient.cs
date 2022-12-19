using System;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System.Net.Http.Headers;
using System.Collections.Generic;

using System.Text;

namespace Azure.ApiWrapper.GraphMail;
public class AzureGraphMailClient
{
    GraphServiceClient graphClient;
    string mailFromUserID;

    public AzureGraphMailClient(string newMailFromUserID)
    {
        SetGraphClientWithCredential(
            new ChainedTokenCredential(
                new ManagedIdentityCredential(),
                new EnvironmentCredential()
            ),
            newMailFromUserID
        );
    }

    public AzureGraphMailClient(ChainedTokenCredential credential, string newMailFromUserID)
    {
        SetGraphClientWithCredential(credential, newMailFromUserID);
    }

    private void SetGraphClientWithCredential(ChainedTokenCredential credential, string newMailFromUserID)
    {
        var token = credential.GetToken(
            new Azure.Core.TokenRequestContext(new[] { "https://graph.microsoft.com/.default" })
        );
        var accessToken = token.Token;

        graphClient  = new GraphServiceClient(
            new DelegateAuthenticationProvider((requestMessage) =>
            {
                requestMessage
                .Headers
                .Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                return Task.CompletedTask;
            }   
        ));

        mailFromUserID = newMailFromUserID;
    }

    public void SendReportEmail(
        string header, 
        IEnumerable<string> recipientsString, 
        string body
    )
    {
        var recipients = new List<Recipient>();
        foreach(var r in recipientsString)
        {
            recipients.Add(new Recipient { EmailAddress = new EmailAddress { Address = r }});
        }

        var message = new Message
        {
            Subject = header,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = body
            },
            ToRecipients = recipients
        };

        graphClient
            .Users[mailFromUserID]
            .SendMail(message, true)
            .Request().PostAsync().Wait();
    }
}