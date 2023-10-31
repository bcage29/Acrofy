using System.Net.Http.Headers;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;

using Newtonsoft.Json.Linq;

namespace Acrofy.Bot;

public class TeamsMessageExtension : TeamsActivityHandler
{
    // Message Extension Code
    // Action.
    protected override Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action, CancellationToken cancellationToken)
    {
        switch (action.CommandId)
        {
            case "createCard":
                return Task.FromResult(CreateCardCommand(turnContext, action));
        }
        return Task.FromResult(new MessagingExtensionActionResponse());
    }

    private MessagingExtensionActionResponse CreateCardCommand(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionAction action)
    {
        // The user has chosen to create a card by choosing the 'Create Card' context menu command.
        var createCardData = ((JObject)action.Data).ToObject<CardResponse>();

        var card = new HeroCard
        {
            Title = createCardData.Title,
            Subtitle = createCardData.Subtitle,
            Text = createCardData.Text,
        };

        var attachments = new List<MessagingExtensionAttachment>();
        attachments.Add(new MessagingExtensionAttachment
        {
            Content = card,
            ContentType = HeroCard.ContentType,
            Preview = card.ToAttachment(),
        });

        return new MessagingExtensionActionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                AttachmentLayout = "list",
                Type = "result",
                Attachments = attachments,
            },
        };
    }

    // Search.
    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        var text = query?.Parameters?[0]?.Value as string ?? string.Empty;
        var packages = await FindPackages(text);

        // We take every row of the results and wrap them in cards wrapped in MessagingExtensionAttachment objects.
        // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.
        var attachments = packages.Select(package =>
        {
            var previewCard = new ThumbnailCard { Title = package.Item1, Tap = new CardAction { Type = "invoke", Value = package } };
            if (!string.IsNullOrEmpty(package.Item5))
            {
                previewCard.Images = new List<CardImage>() { new CardImage(package.Item5, "Icon") };
            }

            var attachment = new MessagingExtensionAttachment
            {
                ContentType = HeroCard.ContentType,
                Content = new HeroCard { Title = package.Item1 },
                Preview = previewCard.ToAttachment()
            };

            return attachment;
        }).ToList();

        // The list of MessagingExtensionAttachments must we wrapped in a MessagingExtensionResult wrapped in a MessagingExtensionResponse.
        return new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = attachments
            }
        };
    }
    protected override Task<MessagingExtensionResponse> OnTeamsMessagingExtensionSelectItemAsync(ITurnContext<IInvokeActivity> turnContext, JObject query, CancellationToken cancellationToken)
    {
        // The Preview card's Tap should have a Value property assigned, this will be returned to the bot in this event.
        var (packageId, version, description, projectUrl, iconUrl) = query.ToObject<(string, string, string, string, string)>();

        // We take every row of the results and wrap them in cards wrapped in in MessagingExtensionAttachment objects.
        // The Preview is optional, if it includes a Tap, that will trigger the OnTeamsMessagingExtensionSelectItemAsync event back on this bot.

        var card = new ThumbnailCard
        {
            Title = $"{packageId}, {version}",
            Subtitle = description,
            Buttons = new List<CardAction>
                    {
                        new CardAction { Type = ActionTypes.OpenUrl, Title = "Nuget Package", Value = $"https://www.nuget.org/packages/{packageId}" },
                        new CardAction { Type = ActionTypes.OpenUrl, Title = "Project", Value = projectUrl },
                    },
        };

        if (!string.IsNullOrEmpty(iconUrl))
        {
            card.Images = new List<CardImage>() { new CardImage(iconUrl, "Icon") };
        }

        var attachment = new MessagingExtensionAttachment
        {
            ContentType = ThumbnailCard.ContentType,
            Content = card,
        };

        return Task.FromResult(new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = new List<MessagingExtensionAttachment> { attachment }
            }
        });
    }

    // Generate a set of substrings to illustrate the idea of a set of results coming back from a query.
    private async Task<IEnumerable<(string, string, string, string, string)>> FindPackages(string text)
    {
        var obj = JObject.Parse(await (new HttpClient()).GetStringAsync($"https://azuresearch-usnc.nuget.org/query?q=id:{text}"));
        return obj["data"].Select(item => (item["id"].ToString(), item["version"].ToString(), item["description"].ToString(), item["projectUrl"]?.ToString(), item["iconUrl"]?.ToString()));
    }

    private async Task FindAcronyms(string text)
    {
        //var graphClient = new GraphServiceClient()
        text = text.ToUpper();
        var acronymListUrl = "https://graph.microsoft.com/v1.0/sites/fa047178-d741-4117-9bc4-dca6c91b2b77/lists/1e78f375-9192-4515-af6e-964866bde52a/items?expand=fields(select=Title,Definition,Description)";
        var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", "Your Oauth token");

        var response = await httpClient.GetAsync(acronymListUrl);
        
    }
    internal class CardResponse
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Text { get; set; }
    }
}

