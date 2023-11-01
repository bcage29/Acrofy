using System.Net.Http.Headers;
using AdaptiveCards;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Newtonsoft.Json;

namespace Acrofy.Bot;

public class TeamsMessageExtension : TeamsActivityHandler
{
    // Process create acronym
    protected override async Task<AdaptiveCardInvokeResponse> OnAdaptiveCardInvokeAsync(ITurnContext<IInvokeActivity> turnContext, AdaptiveCardInvokeValue adaptiveCardInvokeValue, CancellationToken cancellationToken)
    {
        if (turnContext.Activity.Name == "adaptiveCard/action")
        {
            var paths = new[] { ".", "adaptiveCards", "adaptiveCardSuccess.json" };
            var adaptiveCardJson = File.ReadAllText(Path.Combine(paths));
            var adaptiveCardResponse = new AdaptiveCardInvokeResponse()
            {
                StatusCode = 200,
                Type = AdaptiveCard.ContentType,
                Value = JsonConvert.DeserializeObject(adaptiveCardJson)
            };

            return adaptiveCardResponse;
        }

        return null;

    }

    // Search.
    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        // get the query parameter
        var text = query?.Parameters?[0]?.Value as string ?? string.Empty;
        text = text.ToUpper();
        // if the query is empty, set the acronym object to null
        var acronym = string.IsNullOrEmpty(text) ? null : await FindAcronyms(text);
        MessagingExtensionAttachment attachment;
        if (acronym == null)
        {
            var paths = new[] { ".", "adaptiveCards", "newAcronym.json" };
            var filePath = Path.Combine(paths);
            var adaptiveCard = FetchAdaptiveCard(filePath);

            var previewCard = new ThumbnailCard
            {
                Title = $"{text}: Request a new acronym",
                Text = $"Fill out this form to request a new acronym: {text}"
            };

            attachment = new MessagingExtensionAttachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard.Content,
                Preview = previewCard.ToAttachment(),
            };
        }
        else
        {
            var previewCard = new ThumbnailCard
            {
                Title = acronym.Title,
                Text = acronym.Definition
            };
            attachment = new MessagingExtensionAttachment
            {
                ContentType = HeroCard.ContentType,
                Content = new HeroCard { Title = acronym.Title, Subtitle = acronym.Definition, Text = acronym.Description },
                Preview = previewCard.ToAttachment()
            };
        }

        // The list of MessagingExtensionAttachments must we wrapped in a MessagingExtensionResult wrapped in a MessagingExtensionResponse.
        return new MessagingExtensionResponse
        {
            ComposeExtension = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = "list",
                Attachments = new List<MessagingExtensionAttachment> { attachment }
            }
        };
    }

    private async Task<Acronym?> FindAcronyms(string text)
    {
        try
        {
            // This is sample test code for a POC
            // In a real scenario we would use a Graph Service Client
            // and not be creating new HttpClients
            var acronymListUrl = $"https://graph.microsoft.com/v1.0/sites/<siteId>/lists/<listId>/items?$filter=fields/Title eq '{text}'&expand=fields(select=Title,Definition,Description)";
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", "");
            httpClient.DefaultRequestHeaders.Add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");

            var response = await httpClient.GetAsync(acronymListUrl);
            response.EnsureSuccessStatusCode();
            var jsonString = await response.Content.ReadAsStringAsync();
            var data = JsonConvert.DeserializeObject<SpList>(jsonString);

            var item = data.value.FirstOrDefault();
            if (item != null)
            {
                return item.fields;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
        
        return null;
    }

    private Attachment FetchAdaptiveCard(string filePath)
    {
        var adaptiveCardJson = File.ReadAllText(filePath);
        var adaptiveCardAttachment = new Attachment
        {
            ContentType = AdaptiveCard.ContentType,
            Content = JsonConvert.DeserializeObject(adaptiveCardJson)
        };
        return adaptiveCardAttachment;
    }

    internal class SpList
    {
        public List<SpListItems> value;
    }

    internal class SpListItems
    {
        public Acronym fields;
    }

    internal class Acronym
    {
        public string Title { get; set; }
        public string Definition { get; set; }
        public string Description { get; set; }
    }

    internal class CardResponse
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Text { get; set; }
    }
}

