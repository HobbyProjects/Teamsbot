// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;

namespace Microsoft.BotBuilderSamples.Controllers
{

    [Route("api/notify")]
    [ApiController]
    public class NotifyController : ControllerBase
    {
        private readonly IBotFrameworkHttpAdapter _adapter;
        private readonly string _appId;
        private readonly ConcurrentDictionary<string, Bot.Schema.ConversationReference> _conversationReferences;
        private readonly List<Activity> _adaptiveCardActivities;

        // This array contains the file location of our adaptive cards
        private readonly string _card = System.IO.Path.Combine(".", "Resources", "adaptive-cards.json");

        public NotifyController(IBotFrameworkHttpAdapter adapter,
            IConfiguration configuration,
            ConcurrentDictionary<string, Bot.Schema.ConversationReference> conversationReferences,
            List<Activity> adaptiveCardActivities)
        {
            _adapter = adapter;
            _conversationReferences = conversationReferences;
            _appId = configuration["MicrosoftAppId"];
            _adaptiveCardActivities = adaptiveCardActivities;

            // If the channel is the Emulator, and authentication is not in use,
            // the AppId will be null.  We generate a random AppId for this case only.
            // This is not required for production, since the AppId will have a value.
            if (string.IsNullOrEmpty(_appId))
            {
                _appId = System.Guid.NewGuid().ToString(); //if no AppId, use a random Guid
            }
        }

        public async Task<IActionResult> Get()
        {
            foreach (var conversationReference in _conversationReferences.Values)
            {
                await ((BotAdapter)_adapter).ContinueConversationAsync(_appId, conversationReference, BotCallbackCard, default(CancellationToken));
            }

            // Let the caller know proactive messages have been sent
            return new ContentResult()
            {
                Content = "<html><body><h1>Proactive messages have been sent.</h1></body></html>",
                ContentType = "text/html",
                StatusCode = (int)System.Net.HttpStatusCode.OK,
            };
        }

        private async Task BotCallbackCard(ITurnContext turnContext, CancellationToken cancellationToken)
        {
            var cardAttachment = CreateAdaptiveCardAttachment(_card);
            _adaptiveCardActivities.Add((Activity)MessageFactory.Attachment(cardAttachment));
            foreach(var activity in _adaptiveCardActivities)
            {
                await turnContext.SendActivityAsync(activity, cancellationToken);
            }
        }

        private static Attachment CreateAdaptiveCardAttachment(string filePath)
        {
            var adaptiveCardJson = System.IO.File.ReadAllText(filePath);
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = "application/vnd.microsoft.card.adaptive",
                Content = JsonConvert.DeserializeObject(adaptiveCardJson),
            };
            return adaptiveCardAttachment;
        }

    }
}
