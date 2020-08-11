using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Extensions.Configuration;

namespace TeamsConversationBot.Dialogs
{
    public class CreateTeamsMeetingDialog : ComponentDialog
    {
        private readonly string _connectionName;
        public CreateTeamsMeetingDialog (IConfiguration configuration) : base(nameof (CreateTeamsMeetingDialog))
        {
            _connectionName = configuration.GetSection("ConnectionName")?.Value;

            var steps = new WaterfallStep[] {
                CreateMeetingAsync
            };

            AddDialog(new WaterfallDialog(nameof(CreateTeamsMeetingDialog), steps));
        }

        protected override async Task<DialogTurnResult> OnBeginDialogAsync(DialogContext innerDc, object options, CancellationToken cancellationToken = default(CancellationToken))
        {
            return await base.OnBeginDialogAsync(innerDc, options, cancellationToken);
        }

        private async Task<DialogTurnResult> CreateMeetingAsync(WaterfallStepContext context, CancellationToken cancellationToken)
        {
            var adapter = context.Context.Adapter as IUserTokenProvider;
            var token = await adapter.GetUserTokenAsync(context.Context, _connectionName, null, cancellationToken);

            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri("https://graph.microsoft.com/beta/");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

            MeetingInfo meetingInfo = new MeetingInfo
            {
                startDateTime = "2020-09-12T14:30:34.2444915-07:00",
                endDateTime = "2020-09-12T15:00:34.2464912-07:00",
                subject = "sample subject"
            };

            HttpResponseMessage response = await client.PostAsJsonAsync("me/onlineMeetings", meetingInfo);
            response.EnsureSuccessStatusCode();
            string result = await  response.Content.ReadAsStringAsync();
            string meetingUrl = result.Substring((result.IndexOf("joinUrl\":\"") + "joinUrl\":\"".Length),
                (result.IndexOf(",\"joinWebUrl") - (result.IndexOf("joinUrl\":\"") + "joinUrl\":\"".Length + 1)));

            await context.Context.SendActivityAsync($"Here is your meeting url: {meetingUrl}");

            return await context.EndDialogAsync(null, cancellationToken);
        }
    }

    public class MeetingInfo
    {
        public string startDateTime { get; set; }
        public string endDateTime { get; set; }
        public string subject { get; set; }
    }
}
