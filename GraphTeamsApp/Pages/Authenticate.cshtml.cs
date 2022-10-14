// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using Microsoft.AspNetCore.Mvc.RazorPages;

namespace GraphTeamsApp.Pages
{
    public class AuthenticateModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        public string ApplicationId { get; private set; }
        public string State { get; private set; }
        public string Nonce { get; private set; }

        public AuthenticateModel(
            IConfiguration configuration,
            ILogger<IndexModel> logger)
        {
            _logger = logger;

            // Read the application ID from the
            // configuration. This is used to build
            // the authorization URL for the consent prompt
            ApplicationId = configuration
                .GetSection("AzureAd")
                .GetValue<string>("ClientId");

            // Generate a GUID for state and nonce
            State = System.Guid.NewGuid().ToString();
            Nonce = System.Guid.NewGuid().ToString();
        }
    }
}
