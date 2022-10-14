// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace GraphTeamsApp.Models
{
    public class NewEvent
    {
        public string Subject { get; set; } = string.Empty;
        public string Attendees { get; set; } = string.Empty;
        public string Start { get; set; } = string.Empty;
        public string End { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
    }
}
