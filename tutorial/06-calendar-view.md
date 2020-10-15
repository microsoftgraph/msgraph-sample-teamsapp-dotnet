<!-- markdownlint-disable MD002 MD041 -->

In this section you will incorporate Microsoft Graph into the application. For this application, you will use the [Microsoft Graph Client Library for .NET](https://github.com/microsoftgraph/msgraph-sdk-dotnet) to make calls to Microsoft Graph.

# Get a calendar view

A calendar view is a set of events from the user's calendar that occur between two points of time. You'll use this to get the user's events for the current week.

1. Open **./Controllers/CalendarController.cs** and add the following function to the **CalendarController** class.

    :::code language="csharp" source="../demo/GraphTutorial/Controllers/CalendarController.cs" id="GetStartOfWeekSnippet":::

1. Replace the existing `Get` function with the following.

    :::code language="csharp" source="../demo/GraphTutorial/Controllers/CalendarController.cs" id="GetSnippet" highlight="2,14-56":::

    Review the changes. This new version of the function:

    - Returns `IEnumerable<Event>` instead of `string`.
    - Gets the user's mailbox settings using Microsoft Graph.
    - Uses the user's time zone to calculate the start and end of the current week.
    - Gets a calendar view
        - Uses the `.Header()` function to include a `Prefer: outlook.timezone` header, which causes the returned events to have their start and end times converted to the user's timezone.
        - Uses the `.Top()` function to request at most 50 events.
        - Uses the `.Select()` function to request just the fields used by the app.
        - Uses the `OrderBy()` function to sort the results by the start time.

1. Save your changes and restart the app. Refresh the tab in Microsoft Teams. The app displays a JSON listing of the events.

## Display the results

Now you can display the list of events in a more user friendly way.

1. Open **./Pages/Index.cshtml** and add the following functions inside the `<script>` tag.

    :::code language="javascript" source="../demo/GraphTutorial/Pages/Index.cshtml" id="RenderHelpersSnippet":::

1. Replace the existing `renderCalendar` function with the following.

    :::code language="javascript" source="../demo/GraphTutorial/Pages/Index.cshtml" id="RenderCalendarSnippet":::

1. Save your changes and restart the app. Refresh the tab in Microsoft Teams. The app displays events on the user's calendar.

    ![A screenshot of the app displaying the user's calendar](images/calendar-view.png)
