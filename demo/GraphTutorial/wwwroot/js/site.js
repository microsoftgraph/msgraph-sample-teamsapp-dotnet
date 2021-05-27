// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// <ThemeSupportSnippet>
(function () {
  // Support Teams themes
  microsoftTeams.initialize();

  // On load, match the current theme
  microsoftTeams.getContext((context) => {
    if(context.theme !== 'default') {
      // For Dark and High contrast, set text to white
      document.body.style.color = '#fff';
      document.body.style.setProperty('--border-style', 'solid');
    }
  });

  // Register event listener for theme change
  microsoftTeams.registerOnThemeChangeHandler((theme)=> {
    if(theme !== 'default') {
      document.body.style.color = '#fff';
      document.body.style.setProperty('--border-style', 'solid');
    } else {
      // For default theme, remove inline style
      document.body.style.color = '';
      document.body.style.setProperty('--border-style', 'none');
    }
  });
})();
// </ThemeSupportSnippet>
