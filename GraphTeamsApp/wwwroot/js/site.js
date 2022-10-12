// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

$(document).ready(async function () {
  await microsoftTeams.app.initialize();

  const context = await microsoftTeams.app.getContext();

  // On load, match the current theme
  if (context.app.theme !== 'default') {
    // For Dark and High contrast, set text to white
    document.body.style.color = '#fff';
    document.body.style.setProperty('--border-style', 'solid');
  }

  // Register event listener for theme change
  microsoftTeams.app.registerOnThemeChangeHandler((theme) => {
    if(theme !== 'default') {
      document.body.style.color = '#fff';
      document.body.style.setProperty('--border-style', 'solid');
    } else {
      // For default theme, remove inline style
      document.body.style.color = '';
      document.body.style.setProperty('--border-style', 'none');
    }
  });
});
