(function () {
  "use strict";

  let config;
  let settingsDialog;

  Office.onReady(function () {
    function initializeTaskPane() {
      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        document.getElementById("not-configured").style.display = "";
      }

      // When insert button is selected, build the content
      // and insert into the body.
      document.getElementById("insert-button").addEventListener("click", function () {
        const selectedGist = document.querySelector(".ms-ListItem.is-selected");
        const gistId = selectedGist && selectedGist.value;
        getGist(gistId, function (gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(
                  content,
                  { coercionType: Office.CoercionType.Html },
                  function (result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError("Could not insert gist: " + result.error.message);
                    }
                  }
                );
              } else {
                showError("Could not create insertable content: " + error);
              }
            });
          } else {
            showError("Could not retrieve gist: " + error);
          }
        });
      });

      // When the settings icon is selected, open the settings dialog.
      document.getElementById("settings-icon").addEventListener("click", function () {
        // Display settings dialog.
        const url = new URL("dialog.html", window.location.href);
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url.searchParams.set("gitHubUserName", config.gitHubUserName);
          url.searchParams.set("defaultGistId", config.defaultGistId);
        }

        const dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url.toString(), dialogOptions, function (result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      });
    }

    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initializeTaskPane);
    } else {
      initializeTaskPane();
    }
  });

  function loadGists(user) {
    document.getElementById("error-display").style.display = "none";
    document.getElementById("not-configured").style.display = "none";
    document.getElementById("gist-list-container").style.display = "";

    getUserGists(user, function (gists, error) {
      if (error) {
      } else {
        const gistList = document.getElementById("gist-list");
        gistList.textContent = "";
        buildGistList(gistList, gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    document.getElementById("insert-button").disabled = false;
    document.querySelectorAll(".ms-ListItem").forEach(function (item) {
      item.classList.remove("is-selected");
      item.checked = false;
    });
    const selectedItem = this.querySelector(".ms-ListItem");
    selectedItem.classList.add("is-selected");
    selectedItem.checked = true;
  }

  function showError(error) {
    document.getElementById("not-configured").style.display = "none";
    document.getElementById("gist-list-container").style.display = "none";
    const errorDisplay = document.getElementById("error-display");
    errorDisplay.textContent = error;
    errorDisplay.style.display = "";
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function (result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
