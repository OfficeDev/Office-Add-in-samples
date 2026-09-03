(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded.
  Office.onReady(function (reason) {
    function initializeDialog() {
      if (window.location.search) {
        // Check if warning should be displayed.
        const warn = getParameterByName("warn");
        if (warn) {
          document.querySelector(".not-configured-warning").style.display = "block";
        } else {
          // See if the config values were passed.
          // If so, pre-populate the values.
          const user = getParameterByName("gitHubUserName");
          const gistId = getParameterByName("defaultGistId");

          document.getElementById("github-user").value = user;
          loadGists(user, function (success) {
            if (success) {
              document.querySelectorAll(".ms-ListItem").forEach(function (item) {
                item.classList.remove("is-selected");
                if (item.value === gistId) {
                  item.classList.add("is-selected");
                  item.checked = true;
                }
              });
              document.getElementById("settings-done").disabled = false;
            }
          });
        }
      }

      // When the GitHub username changes,
      // try to load gists.
      document.getElementById("github-user").addEventListener("change", function () {
        document.getElementById("gist-list").textContent = "";
        const ghUser = document.getElementById("github-user").value;
        if (ghUser.length > 0) {
          loadGists(ghUser);
        }
      });

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      document.getElementById("settings-done").addEventListener("click", function () {
        const settings = {};

        settings.gitHubUserName = document.getElementById("github-user").value;

        const selectedGist = document.querySelector(".ms-ListItem.is-selected");
        if (selectedGist) {
          settings.defaultGistId = selectedGist.value;

          sendMessage(JSON.stringify(settings));
        }
      });
    }

    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", initializeDialog);
    } else {
      initializeDialog();
    }
  });

  // Load gists for the user using the GitHub API
  // and build the list.
  function loadGists(user, callback) {
    getUserGists(user, function (gists, error) {
      if (error) {
        document.querySelector(".gist-list-container").style.display = "none";
        document.getElementById("error-text").textContent = JSON.stringify(error, null, 2);
        document.querySelector(".error-display").style.display = "block";
        if (callback) callback(false);
      } else {
        document.querySelector(".error-display").style.display = "none";
        buildGistList(document.getElementById("gist-list"), gists, onGistSelected);
        document.querySelector(".gist-list-container").style.display = "block";
        if (callback) callback(true);
      }
    });
  }

  function onGistSelected() {
    document.querySelectorAll(".ms-ListItem").forEach(function (item) {
      item.classList.remove("is-selected");
      item.checked = false;
    });
    const selectedItem = this.querySelector(".ms-ListItem");
    selectedItem.classList.add("is-selected");
    selectedItem.checked = true;
    document.querySelector(".not-configured-warning").style.display = "none";
    document.getElementById("settings-done").disabled = false;
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    return new URL(url || window.location.href).searchParams.get(name);
  }
})();
