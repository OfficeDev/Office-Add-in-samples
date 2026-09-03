function getUserGists(user, callback) {
  const requestUrl = "https://api.github.com/users/" + encodeURIComponent(user) + "/gists";

  fetchJson(requestUrl, callback);
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function (gist) {
    const listItem = document.createElement("div");
    parent.appendChild(listItem);

    const radioItem = document.createElement("input");
    radioItem.classList.add("ms-ListItem", "is-selectable");
    radioItem.type = "radio";
    radioItem.name = "gists";
    radioItem.tabIndex = 0;
    radioItem.value = gist.id;
    listItem.appendChild(radioItem);

    const descPrimary = document.createElement("span");
    descPrimary.classList.add("ms-ListItem-primaryText");
    descPrimary.textContent = gist.description;
    listItem.appendChild(descPrimary);

    const descSecondary = document.createElement("span");
    descSecondary.classList.add("ms-ListItem-secondaryText");
    descSecondary.textContent = " - " + buildFileList(gist.files);
    listItem.appendChild(descSecondary);

    const updated = new Date(gist.updated_at);

    const descTertiary = document.createElement("span");
    descTertiary.classList.add("ms-ListItem-tertiaryText");
    descTertiary.textContent = " - Last updated " + updated.toLocaleString();
    listItem.appendChild(descTertiary);

    listItem.addEventListener("click", clickFunc);
  });
}

function buildFileList(files) {
  let fileList = "";

  for (let file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ", ";
      }

      fileList = fileList + files[file].filename + " (" + files[file].language + ")";
    }
  }

  return fileList;
}

function getGist(gistId, callback) {
  const requestUrl = "https://api.github.com/gists/" + encodeURIComponent(gistId);

  fetchJson(requestUrl, callback);
}

function fetchJson(url, callback) {
  fetch(url, { headers: { Accept: "application/vnd.github+json" } })
    .then(function (response) {
      if (!response.ok) {
        throw new Error("GitHub request failed: " + response.status + " " + response.statusText);
      }
      return response.json();
    })
    .then(function (data) {
      callback(data);
    })
    .catch(function (error) {
      callback(null, error);
    });
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (let filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      const file = gist.files[filename];
      if (!file.truncated) {
        switch (file.language) {
          case "HTML":
            // Insert as is.
            callback(file.content);
            break;
          case "Markdown":
            // Use GitHub's renderer so gist Markdown matches github.com.
            fetch("https://api.github.com/markdown", {
              method: "POST",
              headers: {
                Accept: "application/vnd.github+json",
                "Content-Type": "application/json",
              },
              body: JSON.stringify({ text: file.content, mode: "gfm" }),
            })
              .then(function (response) {
                if (!response.ok) {
                  throw new Error("GitHub Markdown request failed: " + response.status + " " + response.statusText);
                }
                return response.text();
              })
              .then(function (html) {
                callback(html);
              })
              .catch(function (error) {
                callback(null, error);
              });
            break;
          default:
            // Insert contents as a <code> block.
            const codeElement = document.createElement("code");
            codeElement.textContent = file.content;
            const preElement = document.createElement("pre");
            preElement.appendChild(codeElement);
            callback(preElement.outerHTML);
        }
        return;
      }
    }
  }
  callback(null, "No suitable file found in the gist");
}
