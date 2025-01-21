/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { AccountManager, protectedResources } from "./authConfig";

// Use AccountManager to access msal-browser for authentication.
const accountManager = new AccountManager();

// Select DOM elements to work with.
const addTodoButton = document.getElementById('addTodoButton');
const getTodoListbutton = document.getElementById('getTodoListButton');
const textInput = document.getElementById('textInput');
const todoListUI = document.getElementById('toDoListItems');

// The initialize function must be run each time a new page is loaded.
Office.onReady(async (info) => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  // Add event listeners for button clicks.
  if (addTodoButton) {
    addTodoButton.addEventListener("click", addTodo);
  }
  if (getTodoListbutton) {
    getTodoListbutton.addEventListener("click", getTodoList);
  }

  // Initialize Microsoft authentication library (MSAL).
  await accountManager.initialize();
});

/**
 * Adds a new todo item to the user's todo list.
 */
async function addTodo() {
  logMessage(null); // Clear messages.
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(protectedResources.todolistApi.scopes.write);
  const endpoint = protectedResources.todolistApi.endpoint + "/todolist";
  callApi('POST', endpoint, accessToken, { description: textInput.value });
}

/**
 * Gets the todo list for the current user.
 */
async function getTodoList() {
  logMessage(null); // Clear messages.
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(protectedResources.todolistApi.scopes.read);
  const endpoint = protectedResources.todolistApi.endpoint + "/todolist";
  const response = await callApi('GET', endpoint, accessToken);
  if (response) {
    showToDoListItems(response);
  }
}

/**
 * Delets a todo item from the user's todo list.
 * 
 * @param id The id of the todo item to delete.
 */
async function deleteTodo(id: string) {
  logMessage(null); // Clear messages.
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(protectedResources.todolistApi.scopes.write);
  const endpoint = protectedResources.todolistApi.endpoint + "/todolist" + `/${id}`;
  await callApi('DELETE', endpoint, accessToken);
  getTodoList();
}

/**
 * Shows a list of todo items on the task pane.
 * 
 * @param todoListItems An array of todo list items.
 */
function showToDoListItems(todoListItems) {
  todoListUI.replaceChildren();
  if (!!todoListItems.length) {
    todoListItems.forEach((task) => {
      AddTaskToToDoList(task);
    });
  } else {
    // Display that the todo list is empty.
    logMessage("Todo list is empty.");
  }
}

/**
 * Adds a new task to the todo item list on the task pane.
 * 
 * @param task The todo item to add.
 */
function AddTaskToToDoList(task) {
  let li = document.createElement('li');
  let button = document.createElement('button');
  button.innerHTML = 'Delete';
  button.classList.add('btn', 'btn-danger');
  button.addEventListener('click', () => {
    deleteTodo(task.id);
  });
  li.classList.add('list-group-item', 'd-flex', 'justify-content-between', 'align-items-center');
  li.innerHTML = task.description;
  li.appendChild(button);
  todoListUI.appendChild(li);
}

/**
 *  Execute a fetch request with the given options
 * @param {string} method: GET, POST, PUT, DELETE
 * @param {String} endpoint: The endpoint to call
 * @param {Object} data: The data to send to the endpoint, if any
 * @returns response
 */


/**
 * Execute a fetch request with the given options.
 * @param {string} method: GET, POST, PUT, DELETE.
 * @param {String} endpoint: The endpoint to call.
 * @param {String} token: The access token for the API.
 * @param {Object} data: The data to send to the endpoint, if any.
 * @returns The JSON response if there is one, otherwise the response object.
 */
async function callApi(method: string, endpoint: string, token: string, data: object = null): Promise<any> {
  const headers = new Headers();
  const bearer = `Bearer ${token}`;
  headers.append('Authorization', bearer);

  if (data) {
    headers.append('Content-Type', 'application/json');
  }

  const options = {
    method: method,
    headers: headers,
    body: data ? JSON.stringify(data) : null,
  };
  try {
    const response = await fetch(endpoint, options);
    if (response.ok) {
      const contentType = response.headers.get("content-type");
      if (contentType && contentType.indexOf("application/json") !== -1) {
        return response.json();
      } else {
        return response;
      }
    } else {
      // Get message info from the body of the response.
      const message = await response.text();
      logMessage(`HTTP Error: ${response.status} with message: ${message}`);
      return null;
    }
  } catch (error) {
    logMessage(error.message);
  }
}

/**
 * Shows a message on the task pane such as status updates.
 * @param message The message to show.
 */
function logMessage(message) {
  const messageLabel = document.getElementById('messages');
  if (message) {
    messageLabel.value = message;
    messageLabel.style = "visibility:visible;"
  } else {
    messageLabel.value = "";
    messageLabel.style = "visibility:hidden;"
  }
}