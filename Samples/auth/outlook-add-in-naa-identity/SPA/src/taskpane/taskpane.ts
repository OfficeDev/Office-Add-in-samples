/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { AccountManager, protectedResources } from "./authConfig";

// Use AccountManager to access msal-browser for authentication.
const accountManager = new AccountManager();

// Select DOM elements to work with.
const welcomeDiv = document.getElementById('welcome-div');
const tableDiv = document.getElementById('table-div');
const tableBody = document.getElementById('table-body-div');
const addTodoButton = document.getElementById('addTodoButton');
const textInput = document.getElementById('textInput');
const toDoListDiv = document.getElementById('groupDiv');
const todoListItems = document.getElementById('toDoListItems');

// The initialize function must be run each time a new page is loaded.
Office.onReady((info) => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  if (addTodoButton) {
    addTodoButton.addEventListener("click", addTodo);
  }

  // Initialize MSAL.
  accountManager.initialize();
});

async function addTodo(){
  const accessToken = await accountManager.ssoGetAccessToken(protectedResources.todolistApi.scopes.write);
  const endpoint = protectedResources.todolistApi.endpoint + "/todolist";
  const response = await callApi('POST',endpoint,accessToken,{description: textInput.value});
  const body = await response.text();
  console.log(body);
}

async function getTodoList (){
  let task = { description: textInput.innerText };
  // Specify minimum scopes for the token needed.
  const accessToken = await accountManager.ssoGetAccessToken(protectedResources.todolistApi.scopes.read);
  const endpoint = protectedResources.todolistApi.endpoint + "/todolist";
  const response = await callApi('GET',endpoint,accessToken);
  const body = await response.text();
  console.log(body);
//  handleToDoListActions(task, 'POST', protectedResources.todolistApi.endpoint);
}

function welcomeUser(username) {
  welcomeDiv.classList.remove('d-none');
  welcomeDiv.innerHTML = `Welcome ${username}!`;
}

function showToDoListItems(response) {
  todoListItems.replaceChildren();
  tableDiv.classList.add('d-none');
  toDoListDiv.classList.remove('d-none');
  if (!!response.length) {
      response.forEach((task) => {
          AddTaskToToDoList(task);
      });
  }
}

function AddTaskToToDoList(task) {
  let li = document.createElement('li');
  let button = document.createElement('button');
  button.innerHTML = 'Delete';
  button.classList.add('btn', 'btn-danger');
  button.addEventListener('click', () => {
      handleToDoListActions(task, 'DELETE', protectedResources.todolistApi.endpoint + `/${task.id}`);
  });
  li.classList.add('list-group-item', 'd-flex', 'justify-content-between', 'align-items-center');
  li.innerHTML = task.description;
  li.appendChild(button);
  todoListItems.appendChild(li);
}

/**
 *  Execute a fetch request with the given options
 * @param {string} method: GET, POST, PUT, DELETE
 * @param {String} endpoint: The endpoint to call
 * @param {Object} data: The data to send to the endpoint, if any
 * @returns response
 */
function callApi(method, endpoint, token, data = null) {
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

  return fetch(endpoint, options)
      .then((response) => {
          const contentType = response.headers.get("content-type");
          
          if (contentType && contentType.indexOf("application/json") !== -1) {
              return response.json();
          } else {
              return response;
          }
      });
}


/**
* Handles todolist actions
* @param {Object} task
* @param {string} method
* @param {string} endpoint
*/
export async function handleToDoListActions(task, method, endpoint) {
  let listData;
  
  try {
    let scopes = null;
      if (method === "DELETE") {
        scopes = protectedResources.todolistApi.scopes.write;
      }
      else {
        scopes = protectedResources.todolistApi.scopes.read;
      }
      const accessToken = await accountManager.ssoGetAccessToken(scopes);
      const data = await callApi(method, endpoint, accessToken, task);

      switch (method) {
          case 'POST':
              listData = JSON.parse(localStorage.getItem('todolist'));
              listData = [data, ...listData];
              localStorage.setItem('todolist', JSON.stringify(listData));
              AddTaskToToDoList(data);
              break;
          case 'DELETE':
              listData = JSON.parse(localStorage.getItem('todolist'));
              const index = listData.findIndex((todoItem) => todoItem.id === task.id);
              localStorage.setItem('todolist', JSON.stringify([...listData.splice(index, 1)]));
              showToDoListItems(listData);
              break;
          default:
              console.log('Unrecognized method.')
              break;
      }
  } catch (error) {
      console.error(error);
  }
}

/**
* Handles todolist action GET action.
*/
async function getToDos() {
  try {
      const accessToken = await accountManager.ssoGetAccessToken(protectedResources.todolistApi.scopes.read);

      const data = await callApi(
          'GET',
          protectedResources.todolistApi.endpoint,
          accessToken
      );

      if (data) {
          localStorage.setItem('todolist', JSON.stringify(data));
          showToDoListItems(data);
      }
  } catch (error) {
      console.error(error);
  }
}
