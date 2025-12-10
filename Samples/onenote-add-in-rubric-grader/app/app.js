// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const app = (() => {  // jshint ignore:line
  'use strict';

  const self = {};

  // Common initialization function (to be called from each page).
  self.initialize = () => {
    const notificationHtml = 
      '<div id="notification-message">' +
      '<div class="padding">' +
      '<div id="notification-message-close"></div>' +
      '<div id="notification-message-header"></div>' +
      '<div id="notification-message-body"></div>' +
      '</div>' +
      '</div>';
      
    document.body.insertAdjacentHTML('beforeend', notificationHtml);

    document.getElementById('notification-message-close').addEventListener('click', () => {
      document.getElementById('notification-message').style.display = 'none';
    });

    // After initialization, expose a common notification function.
    self.showNotification = (header, text) => {
      document.getElementById('notification-message-header').textContent = header;
      document.getElementById('notification-message-body').textContent = text;
      const notificationElement = document.getElementById('notification-message');
      notificationElement.style.display = 'block';
      
      // Simple slide down effect replacement.
      notificationElement.style.opacity = '0';
      notificationElement.style.transform = 'translateY(-20px)';
      notificationElement.style.transition = 'opacity 0.3s ease, transform 0.3s ease';
      
      setTimeout(() => {
        notificationElement.style.opacity = '1';
        notificationElement.style.transform = 'translateY(0)';
      }, 10);
    };
  };

  return self;
})();
