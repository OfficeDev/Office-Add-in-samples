// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const app = (function(){  // jshint ignore:line
  'use strict';

  let self = {};

  // Common initialization function (to be called from each page)
  self.initialize = function(){
    jQuery('body').append(
      '<div id="notification-message">' +
      '<div class="padding">' +
      '<div id="notification-message-close"></div>' +
      '<div id="notification-message-header"></div>' +
      '<div id="notification-message-body"></div>' +
      '</div>' +
      '</div>');

    jQuery('#notification-message-close').click(function(){
      jQuery('#notification-message').hide();
    });

    // After initialization, expose a common notification function
    self.showNotification = function(header, text){
      jQuery('#notification-message-header').text(header);
      jQuery('#notification-message-body').text(text);
      jQuery('#notification-message').slideDown('fast');
    };
  };

  return self;
})();
