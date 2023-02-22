// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file is the main Node.js server file that defines the express middleware.

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}
var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');

var getUserProfileRoute = require('./routes/getUserProfile');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
//app.set('view engine', 'pug');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());

/* Turn off caching when developing */
if (process.env.NODE_ENV !== 'production') {
  app.use(express.static(path.join(__dirname, 'public'),
                        { etag: false }));

  app.use(function (req, res, next) {
    res.set({
      "Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,
      "Cache-Control": "private, no-cache, no-store, must-revalidate",
      "Expires": "-1",
      "Pragma": "no-cache"
    });
    next()
  });
} else {
  // In production mode, let static files be cached.
  app.use(express.static(path.join(__dirname, 'public')));
  app.use(function (req, res, next) {
    res.set({
      "Content-Security-Policy": "script-src https://appsforoffice.microsoft.com https://ajax.aspnetcdn.com https://alcdn.msauth.net " +  process.env.SERVER_SOURCE,
    });
    next()
  });;
}

app.get('/getUserProfile', getUserProfileRoute);

// Catch 404 and forward to error handler
app.use(function(req, res, next) {
  next(createError(404));
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
