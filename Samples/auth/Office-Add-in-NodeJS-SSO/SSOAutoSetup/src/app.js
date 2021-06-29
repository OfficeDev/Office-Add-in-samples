/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}
var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');

var getGraphData = require('./public/javascripts/msgraph-helper');

var indexRouter = require('./routes/index');
var authRouter = require('./routes/authRoute');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'pug');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());

/* Turn off caching when developing */
if (process.env.NODE_ENV !== 'production') {
  app.use(express.static(path.join(__dirname, 'public'),
                        { etag: false }));

  app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next()
  });
} else {
  // In production mode, let static files be cached.
  app.use(express.static(path.join(__dirname, 'public')));
}

app.use('/home/index', indexRouter);
app.use('/auth', authRouter);

app.get('/dialog.html', (async (req, res) => {
  return res.sendfile('dialog.html');
}));

app.get('/getuserdata', async function(req, res, next) {
  const graphToken = req.get('access_token');

  // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
  // and only the top 10 folder or file names.
  // Note that the last parameter, for queryParamsSegment, is hardcoded. If you reuse this code in
  // a production add-in and any part of queryParamsSegment comes from user input, be sure that it is
  // sanitized so that it cannot be used in a Response header injection attack. 
  const graphData = await getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=10");

  // If Microsoft Graph returns an error, such as invalid or expired token,
  // there will be a code property in the returned object set to a HTTP status (e.g. 401).
  // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
  if (graphData.code) {
      next(createError(graphData.code, "Microsoft Graph error " + JSON.stringify(graphData)));
  }
  else 
  {
    // MS Graph data includes OData metadata and eTags that we don't need.
    // Send only what is actually needed to the client: the item names.
    const itemNames = [];
    const oneDriveItems = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }

    res.send(itemNames)
  }
});


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
