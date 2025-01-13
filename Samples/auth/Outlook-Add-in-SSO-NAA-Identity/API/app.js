// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file is the main Node.js server file that defines the express middleware.

const express = require('express');
const morgan = require('morgan');
const cors = require('cors');
const apiRouter = require('./routes/index');

const app = express();

/**
 * Enable CORS middleware. In production, modify as to allow only designated origins and methods.
 * If you are using Azure App Service, we recommend removing the line below and configure CORS on the App Service itself.
 */
app.use(cors());

app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(morgan('dev'));

app.use('/api', apiRouter);

const port = process.env.PORT || 5000;

module.exports = app;