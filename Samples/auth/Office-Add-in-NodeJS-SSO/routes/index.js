// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file defines the home page request handling.

var express = require('express');
var router = express.Router();

/* GET home page. */
router.get('/', function(req, res, next) {
  res.render('index');
});

module.exports = router;
