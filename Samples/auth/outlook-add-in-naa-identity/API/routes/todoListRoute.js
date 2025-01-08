// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require('express');
var router = express.Router();
const authHelper = require('../server-helpers/obo-auth-helper');
const jwt = require('jsonwebtoken');
const { v4: uuidv4 } = require('uuid');
const authConfig = require('../authConfig');

router.get(
    '/todolist', 
    async function (req, res) {
        res.status(200).send("test");
    });  

    module.exports = router;