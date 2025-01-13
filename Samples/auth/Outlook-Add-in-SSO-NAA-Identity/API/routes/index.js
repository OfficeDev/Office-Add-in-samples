// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require('express');

const todolist = require('../controllers/todolist');
const authHelper = require('../server-helpers/validation-helper');


// initialize router
const router = express.Router();

router.get('/todolist', authHelper.validateJwt, todolist.getTodos);

router.post('/todolist', authHelper.validateJwt, todolist.postTodo);

router.delete('/todolist/:id', authHelper.validateJwt, todolist.deleteTodo);

module.exports = router;
