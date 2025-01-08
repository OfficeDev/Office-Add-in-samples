// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const express = require('express');

const todolist = require('../controllers/todolist');
const authHelper = require('../server-helpers/obo-auth-helper');


// initialize router
const router = express.Router();

router.get('/todolist', authHelper.validateJwt, todolist.getTodos);

router.get('/todolist/:id', todolist.getTodo);

router.post('/todolist', todolist.postTodo);

router.delete('/todolist/:id', todolist.deleteTodo);

module.exports = router;
