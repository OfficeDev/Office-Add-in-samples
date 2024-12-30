const express = require('express');

const todolist = require('../controllers/todolist');

// initialize router
const router = express.Router();

router.get('/todolist', todolist.getTodos);

router.get('/todolist/:id', todolist.getTodo);

router.post('/todolist', todolist.postTodo);

router.delete('/todolist/:id', todolist.deleteTodo);

module.exports = router;