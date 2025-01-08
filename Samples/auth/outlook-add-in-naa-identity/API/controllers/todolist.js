// const lowdb = require('lowdb');
// const FileSync = require('lowdb/adapters/FileSync');
// const adapter = new FileSync('./data/db.json');
// const db = lowdb(adapter);
// const { v4: uuidv4 } = require('uuid');

// const {
//   isAppOnlyToken,
//   hasRequiredDelegatedPermissions,
//   hasRequiredApplicationPermissions
// } = require('../auth/permissionUtils');

// const authConfig = require('../authConfig');

exports.getTodo = (req, res, next) => {
   res.send("test");
}

exports.getTodos = (req, res, next) => {
 res.send("test");
}

exports.postTodo = (req, res, next) => {
res.send("test");
}

exports.deleteTodo = (req, res, next) => {
   res.send("test");
}
