 const lowdb = require('lowdb');
 const FileSync = require('lowdb/adapters/FileSync');
 const adapter = new FileSync('./data/db.json');
 const db = lowdb(adapter);
 const { v4: uuidv4 } = require('uuid');

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
    try {
        const owner = req.authInfo['oid'];

        const todos = db.get('todos')
            .filter({ owner: owner })
            .value();

        res.status(200).send(todos);
    } catch (error) {
        next(error);
    }
}

exports.postTodo = (req, res, next) => {
    try {
        const todo = {
            description: req.body.description,
            id: uuidv4(),
            owner: req.authInfo['oid'] // oid is the only claim that should be used to uniquely identify a user in an Azure AD tenant
        };

        db.get('todos').push(todo).write();

        res.status(200).json(todo);
    } catch (error) {
        next(error);
    }
}

exports.deleteTodo = (req, res, next) => {
   res.send("test");
}
