const lowdb = require('lowdb');
const FileSync = require('lowdb/adapters/FileSync');
const adapter = new FileSync('./data/db.json');
const db = lowdb(adapter);
const { v4: uuidv4 } = require('uuid');

const authConfig = require('../server-helpers/authConfig');

exports.getTodos = (req, res, next) => {
    // Check that caller has the delegated todolist.read permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.read)) {

        try {
            const owner = req.authInfo['oid'];

            const todos = db.get('todos')
                .filter({ owner: owner })
                .value();

            res.status(200).send(todos);
        } catch (error) {
            next(error);
        }
    } else {
        next(new Error('User does not have the required permissions.'));
    }
}

exports.postTodo = (req, res, next) => {
    // Check that caller has the delegated todolist.readwrite permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.readWrite)) {

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
    } else (
        next(new Error('User does not have the required permissions.'))
    )
}

exports.deleteTodo = (req, res, next) => {
    // Check that caller has the delegated todolist.readwrite permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.readWrite)) {
        try {
            const id = req.params.id;
            const owner = req.authInfo['oid'];

            db.get('todos')
                .remove({ owner: owner, id: id })
                .write();

            res.status(200).json({ message: "success" });
        } catch (error) {
            next(error);
        }
    } else {
        next(new Error('User does not have the required permissions'))
    }
}

/**
 * Ensures that the access token has the specified delegated permissions.
 * @param {Object} accessTokenPayload: Parsed access token payload
 * @param {Array} requiredPermission: list of required permissions
 * @returns {boolean}
 */
const hasRequiredDelegatedPermissions = (accessTokenPayload, requiredPermission) => {
    const normalizedRequiredPermissions = requiredPermission.map(permission => permission.toUpperCase());

    if (accessTokenPayload.hasOwnProperty('scp') && accessTokenPayload.scp.split(' ')
        .some(claim => normalizedRequiredPermissions.includes(claim.toUpperCase()))) {
        return true;
    }

    return false;
}

exports.getTodo = (req, res, next) => {
    res.send("test");
}
