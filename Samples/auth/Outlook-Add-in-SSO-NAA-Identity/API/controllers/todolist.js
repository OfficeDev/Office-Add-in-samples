// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const lowdb = require('lowdb');
const FileSync = require('lowdb/adapters/FileSync');
const adapter = new FileSync('./data/db.json');
const db = lowdb(adapter);
const { v4: uuidv4 } = require('uuid');

const authConfig = require('../server-helpers/authConfig');

// Get todo list for user identified by oid claim and return the list to caller.
exports.getTodos = (req, res, next) => {
    // Check that caller has the delegated todolist.read permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.read)) {

        try {
            // For multi-tenant apps, the owner is a combination of oid and tid claims.
            // For more information, see https://learn.microsoft.com/azure/active-directory/develop/scenario-desktop-acquire-token-overview#multi-tenant-apps
            const owner = req.authInfo['oid'] + req.authInfo['tid'];

            const todos = db.get('todos')
                .filter({ owner: owner })
                .value();

            res.status(200).send(todos);
        } catch (error) {
            next(error);
        }
    } else {
        res.status(403).send('User does not have the required permissions.');
    }
}

// Add a new todo item to the db for the user identified by oid claim.
exports.postTodo = (req, res, next) => {
    // Check that caller has the delegated todolist.readwrite permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.readWrite)) {

        try {
            const todo = {
                description: req.body.description,
                id: uuidv4(),
                // For multi-tenant apps, the owner is a combination of oid and tid claims.
                // For more information, see https://learn.microsoft.com/azure/active-directory/develop/scenario-desktop-acquire-token-overview#multi-tenant-apps
                owner: req.authInfo['oid'] + req.authInfo['tid']
            };

            db.get('todos').push(todo).write();

            res.status(200).json(todo);
        } catch (error) {
            next(error);
        }
    } else {
        res.status(403).send('User does not have the required permissions.');
}
}

// Delete a todo item by id and user's oid claim.
exports.deleteTodo = (req, res, next) => {
    // Check that caller has the delegated todolist.readwrite permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.readWrite)) {
        try {
            const id = req.params.id;
            // For multi-tenant apps, the owner is a combination of oid and tid claims.
            // For more information, see https://learn.microsoft.com/azure/active-directory/develop/scenario-desktop-acquire-token-overview#multi-tenant-apps
            const owner = req.authInfo['oid'] + req.authInfo['tid'];

            db.get('todos')
                .remove({ owner: owner, id: id })
                .write();

            res.status(200).json({ message: "success" });
        } catch (error) {
            next(error);
        }
    } else {
        res.status(403).send('User does not have the required permissions.');
    }
}

/**
 * Ensures that the access token has the specified delegated permissions.
 * @param {Object} accessTokenPayload: Parsed access token payload.
 * @param {Array} requiredPermission: list of required permissions.
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
