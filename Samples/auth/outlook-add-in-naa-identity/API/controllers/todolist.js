const lowdb = require('lowdb');
const FileSync = require('lowdb/adapters/FileSync');
const adapter = new FileSync('./data/db.json');
const db = lowdb(adapter);
const { v4: uuidv4 } = require('uuid');

const {
  isAppOnlyToken,
  hasRequiredDelegatedPermissions,
  hasRequiredApplicationPermissions
} = require('../auth/permissionUtils');

const authConfig = require('../authConfig');

exports.getTodo = (req, res, next) => {
    if (isAppOnlyToken(req.authInfo)) {
        if (hasRequiredApplicationPermissions(req.authInfo, authConfig.protectedRoutes.todolist.applicationPermissions.read)) {
            try {
                const id = req.params.id;
    
                const todo = db.get('todos')
                    .find({ id: id })
                    .value();
    
                res.status(200).send(todo);
            } catch (error) {
                next(error);
            }
        } else {
            next(new Error('Application does not have the required permissions'))
        }
    } else {
        if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.read)) {
            try {
                /**
                 * The 'oid' (object id) is the only claim that should be used to uniquely identify
                 * a user in an Azure AD tenant. The token might have one or more of the following claim,
                 * that might seem like a unique identifier, but is not and should not be used as such,
                 * especially for systems which act as system of record (SOR):
                 *
                 * - upn (user principal name): might be unique amongst the active set of users in a tenant but
                 * tend to get reassigned to new employees as employees leave the organization and
                 * others take their place or might change to reflect a personal change like marriage.
                 *
                 * - email: might be unique amongst the active set of users in a tenant but tend to get
                 * reassigned to new employees as employees leave the organization and others take their place.
                 */
                const owner = req.authInfo['oid'];
                const id = req.params.id;
    
                const todo = db.get('todos')
                    .filter({ owner: owner })
                    .find({ id: id })
                    .value();
    
                res.status(200).send(todo);
            } catch (error) {
                next(error);
            }
        } else {
            next(new Error('User does not have the required permissions'))
        }
    }
}

exports.getTodos = (req, res, next) => {
    if (isAppOnlyToken(req.authInfo)) {
        if (hasRequiredApplicationPermissions(req.authInfo, authConfig.protectedRoutes.todolist.applicationPermissions.read)) {
            try {
                const todos = db.get('todos')
                    .value();
    
                res.status(200).send(todos);
            } catch (error) {
                next(error);
            }
        } else {
            next(new Error('Application does not have the required permissions'))
        }
    } else {
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
            next(new Error('User does not have the required permissions'))
        }
    }
}

exports.postTodo = (req, res, next) => {
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.write)
        ||
        hasRequiredApplicationPermissions(req.authInfo, authConfig.protectedRoutes.todolist.applicationPermissions.write)
    ) {
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
        next(new Error('User or application does not have the required permissions'))
    )
}

exports.deleteTodo = (req, res, next) => {
    if (isAppOnlyToken(req.authInfo)) {
        if (hasRequiredApplicationPermissions(req.authInfo, authConfig.protectedRoutes.todolist.applicationPermissions.write)) {
            try {
                const id = req.params.id;
    
                db.get('todos')
                    .remove({ id: id })
                    .write();
    
                res.status(200).json({ message: "success" });
            } catch (error) {
                next(error);
            }
        } else {
            next(new Error('Application does not have the required permissions'))
        }
    } else {
        if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.write)) {
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
}