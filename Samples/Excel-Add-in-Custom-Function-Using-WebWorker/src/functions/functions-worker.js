// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

self.addEventListener('message',
    function(event) {
        let job = event.data;
        if (typeof(job) == "string") {
            job = JSON.parse(job);
        }

        const jobId = job.jobId;
        try {
            const result = invokeFunction(job.name, job.parameters);
            // check whether the result is a promise.
            if (typeof(result) == "function" || typeof(result) == "object" && typeof(result.then) == "function") {
                result.then(function(realResult) {
                    postMessage(
                        {
                            jobId: jobId,
                            result: realResult
                        }
                    );
                })
                .catch(function(ex) {
                    postMessage(
                        {
                            jobId: jobId,
                            error: true
                        }
                    )
                });
            }
            else {
                postMessage({
                    jobId: jobId,
                    result: result
                });
            }
        }
        catch(ex) {
            postMessage({
                jobId: jobId,
                error: true
            });
        }
    }
);

function invokeFunction(name, parameters) {
    if (name == "TEST") {
        return test.apply(null, parameters);
    }
    else if (name == "TEST_PROMISE") {
        return test_promise.apply(null, parameters);
    }
    else if (name == "TEST_ERROR") {
        return test_error.apply(null, parameters);
    }
    else if (name == "TEST_ERROR_PROMISE") {
        return test_error_promise.apply(null, parameters);
    }
    else {
        throw new Error("not supported");
    }
}

function test(n) {
    let ret = 0;
    for (let i = 0; i < n; i++) {
        ret += Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));
        for (let l = 0; l < n; l++) {
            ret -= Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(Math.tan(Math.atan(50))))))))));     
        }
    }
    return ret;
}


function test_promise(n) {
    return new Promise(function(resolve, reject) {
        setTimeout(function() {
            resolve(test(n));
        }, 1000);
    });
}

function test_error(n) {
    throw new Error();
}

function test_error_promise(n) {
    return Promise.reject(new Error());
}