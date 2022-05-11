module.exports = async function (context, req) {
  context.log("JavaScript HTTP trigger function processed a request.");

  //retrieve parameters if passed on URL.
  let first = req.query.first;
  let second = req.query.second;

  //Check if parameters were passed in body text.
  if (req.body !== undefined) {
    if (req.body.first !== undefined) {
      first = req.body.first;
    }
    if (req.body.second !== undefined) {
      second = req.body.second;
    }
  }
  if (isNaN(first) || isNaN(second)) {
    context.res = {
      status: 400, //bad request
      body: "Please pass (first,second) number parameters in the query string or in the request body",
    };
  } else {
    context.res = {
      // status: 200, /* Defaults to 200 */
      body: {
        answer: Number(first) + Number(second),
      },
    };
  }
};
