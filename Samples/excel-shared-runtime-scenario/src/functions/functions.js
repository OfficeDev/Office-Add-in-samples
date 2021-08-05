/* global CustomFunctions */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

function add(first, second) {
  return first + second;
}

/**
 * Returns data for a given category.
 * @customfunction
 * @param category The category to filter the data with
 * @returns {string[][]} A dynamic array with multiple results.
 */
 function getData(category) {
  //Check that we are connected to service
  const g = getGlobal();

  if (g.state.isConnected) {
    if (category === "Communications") {
      return [["1/1/2017", "The Phone Company", "Communications", "$120"]];
    } else if (category === "Transportation") {
      return [["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"]];
    } else if (category === "Groceries") {
      return [
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"],
      ];
    } else {
      return [
        ["1/1/2017", "The Phone Company", "Communications", "$120"],
        ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
        ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
        ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
        ["1/11/2017", "Bellows College", "Education", "$350"],
        ["1/15/2017", "Trey Research", "Other", "$135"],
        ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"],
      ];
    }
  } else {
    // @ts-ignore
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
  }
}
