/* global console, self, window, global */

function getGlobal() {
    console.log("init globals for command buttons");
    return typeof self !== "undefined"
      ? self
      : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
      ? global
      : undefined;
  }
  
  