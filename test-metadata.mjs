import plugin from "./dist/index.js";

const symbol = Symbol.for("openclaw.plugin-sdk.tool-plugin.metadata");

console.log("Metadata:");
console.dir(plugin[symbol], { depth: null });
