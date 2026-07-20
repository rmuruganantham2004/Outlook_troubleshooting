import plugin from "./dist/index.js";

const registered = [];

plugin.register({
  pluginConfig: {},

  registerTool(tool, opts) {
    console.log("REGISTER TOOL:", tool.name);
    registered.push(tool.name);
  }
});

console.log("Registered:", registered);