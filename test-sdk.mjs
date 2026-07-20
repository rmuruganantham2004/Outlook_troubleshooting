import plugin from "./dist/index.js";
import { getToolPluginMetadata } from "openclaw/plugin-sdk/tool-plugin";

console.dir(getToolPluginMetadata(plugin), { depth: null });