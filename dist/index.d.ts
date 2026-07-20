export interface PowerShellResult {
    success: boolean;
    output: string;
    error?: string;
    exitCode: number;
}
declare const _default: import("openclaw/plugin-sdk/tool-plugin").DefinedToolPluginEntry;
export default _default;
