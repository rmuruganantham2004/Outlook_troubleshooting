export interface PowerShellResult {
    success: boolean;
    output: string;
    error?: string;
    exitCode: number;
}
export declare class PowerShellService {
    private executionMode;
    constructor();
    /**
     * Run a PowerShell script file with optional parameters
     */
    runScript(scriptName: string, args?: string[]): Promise<PowerShellResult>;
    /**
     * Run a raw PowerShell command string directly (with safety checks)
     */
    runRawCommand(command: string): Promise<PowerShellResult>;
    /**
     * Simulated script runner containing rich responses for simulation profiles
     */
    private simulateScriptRun;
    private simulateRawCommand;
}
