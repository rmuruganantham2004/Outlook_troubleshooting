import { exec } from 'child_process';
import path from 'path';
import fs from 'fs';
import os from 'os';

export interface PowerShellResult {
  success: boolean;
  output: string;
  error?: string;
  exitCode: number;
}

export class PowerShellService {
  private executionMode: string;

  constructor() {
    this.executionMode = process.env.EXECUTION_MODE || 'simulation';
  }

  /**
   * Run a PowerShell script file with optional parameters
   */
  public async runScript(scriptName: string, args: string[] = []): Promise<PowerShellResult> {
    const isWindows = os.platform() === 'win32';
    
    if (this.executionMode === 'simulation' || !isWindows) {
      return this.simulateScriptRun(scriptName, args);
    }

    // Resolve script location relative to execution context
    const scriptPath = path.resolve(__dirname, './scripts/powershell', scriptName);
    if (!fs.existsSync(scriptPath)) {
      return {
        success: false,
        output: '',
        error: `Script not found at path: ${scriptPath}`,
        exitCode: 1,
      };
    }

    const escapedArgs = args.map(arg => `"${arg.replace(/"/g, '\\"')}"`).join(' ');
    const command = `powershell -NoProfile -ExecutionPolicy Bypass -File "${scriptPath}" ${escapedArgs}`;

    return new Promise((resolve) => {
      exec(command, (error, stdout, stderr) => {
        if (error) {
          resolve({
            success: false,
            output: stdout.trim(),
            error: stderr.trim() || error.message,
            exitCode: error.code || 1,
          });
        } else {
          resolve({
            success: true,
            output: stdout.trim(),
            exitCode: 0,
          });
        }
      });
    });
  }

  /**
   * Run a raw PowerShell command string directly (with safety checks)
   */
  public async runRawCommand(command: string): Promise<PowerShellResult> {
    const isWindows = os.platform() === 'win32';
    
    // Safety check: block command injections or destructive cmds
    const dangerousTokens = ['rm ', 'del ', 'format ', 'sfc /scannow', 'shutdown', 'restart-computer'];
    const lowerCmd = command.toLowerCase();
    if (dangerousTokens.some(token => lowerCmd.includes(token))) {
      return {
        success: false,
        output: '',
        error: 'Execution blocked: Command contains forbidden operations for safety reasons.',
        exitCode: -1,
      };
    }

    if (this.executionMode === 'simulation' || !isWindows) {
      return this.simulateRawCommand(command);
    }

    const powershellCmd = `powershell -NoProfile -ExecutionPolicy Bypass -Command "${command.replace(/"/g, '\\"')}"`;

    return new Promise((resolve) => {
      exec(powershellCmd, (error, stdout, stderr) => {
        if (error) {
          resolve({
            success: false,
            output: stdout.trim(),
            error: stderr.trim() || error.message,
            exitCode: error.code || 1,
          });
        } else {
          resolve({
            success: true,
            output: stdout.trim(),
            exitCode: 0,
          });
        }
      });
    });
  }

  /**
   * Simulated script runner containing rich responses for simulation profiles
   */
  private simulateScriptRun(scriptName: string, args: string[]): PowerShellResult {
    const action = scriptName.replace('.ps1', '');

    switch (action) {
      case 'Check-OutlookRunning':
        return {
          success: true,
          output: JSON.stringify({ isRunning: true, processId: 10452, threads: 14 }),
          exitCode: 0,
        };

      case 'Repair-Office':
        return {
          success: true,
          output: 'Office Click-To-Run Quick Repair completed successfully. Executed in 4.2 seconds.',
          exitCode: 0,
        };

      case 'Scan-OST': {
        const ostPath = args[0] || 'C:\\Users\\MockUser\\AppData\\Local\\Microsoft\\Outlook\\outlook.ost';
        return {
          success: true,
          output: JSON.stringify({
            filePath: ostPath,
            sizeGB: 18.4,
            corrupted: true,
            corruptionPercentage: 1.5,
            message: 'OST has minor corruptions. Recommend recreating the profile or using repair tool.'
          }),
          exitCode: 0,
        };
      }

      case 'Repair-PST': {
        return {
          success: true,
          output: 'ScanPST.exe found 24 folders, 1024 messages. Executed repairs. OST/PST is healthy.',
          exitCode: 0,
        };
      }

      case 'Reset-NavigationPane':
        return {
          success: true,
          output: 'Outlook navigation pane configuration files reset. Cache deleted.',
          exitCode: 0,
        };

      case 'Test-ExchangeConnection':
        return {
          success: true,
          output: JSON.stringify({
            latencyMs: 45,
            exchangeServer: 'outlook.office365.com',
            connectivity: 'Connected',
            authType: 'ModernAuth',
            tlsVersion: 'TLS 1.3',
          }),
          exitCode: 0,
        };

      case 'Collect-EventViewerLogs':
        return {
          success: true,
          output: JSON.stringify([
            { id: 1000, source: 'Application Error', message: 'Faulting application name: OUTLOOK.EXE, version: 16.0.14326.20454, exception code: 0xc0000005', time: new Date().toISOString() },
            { id: 2002, source: 'Outlook', message: 'Connection to server lost. Retrying...', time: new Date(Date.now() - 50000).toISOString() }
          ]),
          exitCode: 0,
        };

      default:
        return {
          success: true,
          output: `Simulated output for ${scriptName}`,
          exitCode: 0,
        };
    }
  }

  private simulateRawCommand(command: string): PowerShellResult {
    const lower = command.toLowerCase();

    if (lower.includes('get-process outlook')) {
      return {
        success: true,
        output: 'OUTLOOK (PID 10452)',
        exitCode: 0,
      };
    }

    if (lower.includes('nslookup') || lower.includes('resolve-dnsname')) {
      return {
        success: true,
        output: 'Name: outlook.office365.com\nAddress: 40.97.120.242\nName: autodiscover.outlook.com\nCNAME: autodiscover.outlook.office365.com',
        exitCode: 0,
      };
    }

    if (lower.includes('test-netconnection')) {
      return {
        success: true,
        output: 'ComputerName: outlook.office365.com\nRemoteAddress: 40.97.120.242\nInterfaceAlias: Wi-Fi\nTcpTestSucceeded: True',
        exitCode: 0,
      };
    }

    return {
      success: true,
      output: 'Command completed successfully in simulation mode.',
      exitCode: 0,
    };
  }
}
