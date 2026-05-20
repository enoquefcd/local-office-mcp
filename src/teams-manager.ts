import { spawn } from 'child_process';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

export interface TeamsChat {
  id: string;
  displayName: string;
  chatType: string;
  members: string[];
}

export interface TeamsMessage {
  id: string;
  content: string;
  from: string;
  time: string;
  chat_id: string;
  type?: string;
}

export interface TeamsChannel {
  id: string;
  channelName: string;
  teamName: string;
  teamId: string;
  channelType: string;
}

const IS_WSL = process.platform === 'linux' &&
  (process.env.WSL_DISTRO_NAME !== undefined ||
   process.env.WSLENV !== undefined ||
   require('fs').existsSync('/proc/version') &&
     require('fs').readFileSync('/proc/version', 'utf8').toLowerCase().includes('microsoft'));

const PYTHON_CANDIDATES = IS_WSL
  ? ['python3', 'python']
  : ['python.exe', 'python3.exe', 'py.exe'];

export class TeamsManager {
  private scriptPath: string;
  private idbPathOverride?: string;

  constructor(idbPathOverride?: string) {
    const __filename = fileURLToPath(import.meta.url);
    const __dirname = dirname(__filename);
    this.scriptPath = join(__dirname, '..', 'scripts', 'read_teams_idb.py');
    this.idbPathOverride = idbPathOverride;
  }

  private async executePython(action: string, params: Record<string, string | number | undefined> = {}): Promise<any> {
    const args = [this.scriptPath, '--action', action];

    for (const [key, value] of Object.entries(params)) {
      if (value !== undefined && value !== null) {
        args.push(`--${key}`, String(value));
      }
    }

    if (this.idbPathOverride) {
      args.push('--idb_path', this.idbPathOverride);
    }

    return new Promise((resolve, reject) => {
      const tryPython = async (candidates: string[]): Promise<void> => {
        const pythonExe = candidates[0];
        const remaining = candidates.slice(1);

        const py = spawn(pythonExe, args, {
          env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
        });

        let stdout = '';
        let stderr = '';

        py.stdout.setEncoding('utf8');
        py.stderr.setEncoding('utf8');
        py.stdout.on('data', (d) => { stdout += d; });
        py.stderr.on('data', (d) => { stderr += d; });

        py.on('error', (err: NodeJS.ErrnoException) => {
          if (err.code === 'ENOENT' && remaining.length > 0) {
            tryPython(remaining).then(resolve).catch(reject);
          } else {
            reject(new Error(`Python not found. Install Python 3 and run: pip install ccl-chromium-indexeddb`));
          }
        });

        py.on('close', (code) => {
          const raw = stdout.trim();
          if (!raw) {
            reject(new Error(`Teams script produced no output (exit ${code}): ${stderr.trim()}`));
            return;
          }
          try {
            const result = JSON.parse(raw);
            if (result.error) {
              reject(new Error(result.error));
            } else {
              resolve(result);
            }
          } catch {
            reject(new Error(`Failed to parse Teams script output: ${raw.substring(0, 200)}`));
          }
        });
      };

      tryPython(PYTHON_CANDIDATES).catch(reject);
    });
  }

  async getChats(count = 20): Promise<TeamsChat[]> {
    const result = await this.executePython('get_chats', { count });
    return (result.data ?? []) as TeamsChat[];
  }

  async getMessages(chatId: string, count = 50): Promise<TeamsMessage[]> {
    const result = await this.executePython('get_messages', { chat_id: chatId, count });
    return (result.data ?? []) as TeamsMessage[];
  }

  async searchMessages(query: string, count = 20): Promise<TeamsMessage[]> {
    const result = await this.executePython('search_messages', { query, count });
    return (result.data ?? []) as TeamsMessage[];
  }

  async getChannels(count = 50): Promise<TeamsChannel[]> {
    const result = await this.executePython('get_channels', { count });
    return (result.data ?? []) as TeamsChannel[];
  }

  async listStores(): Promise<string[]> {
    const result = await this.executePython('list_stores');
    return (result.data ?? []) as string[];
  }
}
