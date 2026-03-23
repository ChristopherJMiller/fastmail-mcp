export interface FastmailConfig {
  apiToken: string;
  baseUrl?: string;
  cookies?: string;
}

function normalizeBaseUrl(input?: string): string {
  const DEFAULT = 'https://api.fastmail.com';
  if (!input) return DEFAULT;
  let url = input.trim();
  if (!url) return DEFAULT;
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(url)) {
    url = 'https://' + url;
  }
  url = url.replace(/\/+$/, '');
  return url;
}

export class FastmailAuth {
  private apiToken: string;
  private baseUrl: string;
  private cookies: string | undefined;

  constructor(config: FastmailConfig) {
    this.apiToken = config.apiToken;
    this.baseUrl = normalizeBaseUrl(config.baseUrl);
    this.cookies = config.cookies;
  }

  getAuthHeaders(): Record<string, string> {
    const headers: Record<string, string> = {
      'Authorization': `Bearer ${this.apiToken}`,
      'Content-Type': 'application/json'
    };
    if (this.cookies) {
      headers['Cookie'] = this.cookies;
    }
    return headers;
  }

  getSessionUrl(): string {
    return `${this.baseUrl}/jmap/session`;
  }

  getApiUrl(): string {
    return `${this.baseUrl}/jmap/api/`;
  }

  getCookieHeaders(): Record<string, string> | null {
    if (!this.cookies) return null;
    return {
      'Content-Type': 'application/json',
      'Cookie': this.cookies,
      'X-TrustedClient': 'Yes',
      'Origin': 'https://app.fastmail.com'
    };
  }

  getBaseUrl(): string {
    return this.baseUrl;
  }
}