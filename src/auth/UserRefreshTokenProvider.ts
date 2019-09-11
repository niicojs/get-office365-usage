import * as proxyAgent from 'https-proxy-agent';
import * as got from 'got';

export class UserRefreshTokenProvider {
  tenant: string;
  token: string;
  expire: number;
  refreshToken: string;
  scope: string;

  constructor(options) {
    this.tenant = options.tenant;
    this.token = options.token;
    this.expire = Date.now() + 1000 * +options.expire;
    this.refreshToken = options.refreshToken;
    this.scope = options.scope;
  }

  async refresh() {
    const result = await got.post(
      `https://login.microsoftonline.com/${this.tenant}/oauth2/v2.0/token`,
      {
        form: true,
        body: {
          client_id: process.env.APP_ID,
          scope: this.scope,
          refresh_token: this.refreshToken,
          redirect_uri: 'http://localhost',
          client_secret: process.env.APP_SECRET,
          grant_type: 'refresh_token'
        },
        json: true,
        agent: process.env.PROXY ? new proxyAgent(process.env.PROXY) : undefined
      }
    );

    const data = result.body;
    this.token = data.access_token;
    this.expire = Date.now() + 1000 * +data.expires_id;
    this.refreshToken = data.refresh_token;
  }

  public async getAccessToken(): Promise<string> {
    if (Date.now() >= this.expire) {
      await this.refresh();
    }
    return this.token;
  }
}
