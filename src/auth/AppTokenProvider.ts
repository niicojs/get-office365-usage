import * as proxyAgent from 'https-proxy-agent';
import * as got from 'got';

export class AppTokenProvider {
  tenant: string;
  appId: string;
  appSecret: string;
  token: string;
  expire: number;

  constructor(options) {
    this.tenant = options.tenant;
    this.appId = options.appId;
    this.appSecret = options.appSecret;
  }

  async getAppToken() {
    // const tenant = ths;
    const result = await got.post(
      `https://login.microsoftonline.com/${this.tenant}/oauth2/v2.0/token`,
      {
        form: true,
        body: {
          client_id: this.appId,
          scope: 'https://graph.microsoft.com/.default',
          client_secret: this.appSecret,
          grant_type: 'client_credentials'
        },
        json: true,
        agent: process.env.PROXY ? new proxyAgent(process.env.PROXY) : undefined
      }
    );

    return {
      token: result.body.access_token,
      expire: Date.now() + 1000 * +result.body.expires_in
    };
  }

  public async getAccessToken(): Promise<string> {
    // TODO check token expiration
    if (!this.token) {
      const { token, expire } = await this.getAppToken();
      this.token = token;
      this.expire = expire;
    }
    return this.token;
  }
}
