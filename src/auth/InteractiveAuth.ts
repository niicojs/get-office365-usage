import * as fs from 'fs';
import * as express from 'express';
import * as open from 'open';
import * as proxyAgent from 'https-proxy-agent';
import * as got from 'got';
import { URLSearchParams } from 'url';
import { UserRefreshTokenProvider } from './UserRefreshTokenProvider';

export class InteractiveAuth {
  persist: boolean;
  tenant: string;
  auth: any;

  constructor(options: any = {}) {
    this.persist = !!options.persist;
    this.tenant = options.tenant || process.env.TENANT;
  }

  scope() {
    return [
      'openid',
      'profile',
      'offline_access',
      'Reports.Read.All'
    ].join(' ');
  }

  buildProvider(tokeninfo) {
    return new UserRefreshTokenProvider({
      tenant: this.tenant,
      expire: tokeninfo.expire,
      token: tokeninfo.token,
      refreshToken: tokeninfo.refreshToken,
      scope: this.scope()
    });
  }

  launch() {
    return new Promise(async (resolve, reject) => {
      try {
        // if existing token exists on disk, use it
        if (this.persist && fs.existsSync('token.bin')) {
          const token = await fs.promises.readFile('token.bin', 'utf8');
          return resolve({
            refreshToken: token,
            expire: 0
          });
        }

        // launch local webserver to get response

        const app = express();
        app.get('/', async (req, res) => {
          try {
            const code = req.query.code;
            const error = req.query.error_description;
            if (error) {
              return res.send(`<div>Error</div><div>${error}</div>`);
            }

            const result = await got.post(
              `https://login.microsoftonline.com/${this.tenant}/oauth2/v2.0/token`,
              {
                form: true,
                body: {
                  client_id: process.env.APP_ID,
                  scope: this.scope(),
                  code,
                  redirect_uri: 'http://localhost',
                  client_secret: process.env.APP_SECRET,
                  grant_type: 'authorization_code'
                },
                json: true,
                agent: process.env.PROXY
                  ? new proxyAgent(process.env.PROXY)
                  : undefined
              }
            );

            const data = result.body;
            this.auth = data;
            res.send('ok').end();

            if (this.persist) {
              await fs.promises.writeFile(
                'token.bin',
                data.refresh_token,
                'utf8'
              );
            }

            resolve({
              token: data.access_token,
              expire: data.expires_in,
              refreshToken: data.refresh_token
            });
          } catch (e) {
            res.send('an error occured').end();
            reject(e);
          }
        });

        app.listen(80);

        // launch url

        const url = `https://login.microsoftonline.com/${this.tenant}/oauth2/v2.0/authorize`;

        const params = new URLSearchParams({
          client_id: process.env.APP_ID,
          response_type: 'code',
          redirect: 'http://localhost',
          response_mode: 'query',
          scope: this.scope()
        });

        console.log('Launching browser');

        await open(`${url}?${params.toString()}`);

        console.log('Waiting for user login');
      } catch (e) {
        reject(e);
      }
    });
  }
}
