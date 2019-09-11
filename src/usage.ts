import { AppTokenProvider } from './auth/AppTokenProvider';
import * as moment from 'moment';
import * as proxyAgent from 'https-proxy-agent';
import { Client as GraphClient } from '@microsoft/microsoft-graph-client';
import { iterate } from './utils';

const main = async () => {
  const auth = new AppTokenProvider({
    tenant: process.env.TENANT_ID,
    appId: process.env.APP_ID,
    appSecret: process.env.APP_SECRET
  });
  const token = await auth.getAccessToken();
  console.log('Logged in.');

  const client = GraphClient.init({
    defaultVersion: 'beta',
    authProvider: done => done(null, token),
    fetchOptions: {
      agent: process.env.PROXY ? new proxyAgent(process.env.PROXY) : undefined
    }
  });

  const day = moment()
    .add(-2, 'day')
    .format('YYYY-MM-DD');

  {
    const response = await client
      .api(`/reports/getEmailActivityUserDetail(date=${day})`)
      .query('$format=application/json')
      .get();
    const data = await iterate(client, response);
    console.log(data);
  }

  {
    const response = await client
      .api(`/reports/getEmailAppUsageUserDetail(date=${day})`)
      .query('$format=application/json')
      .get();
    const data = await iterate(client, response);
    console.log(data);
  }

  {
    const response = await client
      .api(`/reports/getMailboxUsageDetail(date=${day})`)
      .query('$format=application/json')
      .get();
    const data = await iterate(client, response);
    console.log(data);
  }

  {
    const response = await client
      .api(`/reports/getOneDriveActivityUserDetail(date=${day})`)
      .query('$format=application/json')

      .get();
    const data = await iterate(client, response);
    console.log(data);
  }
};

export const run = async () => {
  try {
    console.log('Getting usage reports datas...');
    await main();
  } catch (e) {
    console.log(`Error ${e.statusCode} ${e.code}`);
    console.log(`${e.message}`);
  }
};
