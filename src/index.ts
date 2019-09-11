import 'isomorphic-fetch';
import * as usage from './usage';

require('dotenv').config();

(async () => {
  console.log('Starting...')
  await usage.run();
  console.log('Done.');
})();
