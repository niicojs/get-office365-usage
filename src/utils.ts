import {
  PageIterator,
  Client as GraphClient
} from '@microsoft/microsoft-graph-client';

export const iterate = async (client: GraphClient, response) => {
  return new Promise(async (resolve, reject) => {
    const result = [];
    try {
      const pageIterator = new PageIterator(client, response, data => {
        result.push(data);
        return true;
      });
      await pageIterator.iterate();
      resolve(result);
    } catch (e) {
      reject(e);
    }
  });
};
