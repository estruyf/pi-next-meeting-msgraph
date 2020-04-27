import fetch from 'node-fetch';

export class MsGraphService {
  
  /**
   * Perform a get request against the MS Graph
   * 
   * @param url 
   * @param accessToken 
   */
  public static async get(url: string, accessToken: string, debug: boolean = false) {
    if (debug) {
      console.log(`Calling the MS Graph.`);
    }

    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': `application/json`,
        'Accept': `application/json`
      }
    });

    const data = await response.json();
    if (debug) {
      console.log(`MS Graph response: ${JSON.stringify(data)}`);
    }
    
    return data;
  }
}