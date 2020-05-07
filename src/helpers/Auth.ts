import { AuthenticationContext, TokenResponse, ErrorResponse, UserCodeInfo, Logging, LoggingLevel } from 'adal-node';
import { AccessToken } from '../models';

const MS_LOGIN_URL = "https://login.microsoftonline.com";

export interface AuthLogging {
  [name: string]: string;
}

export interface Hash<TValue> {
  [key: string]: TValue;
}

class AuthService {
  connected: boolean = false;
  refreshToken?: string;
  accessTokens: Hash<AccessToken>;
  tenantId?: string;

  constructor() {
    this.accessTokens = {};
  }

  public logout(): void {
    this.connected = false;
    this.accessTokens = {};
    this.refreshToken = undefined;
    this.tenantId = undefined;
  }
}

export class Auth {
  private authCtx: AuthenticationContext = null;
  private service: AuthService = null;
  private userCodeInfo?: UserCodeInfo;

  constructor(private appId: string, private name?: string) {
    this.service = new AuthService();
    this.authCtx = new AuthenticationContext(`${MS_LOGIN_URL}/common`)
  }

  /**
   * Retrieve an accessToken
   * 
   * @param resource 
   * @param log 
   * @param debug 
   * @param fetchNew 
   */
  public async ensureAccessToken(resource: string, authMsg: AuthLogging, debug: boolean = false, fetchNew: boolean = false): Promise<string | null> {
    try {
      const now: Date = new Date();
      const accessToken: AccessToken | undefined = this.service.accessTokens[resource];
      const expiresOn: Date = accessToken ? new Date(accessToken.expiresOn) : new Date(0);

      // Check if there is still an accessToken available
      if (!fetchNew && accessToken && expiresOn > now) {
        if (debug) {
          console.log(`Existing access token ${accessToken.value} still valid. Returning...`);
        }
        return accessToken.value;
      } else {
        if (debug) {
          if (!accessToken) {
            console.log(`No token found for resource ${resource}`);
          } else {
            console.log(`Access token expired. Token: ${accessToken.value}, ExpiresAt: ${accessToken.expiresOn}`);
          }
        }
      }

      let getTokenPromise = this.ensureAccessTokenWithDeviceCode;
      if (this.service.refreshToken) {
        getTokenPromise = this.ensureAccessTokenWithRefreshToken;
      }

      const tokenResponse = await getTokenPromise(resource, debug, authMsg);
      if (!tokenResponse) {
        return null;
      }

      this.service.accessTokens[resource] = {
        expiresOn: tokenResponse.expiresOn as string,
        value: tokenResponse.accessToken
      };
      this.service.refreshToken = tokenResponse.refreshToken;
      this.service.connected = true;

      return this.service.accessTokens[resource].value;
    } catch (error) {
      console.error(`Failed to retrieve an accessToken: ${error.message}`);
      return null;
    }
  }

  /**
   * Cancel the device token flow
   */
  public cancel(): void {
    if (this.userCodeInfo) {
      this.authCtx.cancelRequestToGetTokenWithDeviceCode(this.userCodeInfo as UserCodeInfo, (error: Error, response: TokenResponse | ErrorResponse): void => { });
    }
  }

  /**
   * Retrieve a new accessToken via the device flow
   * 
   * @param resource
   * @param log
   * @param debug
   */
  private ensureAccessTokenWithDeviceCode = (resource: string, debug: boolean, authMsg: AuthLogging): Promise<TokenResponse> => {
    if (debug) {
      console.log(`Starting Auth.ensureAccessTokenWithDeviceCode. resource: ${resource}, debug: ${debug}`);
    }

    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (err: any) => void) => {
      if (debug) {
        console.log('No existing refresh token. Starting new device code flow...');
      }

      this.authCtx.acquireUserCode(resource, this.appId as string, 'en-us',
        (error: Error, response: UserCodeInfo): void => {
          if (debug) {
            console.log('Response:');
            console.log(response);
            console.log('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          authMsg[this.name] = response.message;
          console.log(`${this.name}: ${response.message}`);

          this.userCodeInfo = response;
          this.authCtx.acquireTokenWithDeviceCode(resource, this.appId as string, response,
            (error: Error, response: TokenResponse | ErrorResponse): void => {
              if (debug) {
                console.log('Response:');
                console.log(response);
                console.log('');
              }

              if (error) {
                reject((response && (response as any).error_description) || error.message || (error as any).error_description);
                return;
              }

              authMsg[this.name] = "Already authenticated";
              this.userCodeInfo = undefined;
              resolve(<TokenResponse>response);
            });
        });
    });
  }

  /**
   * Retrieve a new accessToken via the refresh token
   * 
   * @param resource
   * @param log
   * @param debug
   */
  private ensureAccessTokenWithRefreshToken = (resource: string, debug: boolean): Promise<TokenResponse> => {
    return new Promise<TokenResponse>((resolve: (tokenResponse: TokenResponse) => void, reject: (error: any) => void): void => {
      if (debug) {
        console.log(`Retrieving new access token using existing refresh token ${this.service.refreshToken}`);
      }

      this.authCtx.acquireTokenWithRefreshToken(
        this.service.refreshToken as string,
        this.appId as string,
        resource,
        (error: Error, response: TokenResponse | ErrorResponse): void => {
          if (debug) {
            console.log('Response:');
            console.log(response);
            console.log('');
          }

          if (error) {
            reject((response && (response as any).error_description) || error.message);
            return;
          }

          resolve(<TokenResponse>response);
        });
    });
  }
}