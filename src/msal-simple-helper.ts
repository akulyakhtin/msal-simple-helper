import type { AuthenticationResult, IPublicClientApplication } from '@azure/msal-browser'
import { PublicClientApplication } from '@azure/msal-browser'

export interface AuthConfig {
  tenantId: string
  clientId: string
  scopes?: string[]
  useRedirectFlow?: boolean
  redirectResponseHandler?: ((authResult: AuthenticationResult) => void)
  cacheLocation?: string
  noSso?: boolean
}

let msalInstance: IPublicClientApplication | null = null
let authConfig: AuthConfig

/**
 * Creates MSAL singleton instance
 * Throws exception if the instance has already been created with different config
 * @param config config parameters to initialize  MSAL instance with
 * @returns the initialized  instance
 */
export async function msalInit (config: AuthConfig, fnInit: (config: AuthConfig) => Promise<IPublicClientApplication> = doInit): Promise<IPublicClientApplication> {
  validateConfig(config)

  if (msalInstance) throw new Error('MSAL has already been initialized')

  msalInstance = await fnInit(config)

  if (config.useRedirectFlow) {
    if (!config.redirectResponseHandler) throw new Error('Please, set redirectResponseHandler for redirect flow')
    msalInstance.handleRedirectPromise().then(handleRedirectResponse)
      .then((resp) => { if (resp) { config.redirectResponseHandler!(resp) } })
  }

  authConfig = config
  return msalInstance
}

async function doInit (config: AuthConfig): Promise<IPublicClientApplication> {
  const msal = await PublicClientApplication.createPublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: 'https://login.microsoftonline.com/' + config.tenantId
    },
    cache: {
      cacheLocation: config.cacheLocation ?? 'localStorage'
    }
  })
  return msal
}

function handleRedirectResponse (loginResponse: AuthenticationResult | null): AuthenticationResult | null {
  if (loginResponse != null) {
    msalInstance!.setActiveAccount(loginResponse.account)
  } else {
    const currentAccounts = msalInstance!.getAllAccounts()
    if (currentAccounts.length > 0) {
      const activeAccount = currentAccounts[0]
      msalInstance!.setActiveAccount(activeAccount)
    }
  }
  return loginResponse
}

function validateConfig (config?: AuthConfig): void {
  if (!config) throw new Error('Please, provide a valid config')
  if (config.useRedirectFlow && !config.redirectResponseHandler) {
    throw new Error('Please, specify response handler for redirect flow')
  }
}

/**
 * Initialize MSAL, if needed, and performs login
 * If MSAL was already initialized with a different config throws exception
 * @param config config to init MSAL with. Pass null if you have already called msalInit (for redirect flow)
 * @returns authentication result
 */
export async function msalLogin (config?: AuthConfig,
  fnLogin: (msalInstance: IPublicClientApplication, config: AuthConfig) => Promise<AuthenticationResult | undefined> = doLogin): Promise<AuthenticationResult | undefined> {
  // For redirect flow we pass null as config and use the config specified in msalInit before
  if (!config) config = authConfig

  if (!msalInstance) {
    await msalInit(config)
  }
  return await fnLogin(msalInstance!, config)
}

async function doLogin (msalInstance: IPublicClientApplication, config: AuthConfig): Promise<AuthenticationResult | undefined> {
  let loginResponse
  if (config.noSso) {
    loginResponse = await doPopupOrRedirectLogin(msalInstance, config)
  } else {
    try {
      loginResponse = await doSsoLogin(msalInstance)
    } catch (e) {
      console.warn('SSO login failed')
      loginResponse = await doPopupOrRedirectLogin(msalInstance, config)
    }
  }
  msalInstance.setActiveAccount(loginResponse!.account)
  return loginResponse
}

async function doSsoLogin (msalInstance: IPublicClientApplication): Promise<AuthenticationResult | undefined> {
  return await msalInstance.ssoSilent({})
}

async function doPopupOrRedirectLogin (msalInstance: IPublicClientApplication, config: AuthConfig): Promise<AuthenticationResult | undefined> {
  if (config.useRedirectFlow) {
    await msalInstance.loginRedirect() // won't go past this line
  } else {
    return await msalInstance.loginPopup()
  }
}

/**
 * Retrieves access token
 * Throws error if MSAL hasn't been initialized
 * @returns access token response
 */
export async function msalGetAccessToken (
  fnGetToken: (msalInstance: IPublicClientApplication, tokenRequest: { scopes: string[] }) => Promise<AuthenticationResult | undefined> = doGetAccessToken): Promise<AuthenticationResult | undefined> {
  if (msalInstance === null) throw Error('Please, initialize MSAL first.')
  const tokenRequest = {
    scopes: authConfig.scopes ?? []
  }
  return await fnGetToken(msalInstance, tokenRequest)
}

async function doGetAccessToken (msalInstance: IPublicClientApplication, tokenRequest: { scopes: string[] }): Promise<AuthenticationResult | never> {
  let tokenResponse
  try {
    tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest)
  } catch (e) {
    console.error('acquireTokenSilent failed', e)
    tokenResponse = await msalInstance.acquireTokenPopup(tokenRequest)
  }
  return tokenResponse
}

/**
 * Logs out and destroys MSAL
 */
export async function msalLogout (fnLogout: (msalInstance: IPublicClientApplication) => Promise<void> = doLogout): Promise<void> {
  if (msalInstance == null) return
  await fnLogout(msalInstance)
  msalInstance = null
}

async function doLogout (msalInstance: IPublicClientApplication): Promise<void> {
  await msalInstance.logoutPopup()
}

/**
 * @returns Gets MSAL instance (if you want to use the instance directly)
 */
export function msalGetMsal (): IPublicClientApplication | null {
  return msalInstance
}

// For testing
export function msalDestroy (): void {
  msalInstance = null
}
