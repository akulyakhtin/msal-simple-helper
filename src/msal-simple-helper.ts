import type { AuthenticationResult, IPublicClientApplication } from '@azure/msal-browser'
import { PublicClientApplication } from '@azure/msal-browser'

export interface AuthConfig {
  tenantId: string
  clientId: string
  scopes?: string[]
  flow?: string
  redirectResponseHandler?: ((authResult: AuthenticationResult) => void)
  cacheLocation?: string
}

let msalInstance: IPublicClientApplication | null = null
let authConfig: AuthConfig | null = null

/**
 * Creates MSAL singleton instance
 * Throws exception if the instance has already been created with different config
 * @param config config parameters to initialize  MSAL instance with
 * @returns the initialized  instance
 */
export async function msalInit (config: AuthConfig, fnInit: (config: AuthConfig) => Promise<IPublicClientApplication> = doInit): Promise<IPublicClientApplication> {
  validateConfig(config)

  // Use the existing MSAL if any
  // and don't allow recreating unless destroyed before
  if (msalInstance) {
    if (authConfig && configsDiffer(config, authConfig)) {
      throw new Error('MSAL has already been initialized with different config')
    } else {
      return await Promise.resolve(msalInstance)
    }
  }

  msalInstance = await fnInit(config)

  if (config.flow?.toLowerCase() === 'redirect') {
    msalInstance.handleRedirectPromise().then(handleRedirectResponse)
      .then((resp) => { if (resp) { config.redirectResponseHandler!(resp) } })
  }

  authConfig = config
  return msalInstance
}

/**
 * Initialize MSAL, if needed, and performs login
 * If MSAL was already initialized with a different config throws exception
 * @param config Iconfig to init MSAL with
 * @returns authentication result
 */
export async function msalLogin (config: AuthConfig,
  fnLogin: (msalInstance: IPublicClientApplication, config: AuthConfig) => Promise<AuthenticationResult | undefined> = doLogin): Promise<AuthenticationResult | undefined> {
  const msalInstance = await msalInit(config)
  return await fnLogin(msalInstance, config)
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
    scopes: authConfig?.scopes ?? []
  }
  return await fnGetToken(msalInstance, tokenRequest)
}

/**
 * Logs out and destroys MSAL
 */
export async function msalLogout (fnLogout: (msalInstance: IPublicClientApplication) => Promise<void> = doLogout): Promise<void> {
  if (msalInstance == null) return
  await fnLogout(msalInstance)
  msalInstance = null
  authConfig = null
}

/**
 *
 * @returns Gets MSAL instance (if you want to use the instance directly)
 */
export function msalGetMsal (): IPublicClientApplication | null {
  return msalInstance
}

// Implementation

function validateConfig (config?: AuthConfig): void {
  if (!config) throw new Error('Please, provide a valid config')
  if (config.flow && !(['popup', 'redirect'].includes(config.flow.toLowerCase()))) throw new Error('Flow should be either popup or redirect')
  if (config.flow?.toLowerCase() === 'redirect' && !config.redirectResponseHandler) {
    throw new Error('Please, specify response handler for redirect flow')
  }
}

function configsDiffer (cf1: AuthConfig, cf2: AuthConfig): boolean {
  return (cf1.tenantId !== cf2.tenantId ||
    cf1.clientId !== cf2.clientId ||
    cf1.scopes !== cf2.scopes)
}

async function doLogout (msalInstance: IPublicClientApplication): Promise<void> {
  await msalInstance.logoutPopup()
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

async function doLogin (msalInstance: IPublicClientApplication, config: AuthConfig): Promise<AuthenticationResult | undefined> {
  let loginResponse = null
  try {
    loginResponse = await msalInstance.ssoSilent({})
    msalInstance.setActiveAccount(loginResponse.account)
    return loginResponse
  } catch (e) {
    console.warn('ssoSilent failed', e)
    if (config.flow?.toLowerCase() === 'popup') {
      loginResponse = await msalInstance.loginPopup()
      msalInstance.setActiveAccount(loginResponse.account)
      return loginResponse
    } else {
      await msalInstance.loginRedirect() // won't go past this line
    }
  }
}

function handleRedirectResponse (loginResponse: AuthenticationResult | null): AuthenticationResult | null {
  if (msalInstance === null) throw new Error('msalInstance must not be null')
  if (loginResponse != null) {
    msalInstance.setActiveAccount(loginResponse.account)
  } else {
    const currentAccounts = msalInstance.getAllAccounts()
    if (currentAccounts.length > 0) {
      const activeAccount = currentAccounts[0]
      msalInstance.setActiveAccount(activeAccount)
    }
  }
  return loginResponse
}

// For testing
export function msalDestroy (): void {
  msalInstance = null
  authConfig = null
}
