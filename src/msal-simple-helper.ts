import type { AuthenticationResult, PopupRequest, RedirectRequest } from '@azure/msal-browser'
import { PublicClientApplication } from '@azure/msal-browser'
import type { MsalLike } from './msalLike'

export interface AuthConfig {
  tenantId: string
  clientId: string
  scopes?: string[]
  useRedirectFlow?: boolean
  redirectResponseHandler?: ((authResult: AuthenticationResult) => void)
  cacheLocation?: string
  noSso?: boolean
  noAcquireTokenSilent?: boolean
}

let msalInstance: MsalLike
let authConfig: AuthConfig

/**
 * Creates MSAL singleton instance
 * Throws exception if the instance has already been created with different config
 * @param config config parameters to initialize  MSAL instance with
 * @returns the initialized  instance
 */
export async function msalInit (config: AuthConfig): Promise<MsalLike> {

  if (msalInstance) throw new Error('MSAL has already been initialized')

  validateConfig(config)

  authConfig = config

  msalInstance = await msalCreator(authConfig)

  if (authConfig.useRedirectFlow) {
    msalInstance.handleRedirectPromise().then(handleRedirectResponse)
      .then((resp) => { if (resp) { config.redirectResponseHandler!(resp) } })
  }

  return msalInstance
}

/**
 * Initialize MSAL, if needed, and performs login
 * @param config config to init MSAL with. Pass null if you have already called msalInit (for redirect flow)
 * @returns authentication result
 */
export async function msalLogin (config?: AuthConfig): Promise<AuthenticationResult | undefined> {

  if (!msalInstance) {
    if (!config) throw Error('config must not be null')
    await msalInit(config)
  }
  return await doLogin()
}

/**
 * Retrieves access token
 * Throws error if MSAL hasn't been initialized
 * @returns access token response
 */
export async function msalGetAccessToken (): Promise<AuthenticationResult | undefined> {

  if (msalInstance === null) throw Error('Please, initialize MSAL first.')

  const tokenRequest = {
    scopes: authConfig.scopes ?? []
  }
  return await doGetAccessToken(tokenRequest)
}

/**
 * Logs out
 */
export async function msalLogout (): Promise<void> {
  if (msalInstance == null) return
  if (authConfig.useRedirectFlow) {
    await msalInstance.logoutRedirect()
  } else {
    await msalInstance.logoutPopup()
  }
}

/**
 * @returns Gets MSAL instance (if you want to use the instance directly)
 */
export function msalGetMsal (): MsalLike | null {
  return msalInstance
}

let msalCreator = async (config: AuthConfig): Promise<MsalLike> => {
  return await PublicClientApplication.createPublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: 'https://login.microsoftonline.com/' + config.tenantId
    },
    cache: {
      cacheLocation: config.cacheLocation ?? 'localStorage'
    }
  }) as unknown as MsalLike
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

async function doLogin (): Promise<AuthenticationResult | undefined> {
  let loginResponse

  if (authConfig.noSso) {
    loginResponse = await doPopupOrRedirectLogin()
  } else {
    try {
      loginResponse = await doSsoLogin()
    } catch (e) {
      console.warn('SSO login failed')
      loginResponse = await doPopupOrRedirectLogin()
    }
  }

  if (loginResponse) {
    msalInstance.setActiveAccount(loginResponse.account)
  }

  return loginResponse
}

async function doSsoLogin (): Promise<AuthenticationResult | undefined> {
  return await msalInstance.ssoSilent({})
}

async function doPopupOrRedirectLogin (): Promise<AuthenticationResult | undefined> {
  if (authConfig.useRedirectFlow) {
    await msalInstance.loginRedirect() // won't go past this line
  } else {
    return await msalInstance.loginPopup()
  }
}

async function doGetAccessToken (tokenRequest: { scopes: string[] }): Promise<AuthenticationResult | undefined> {
  let tokenResponse
  if (authConfig.noAcquireTokenSilent) {
    tokenResponse = await acquireTokenWithRedirectOrPopup(tokenRequest)
  } else {
    try {
      tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest)
    } catch (e) {
      console.error('acquireTokenSilent failed', e)
      tokenResponse = await acquireTokenWithRedirectOrPopup(tokenRequest)

    }
  }
  return tokenResponse
}

async function acquireTokenWithRedirectOrPopup (tokenRequest: PopupRequest | RedirectRequest): Promise<undefined | AuthenticationResult> {
  let tokenResponse
  if (authConfig.useRedirectFlow) {
    await msalInstance.acquireTokenRedirect(tokenRequest) // won't go past this line
  } else {
    tokenResponse = await msalInstance.acquireTokenPopup(tokenRequest)
  }
  return tokenResponse
}

/**
 * For unit testing only
 */
export function msalSetMsalCreator (fnCreator: (config: AuthConfig) => Promise<MsalLike>): void {
  // @ts-expect-error this method is for unit testing only
  msalInstance = null
  msalCreator = fnCreator
}
