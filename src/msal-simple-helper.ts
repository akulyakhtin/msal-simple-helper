import type { AuthenticationResult, IPublicClientApplication } from '@azure/msal-browser'
import { PublicClientApplication } from '@azure/msal-browser'

export interface AuthConfig {
  tenantId: string
  clientId: string
  scopes?: string[]
  cacheLocation?: string
}

let msalInstance: IPublicClientApplication | null = null
let authConfig: AuthConfig | null = null

async function msalInit (config: AuthConfig): Promise<IPublicClientApplication> {
  const msal = await PublicClientApplication.createPublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: 'https://login.microsoftonline.com/' + config.tenantId
    },
    cache: {
      cacheLocation: config.cacheLocation ?? 'localStorage'
    }
  })
  authConfig = config
  return msal
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

export async function msalInitForRedirect (config: AuthConfig, loginResponseHandler: (loginResponse: AuthenticationResult | null) => void): Promise<void> {
  msalInstance = await msalInit(config)
  msalInstance.handleRedirectPromise().then(handleRedirectResponse)
    .then((resp) => { if (resp !== null) { loginResponseHandler(resp) } })
    .catch(e => { console.error(e); throw e })
}

function configsDiffer (cf1: AuthConfig, cf2: AuthConfig): boolean {
  return cf1.tenantId !== cf2.tenantId ||
    cf1.clientId !== cf2.clientId ||
    cf1.scopes !== cf2.scopes
}

export async function msalLogin (config: AuthConfig): Promise<AuthenticationResult> {
  if (msalInstance === null) {
    if ((authConfig !== null) && configsDiffer(config, authConfig)) {
      throw new Error('Already logged in with different config. Please. log out, first.')
    } else {
      console.warn('Already logged in with the same config.')
    }
  }

  try {
    msalInstance = await msalInit(config)
    let loginResponse = null
    try {
      loginResponse = await msalInstance.ssoSilent({})
    } catch (e) {
      console.error('ssoSilent failed', e)
      loginResponse = await msalInstance.loginPopup()
    }
    msalInstance.setActiveAccount(loginResponse.account)
    return loginResponse
  } catch (e) {
    msalInstance = null
    throw e
  }
}

export async function msalLoginPopup (config: AuthConfig): Promise<AuthenticationResult> {
  return await msalLogin(config)
}

export async function msalLoginRedirect (): Promise<undefined | AuthenticationResult> {
  if (msalInstance === null) throw new Error('msalInstance should not be null. Did you call msalInitForRedirect before calling login?')
  let loginResponse
  try {
    loginResponse = await msalInstance.ssoSilent({})
    msalInstance.setActiveAccount(loginResponse.account)
    return await Promise.resolve(loginResponse)
  } catch (e) {
    console.error('ssoSilent failed', e)
    await msalInstance.loginRedirect()
  }
}

export async function msalGetAccessToken (): Promise<AuthenticationResult> {
  if (msalInstance === null) throw Error('Please, login first.')
  const tokenRequest = {
    scopes: authConfig?.scopes ?? []
  }

  let tokenResponse
  try {
    tokenResponse = await msalInstance?.acquireTokenSilent(tokenRequest)
  } catch (e) {
    console.error('acquireTokenSilent failed', e)
    tokenResponse = await msalInstance.acquireTokenPopup(tokenRequest)
  }
  return tokenResponse
}

export async function msalLogout (): Promise<void> {
  if (msalInstance == null) return
  await msalInstance.logoutPopup()
  msalInstance = null
  authConfig = null
}

export function msalGetMsal (): IPublicClientApplication | null {
  return msalInstance
}
