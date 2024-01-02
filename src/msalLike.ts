import type { AccountInfo, AuthenticationResult, EndSessionPopupRequest, EndSessionRequest, PopupRequest, RedirectRequest, SilentRequest, SsoSilentRequest } from '@azure/msal-browser'

export interface MsalLike {
  loginRedirect: (request?: RedirectRequest) => Promise<void>
  handleRedirectPromise: (hash?: string) => Promise<AuthenticationResult | null>
  setActiveAccount: (account: AccountInfo | null) => void
  getAllAccounts: () => AccountInfo[]
  loginPopup: (request?: PopupRequest) => Promise<AuthenticationResult>
  acquireTokenSilent: (silentRequest: SilentRequest) => Promise<AuthenticationResult>
  acquireTokenPopup: (request: PopupRequest) => Promise<AuthenticationResult>
  acquireTokenRedirect: (request: RedirectRequest) => Promise<void>
  logoutRedirect: (logoutRequest?: EndSessionRequest) => Promise<void>
  logoutPopup: (logoutRequest?: EndSessionPopupRequest) => Promise<void>
  ssoSilent: (request: SsoSilentRequest) => Promise<AuthenticationResult>
}
