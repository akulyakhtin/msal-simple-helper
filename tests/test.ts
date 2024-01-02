import { AccountInfo, AuthenticationResult, EndSessionPopupRequest, EndSessionRequest, PopupRequest, RedirectRequest, SilentRequest, SsoSilentRequest } from '@azure/msal-browser'
import { msalGetAccessToken, msalLogin, msalSetMsalCreator } from "../src/msal-simple-helper";
import { MsalLike } from '../src/msalLike';

beforeEach(() => {
   msalSetMsalCreator(() => Promise.resolve(MSAL_STUB))
});

const mockConfig = {
   clientId: 'mockClientId',
   tenantId: 'mockTenantId'
}

class MsalStub implements MsalLike {
   loginPopup(request?: PopupRequest | undefined): Promise<AuthenticationResult> {
      console.log('loginPopup stub')
      return Promise.resolve({} as unknown as AuthenticationResult)
   }
   loginRedirect(request?: RedirectRequest | undefined): Promise<void> {
      console.log('loginRedirect stub')
      return Promise.resolve()
   }
   handleRedirectPromise(hash?: string | undefined): Promise<AuthenticationResult | null> {
      console.log('handleRedirectPromise stub')
      return Promise.resolve({} as unknown as AuthenticationResult)
   }
   setActiveAccount(account: AccountInfo | null): void {
      console.log('setActiveAccount stub')
   }
   getAllAccounts(): AccountInfo[] {
      console.log('getAllAccounts stub')
      return []
   }
   acquireTokenSilent(silentRequest: SilentRequest): Promise<AuthenticationResult> {
      console.log('acquireTokenSilent stub')
      return Promise.resolve({} as unknown as AuthenticationResult)
   }
   acquireTokenPopup(request: PopupRequest): Promise<AuthenticationResult> {
      console.log('acquireTokenPopup stub')
      return Promise.resolve({} as unknown as AuthenticationResult)
   }
   acquireTokenRedirect(request: RedirectRequest): Promise<void> {
      console.log('acquireTokenRedirect stub')
      return Promise.resolve()
   }
   logoutRedirect(logoutRequest?: EndSessionRequest | undefined): Promise<void> {
      console.log('loginRedirect stub')
      return Promise.resolve()
   }
   logoutPopup(logoutRequest?: EndSessionPopupRequest | undefined): Promise<void> {
      throw new Error('Method not implemented.');
   }
   ssoSilent(request: SsoSilentRequest): Promise<AuthenticationResult> {
      console.log('ssoSilent stub')
      return Promise.resolve({} as unknown as AuthenticationResult)
   }
}

const MSAL_STUB = new MsalStub()

test('when logging in first try sso', async () => {

   MSAL_STUB.ssoSilent = jest.fn(() => Promise.resolve({ account: "mockAccount" } as unknown as AuthenticationResult))

   await msalLogin(mockConfig)
   expect(MSAL_STUB.ssoSilent).toHaveBeenCalled()
})

test('when noSso specified sso login not performed', async () => {
   MSAL_STUB.ssoSilent = jest.fn(() => Promise.resolve({ account: "mockAccount" } as unknown as AuthenticationResult))
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      noSso: true
   })
   expect(MSAL_STUB.ssoSilent).not.toHaveBeenCalled()
})

test('if sso fails popup login will be performed', async () => {
   MSAL_STUB.ssoSilent = (request: SsoSilentRequest) => { throw Error('sso failed') }
   MSAL_STUB.loginPopup = jest.fn(() => Promise.resolve({ account: "mockAccount" } as unknown as AuthenticationResult))
   await msalLogin(mockConfig)
   expect(MSAL_STUB.loginPopup).toHaveBeenCalled()
})

test('when noSso specified popup login will be performed', async () => {
   MSAL_STUB.ssoSilent = (request: SsoSilentRequest) => { throw Error('sso failed') }
   MSAL_STUB.loginPopup = jest.fn(() => Promise.resolve({ account: "mockAccount" } as unknown as AuthenticationResult))
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      noSso: true
   })
   expect(MSAL_STUB.loginPopup).toHaveBeenCalled()
})

test('if sso fails and redirect is specified then redirect login will be performed', async () => {
   MSAL_STUB.ssoSilent = (request: SsoSilentRequest) => { throw Error('sso failed') }
   MSAL_STUB.loginRedirect = jest.fn(() => Promise.resolve())
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      useRedirectFlow: true,
      redirectResponseHandler: (response) => { }
   })
   expect(MSAL_STUB.loginRedirect).toHaveBeenCalled()
})

test('if noSso specified and redirect is specified then redirect login will be performed', async () => {
   MSAL_STUB.loginRedirect = jest.fn(() => Promise.resolve())
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      noSso: true,
      useRedirectFlow: true,
      redirectResponseHandler: (response) => { }
   })
   expect(MSAL_STUB.loginRedirect).toHaveBeenCalled()
})

test('if redirect flow is requested then redirectResponseHandler should be given', async () => {
   expect(async () => {
      await msalLogin({
         clientId: 'mockClientId',
         tenantId: 'mockTenantId',
         useRedirectFlow: true,
      })
   }).rejects.toThrow()
})

test('if redirect flow is requested and noSso is specified then redirectResponseHandler should be given', async () => {
   expect(async () => {
      await msalLogin({
         clientId: 'mockClientId',
         tenantId: 'mockTenantId',
         useRedirectFlow: true,
         noSso: true
      })
   }).rejects.toThrow()
})

test('when getting access token, acquire token silent gets called first', async() => {
   MSAL_STUB.acquireTokenSilent = jest.fn()
   await msalLogin(mockConfig)
   await msalGetAccessToken()
   expect(MSAL_STUB.acquireTokenSilent).toHaveBeenCalled()
})

test('when getting access token, if acquire token silent fails then popup gets called', async() => {
   MSAL_STUB.acquireTokenSilent = () =>  { throw Error('acquireTokenSilent failed')}
   MSAL_STUB.acquireTokenPopup = jest.fn()
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId'
   })
   await msalGetAccessToken()
   expect(MSAL_STUB.acquireTokenPopup).toHaveBeenCalled()
})

test('when getting access token with redirect flow, if acquire token silent fails then redirect gets called', async() => {
   MSAL_STUB.acquireTokenSilent = () =>  { throw Error('acquireTokenSilent failed')}
   MSAL_STUB.acquireTokenRedirect = jest.fn()
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      useRedirectFlow: true,
      redirectResponseHandler: (response) => { }
   })
   await msalGetAccessToken()
   expect(MSAL_STUB.acquireTokenRedirect).toHaveBeenCalled()
})

test('when noAcquirTokenSilent specified, token is not acquired silently', async() => {
   MSAL_STUB.acquireTokenSilent = jest.fn()
   await msalLogin({
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      noAcquireTokenSilent: true
   })
   await msalGetAccessToken()
   expect(MSAL_STUB.acquireTokenSilent).not.toHaveBeenCalled()
})
