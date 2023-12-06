import { AuthenticationResult, IPublicClientApplication, PublicClientApplication } from '@azure/msal-browser'
import { AuthConfig, msalDestroy, msalGetAccessToken, msalGetMsal, msalInit, msalLogin, msalLogout } from "../src/msal-simple-helper";

beforeEach(() => {
   msalDestroy()
});

const mockConfig = {
   clientId: 'mockClientId',
   tenantId: 'mockTenantId'
}

const mockConfig2 = {
   clientId: 'mockClientId2',
   tenantId: 'mockTenantId2'
}

const mockInit = async (unused: AuthConfig) => Promise.resolve(new PublicClientApplication({
   auth: {
      clientId: 'mockClientId',
      authority: 'https://login.microsoftonline.com/' + 'mockTenantId'
   }
}))

 test('logout does nothing when not logged in', async () => {
    await msalLogout()
 })

test('initializing MSAL twice with the same config results in the same instance', async () => {
   const msal1 = await msalInit(mockConfig, mockInit)
   const msal2  = await msalInit(mockConfig, mockInit)
   expect(msal2).toEqual(msal1)
 })

 test('initializing MSAL twice with different config results in an excepion', async () => {
   await msalInit(mockConfig, mockInit)
   expect(async() => {
      await msalInit(mockConfig2, mockInit)
   }).rejects.toThrow()
 })

 test('logging in with the same config is ok', async() => {
   const mockLogin: (msalInstance: IPublicClientApplication, config: AuthConfig) => Promise<AuthenticationResult|undefined> = (msalInstance: IPublicClientApplication, config: AuthConfig) => {
      return null!
   }
   await msalLogin(mockConfig, mockLogin)
   await msalLogin(mockConfig, mockLogin)
 })

 test('logging in with the different config throws exception', async() => {
   const mockLogin: (msalInstance: IPublicClientApplication, config: AuthConfig) => Promise<AuthenticationResult | undefined> = (msalInstance: IPublicClientApplication, config: AuthConfig) => {
      return null!
   }
   await msalLogin(mockConfig, mockLogin)
   expect(async() => {
      await msalLogin(mockConfig2, mockLogin)
   }).rejects.toThrow()
 })

test('redirect flow requires response handler', async () => {
   const mockConfigRedirect = {
      clientId: 'mockClientId2',
      tenantId: 'mockTenantId2',
      flow: 'redirect'
   }
   expect(async() => {
      await msalInit(mockConfigRedirect, mockInit)
   }).rejects.toThrow()
})


 test('redirect login requires response handler', async() => {
   const authConfig: AuthConfig = {
      clientId: 'mockClientId',
      tenantId: 'mockTenantId',
      flow: 'redirect'
   }
   expect( async() => {
      await msalLogin(authConfig)
  }).rejects.toThrow()
 })

 test('get access token throws exception if MSAL not initialized', async() => {
   expect( async() => {
      await msalGetAccessToken()
  }).rejects.toThrow()
 })

 test('logout sets MSAL to null', async() => {
   await msalInit(mockConfig, mockInit)
   expect(msalGetMsal()).toBeDefined()
   await msalLogout(async() => {})
   expect(msalGetMsal()).toBeNull()
 })

 test('msalGetInstance returns the correct instance', async() => {
   const msal = await msalInit(mockConfig, mockInit)
   expect(msalGetMsal()).toBe(msal)
   await msalLogout(async() => {})
   expect(msalGetMsal()).toBeNull()
 })
