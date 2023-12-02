import { AuthenticationResult, PublicClientApplication } from "@azure/msal-browser"

export type AuthConfig = {
    tenantId: string,
    clientId: string,
    scopes?: string[],
    cacheLocation?: string
}

let msalInstance: PublicClientApplication | null = null
let authConfig: AuthConfig | null = null

async function msalInit(config: AuthConfig): Promise<PublicClientApplication> {
    const msal = new PublicClientApplication({
        auth: {
            clientId: config.clientId,
            authority: 'https://login.microsoftonline.com/' + config.tenantId,
        },
        cache: {
            cacheLocation: config.cacheLocation || 'localStorage'
        }
    })            
    await msal.initialize()
    return msal;
}

function checkConfig(config: AuthConfig) {
    if (!config) throw Error('Auth config can not be null')
    if (!config.tenantId) throw Error('Tenant ID can not be null')
    if (!config.clientId) throw Error('Client ID can not be null')
}

function configsDiffer(cf1: AuthConfig, cf2: AuthConfig): boolean {
    return cf1.tenantId != cf2.tenantId ||
        cf1.clientId != cf2.clientId ||
        cf1.scopes != cf2.scopes
}

export async function msalLogin(config: AuthConfig): Promise<AuthenticationResult> {
    if (msalInstance) {
        if (authConfig && configsDiffer(config, authConfig)) {
            throw new Error('Already logged in with different config. Please. log out, first.');
        }
    }
    checkConfig(config);

    try {
        msalInstance = await msalInit(config);
        let loginResponse = null
        try {
            loginResponse = await msalInstance.ssoSilent({})
        } catch(e) {
            console.error('ssoSilent failed', e);
            loginResponse = await msalInstance.loginPopup();
        }
        msalInstance.setActiveAccount(loginResponse.account);
        authConfig = config;
        return loginResponse;
    } catch (e) {
        msalInstance = null;
        throw e;
    }
}

export async function msalGetAccessToken() {
    if(!msalInstance) throw Error('Please, login first.')
    const tokenRequest = {
        scopes: authConfig?.scopes || []
    }

    let tokenResponse;
    try {
        tokenResponse = await msalInstance?.acquireTokenSilent(tokenRequest)
    } catch(e) {
        console.error('acquireTokenSilent failed', e);
        tokenResponse = await msalInstance.acquireTokenPopup(tokenRequest)
    }
    return tokenResponse;
}

export async function msalLogout() {
    console.log('msal logout')
    if (!msalInstance) return;
    await msalInstance.logoutPopup();
    msalInstance = null;
    authConfig = null;
}

export function msalGetMsal(): PublicClientApplication | null {
    return msalInstance;
}