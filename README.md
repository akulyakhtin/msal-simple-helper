# MSAL-SIMPLE_HELPER
Simple wrapper around MSAL4JS allowing for just a single login 
(which in many cases is enough)

Usage:
```
// Prepare your Azure Entra ID data
const authConfig: AuthConfig = {
      tenantId: yourTenantId,
      clientId: yourClientId,
      // scopes: [your_scope_1, your_scope_2]
    }
    
// Log in (currently, tries to login silently then performs popup login if silent login fails)    
const loginResult = await msalLogin(authConfig)

// msalLogin returns login response from MSAL, so you can use the data as you want to
console.log('Logged in username:', loginResult.acccount.name)

// Get access token
const tokenResponse = await msalGetAccessToken();

// msalGetAccessToken returns access token response from MSAL
// so you can pass the access token like this:
// console.log('Authorization: Bearer ', tokenResponse.accessToken)

// Finally, logout
await msalLogout()
```
