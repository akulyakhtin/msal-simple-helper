 import { msalLogout } from "../src/msal-simple-helper";

 test('logout does nothing when not logged in', async () => {
    await msalLogout()
 })