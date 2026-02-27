import type { Configuration, RedirectRequest } from "@azure/msal-browser";

export const msalConfig: Configuration = {
    auth: {
        clientId:    "3a50a290-29f1-4e09-bc84-fd36913dca59",
        authority:   "https://login.microsoftonline.com/pecportaldev.onmicrosoft.com",
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: "localStorage",
    },
};

export const loginRequest: RedirectRequest = {
    scopes: [
        "openid",
        "profile",
        "email",
        "User.Read",
    ],
};

export const graphConfig = {
    baseUrl:  "https://graph.microsoft.com/v1.0",
    scopes:   ["https://graph.microsoft.com/.default"],
    tenantId: "pecportaldev.onmicrosoft.com",
};
