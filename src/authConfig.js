export const msalConfig = {
    auth: {
        clientId: "499f7779-5088-4a08-b19d-2a23edcf2640",

        authority:
            "https://login.microsoftonline.com/828ad09d-e51b-4a6a-92e3-6e0737a5c395",

        redirectUri: "http://localhost:3000/"
    },

    cache: {
        cacheLocation: "sessionStorage"
    }
};

export const loginRequest = {
    scopes: ["openid", "profile", "email"]
};