export const msalConfig = {
    auth: {
        clientId: "86c59355-8f99-4081-96c9-39db6e13ca08",

        authority:
            "https://login.microsoftonline.com/ba9c3522-adab-4c57-ac78-dc6131594e07",

        redirectUri: "http://localhost:3000/"
    },

    cache: {
        cacheLocation: "sessionStorage"
    }
};

export const loginRequest = {
    scopes: ["openid", "profile", "email"]
};

export const apiRequest = {
    scopes: [
        "api://c639b2a3-10c3-4611-b3c8-8acb71c2afc6/access_as_user"
    ]
};