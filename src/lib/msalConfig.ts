import { Configuration, LogLevel } from "@azure/msal-browser";

const CLIENT_ID = "9d68763b-80bc-494b-9f06-8d109490d604";
const TENANT_ID = "827dd3b6-a0ce-4577-a71f-06ac510929a7";

export const msalConfig: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === LogLevel.Error) console.error(message);
      },
      logLevel: LogLevel.Error,
    },
  },
};

export const loginRequest = {
  scopes: ["User.Read", "GroupMember.Read.All"],
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
};

export const groups = [
  {
    id: "c214659a-a666-4a17-94a8-2c649dcbd59c",
    name: "Amministrazione",
    endpoint: "https://graph.microsoft.com/v1.0/groups/c214659a-a666-4a17-94a8-2c649dcbd59c/members",
  },
];
