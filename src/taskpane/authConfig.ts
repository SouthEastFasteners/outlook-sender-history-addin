// Fill these in from your Entra ID app registration (see README Step 1).
export const msalConfig = {
  auth: {
    // Application (client) ID from the app registration overview page
    clientId: "63839c75-0eb2-46dd-9fed-c7faab5ac2c4",
    // Directory (tenant) ID — using your specific tenant (rather than
    // "common") is recommended for an internal-only add-in.
    authority: "https://login.microsoftonline.com/dca2f5c2-445f-4b37-9783-6fbefd0992d4",
  },
};

// Delegated Graph scopes the add-in requests. Must match what was
// granted admin consent in the app registration's API permissions.
export const graphScopes = ["Mail.Read", "User.Read"];
