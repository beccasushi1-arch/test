import { createExtension, createNodeDescriptor, INodeFunctionBaseParams } from "@cognigy/extension-tools";
import axios, { AxiosRequestConfig } from "axios";

// ─── Helper: fetch token from Azure Key Vault ────────────────────────────────

async function getKeyVaultSecret(
  vaultUrl: string,
  secretName: string,
  tenantId: string,
  clientId: string,
  clientSecret: string
): Promise<string> {
  // Get AAD token
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const tokenParams = new URLSearchParams();
  tokenParams.append("grant_type", "client_credentials");
  tokenParams.append("client_id", clientId);
  tokenParams.append("client_secret", clientSecret);
  tokenParams.append("scope", "https://vault.azure.net/.default");

  const tokenResponse = await axios.post(tokenUrl, tokenParams.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" }
  });
  const accessToken = tokenResponse.data.access_token;

  // Get secret from Key Vault
  const secretUrl = `${vaultUrl}/secrets/${secretName}?api-version=7.4`;
  const secretResponse = await axios.get(secretUrl, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });

  return secretResponse.data.value;
}

// ─── Helper: make one API call ───────────────────────────────────────────────

async function makeApiCall(
  apiUrl: string,
  method: string,
  body: any,
  queryParams: Record<string, string>,
  bearerToken: string,
  additionalHeaders: Record<string, string>
): Promise<any> {
  const config: AxiosRequestConfig = {
    method: method.toLowerCase() as any,
    url: apiUrl,
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
      ...additionalHeaders
    },
    params: queryParams || {}
  };

  if (["post", "put", "patch"].includes(method.toLowerCase()) && body) {
    config.data = typeof body === "string" ? JSON.parse(body) : body;
  }

  const response = await axios(config);
  return response.data;
}

// ─── Node: AZ Dual API Parallel Request ─────────────────────────────────────

export const azDualApiNode = createNodeDescriptor({
  type: "azDualApiParallel",
  defaultLabel: "AZ Dual API Parallel Request",
  summary: "Fires two authenticated API calls in parallel using Azure Key Vault credentials",

  fields: [
    // ── Key Vault Connection ──────────────────────────────────────────────
    {
      key: "connection",
      label: "AZ Key Vault Connection",
      type: "connection",
      params: {
        connectionType: "keyVaultConnection",
        required: true
      }
    },
    {
      key: "dataConnection",
      label: "AZ Key Vault Data Configuration",
      type: "connection",
      params: {
        connectionType: "keyVaultData",
        required: true
      }
    },

    // ── API Call 1 ────────────────────────────────────────────────────────
    {
      key: "apiUrl1",
      label: "API 1 - URL",
      type: "cognigyText",
      defaultValue: "",
      params: { required: true }
    },
    {
      key: "method1",
      label: "API 1 - Method",
      type: "select",
      defaultValue: "POST",
      params: {
        options: [
          { label: "GET", value: "GET" },
          { label: "POST", value: "POST" },
          { label: "PUT", value: "PUT" },
          { label: "PATCH", value: "PATCH" },
          { label: "DELETE", value: "DELETE" }
        ]
      }
    },
    {
      key: "queryParams1",
      label: "API 1 - Query Params (JSON)",
      type: "json",
      defaultValue: {}
    },
    {
      key: "body1",
      label: "API 1 - Body Data",
      type: "cognigyText",
      defaultValue: ""
    },
    {
      key: "resultKey1",
      label: "API 1 - Store result in context as",
      type: "cognigyText",
      defaultValue: "result1"
    },

    // ── API Call 2 ────────────────────────────────────────────────────────
    {
      key: "apiUrl2",
      label: "API 2 - URL",
      type: "cognigyText",
      defaultValue: "",
      params: { required: true }
    },
    {
      key: "method2",
      label: "API 2 - Method",
      type: "select",
      defaultValue: "POST",
      params: {
        options: [
          { label: "GET", value: "GET" },
          { label: "POST", value: "POST" },
          { label: "PUT", value: "PUT" },
          { label: "PATCH", value: "PATCH" },
          { label: "DELETE", value: "DELETE" }
        ]
      }
    },
    {
      key: "queryParams2",
      label: "API 2 - Query Params (JSON)",
      type: "json",
      defaultValue: {}
    },
    {
      key: "body2",
      label: "API 2 - Body Data",
      type: "cognigyText",
      defaultValue: ""
    },
    {
      key: "resultKey2",
      label: "API 2 - Store result in context as",
      type: "cognigyText",
      defaultValue: "result2"
    },

    // ── Error handling ────────────────────────────────────────────────────
    {
      key: "errorKey",
      label: "Store errors in context as",
      type: "cognigyText",
      defaultValue: "dualApiError"
    },
    {
      key: "stopOnError",
      label: "Stop flow on error",
      type: "toggle",
      defaultValue: false
    }
  ],

  sections: [
    {
      key: "authSection",
      label: "Authentication",
      defaultCollapsed: false,
      fields: ["connection", "dataConnection"]
    },
    {
      key: "api1Section",
      label: "API Call 1",
      defaultCollapsed: false,
      fields: ["apiUrl1", "method1", "queryParams1", "body1", "resultKey1"]
    },
    {
      key: "api2Section",
      label: "API Call 2",
      defaultCollapsed: false,
      fields: ["apiUrl2", "method2", "queryParams2", "body2", "resultKey2"]
    },
    {
      key: "errorSection",
      label: "Error Handling",
      defaultCollapsed: true,
      fields: ["errorKey", "stopOnError"]
    }
  ],

  form: [
    { type: "section", key: "authSection" },
    { type: "section", key: "api1Section" },
    { type: "section", key: "api2Section" },
    { type: "section", key: "errorSection" }
  ],

  appearance: {
    color: "#003781" // Allianz blue
  },

  function: async ({ cognigy, config }: INodeFunctionBaseParams) => {
    const { api, context } = cognigy;
    const {
      connection,
      dataConnection,
      apiUrl1, method1, queryParams1, body1, resultKey1,
      apiUrl2, method2, queryParams2, body2, resultKey2,
      errorKey,
      stopOnError
    } = config as any;

    try {
      // ── 1. Pull credentials from Key Vault connections ──────────────────
      const { tenantId, clientId, clientSecret, vaultUrl } = connection;
      const { secretName } = dataConnection;

      // Fetch the bearer token secret from Key Vault
      const bearerToken = await getKeyVaultSecret(
        vaultUrl,
        secretName,
        tenantId,
        clientId,
        clientSecret
      );

      // ── 2. Fire both API calls in parallel ──────────────────────────────
      const [response1, response2] = await Promise.all([
        makeApiCall(apiUrl1, method1, body1, queryParams1 || {}, bearerToken, {}),
        makeApiCall(apiUrl2, method2, body2, queryParams2 || {}, bearerToken, {})
      ]);

      // ── 3. Store results in context ─────────────────────────────────────
      api.setContext(resultKey1 || "result1", response1);
      api.setContext(resultKey2 || "result2", response2);

      // Clear any previous error
      api.setContext(errorKey || "dualApiError", null);

    } catch (error: any) {
      const errorPayload = {
        message: error?.message || "Unknown error",
        status: error?.response?.status || null,
        data: error?.response?.data || null
      };

      api.setContext(errorKey || "dualApiError", errorPayload);

      if (stopOnError) {
        api.log("error", `AZ Dual API Error: ${JSON.stringify(errorPayload)}`);
      }
    }
  }
});

// ─── Connection definitions ──────────────────────────────────────────────────

export const keyVaultConnection = {
  type: "keyVaultConnection",
  label: "AZ Key Vault Connection",
  fields: [
    { fieldName: "vaultUrl", label: "Key Vault URL", fieldType: "url" },
    { fieldName: "tenantId", label: "Tenant ID", fieldType: "text" },
    { fieldName: "clientId", label: "Client ID", fieldType: "text" },
    { fieldName: "clientSecret", label: "Client Secret", fieldType: "secret" }
  ]
};

export const keyVaultDataConnection = {
  type: "keyVaultData",
  label: "AZ Key Vault Data Configuration",
  fields: [
    { fieldName: "secretName", label: "Secret Name", fieldType: "text" }
  ]
};

// ─── Export extension ────────────────────────────────────────────────────────

export default createExtension({
  nodes: [azDualApiNode],
  connections: [keyVaultConnection, keyVaultDataConnection]
});
