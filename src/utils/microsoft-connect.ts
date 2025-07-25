import axios from "axios";
import * as qs from "qs";
import type { MicrosoftConfig } from "../types/config";

type CachedToken = {
  token: string;
  expiresAt: number;
};

const microsoftTokenCache = new Map<string, CachedToken>();

export async function generateMicrosoftAccessToken(
  dto: MicrosoftConfig,
  log: boolean = false,
): Promise<Object> {
  const missingParams = [];
  if (!dto.tenentId) missingParams.push("tenentId");
  if (!dto.clientId) missingParams.push("clientId");
  if (!dto.clientSecret) missingParams.push("clientSecret");
  if (!dto.scope) missingParams.push("scope");
  if (!dto.grandType) missingParams.push("grandType");
  if (missingParams.length > 0) {
    throw new Error(
      `[kas-on-cloud]: Missing required Microsoft configuration parameters for generate access token: ${missingParams.join(", ")}`,
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${dto.tenentId}/oauth2/v2.0/token`;

  const clientId = dto.clientId;
  if (!clientId) {
    throw new Error(
      "[kas-on-cloud]: Client ID is required to generate access token",
    );
  }
  const clientSecret = dto.clientSecret;

  const cached = microsoftTokenCache.get(clientId);
  const now = Date.now();

  if (cached && now < cached.expiresAt) {
    if (log) console.log("✅[kas-on-cloud]: Token from cache for: ", clientId);
    return cached.token;
  }

  const data = qs.stringify({
    grant_type: dto.grandType,
    scope: dto.scope,
  });

  const response = await axios.post(tokenUrl, data, {
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    auth: {
      username: clientId || "",
      password: clientSecret || "",
    },
  });

  const token = response.data.access_token;
  const expiresIn = response.data.expires_in;

  const expiresAt = now + expiresIn * 1000;

  microsoftTokenCache.set(clientId, {
    token,
    expiresAt,
  });

  if (log) {
    console.log("[kas-on-cloud]: Microsoft Access Token Request:", {
      url: tokenUrl,
      data: data,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      token: response.data.access_token,
      expires_in: response.data.expires_in,
      ext_expires_in: response.data.ext_expires_in,
    });
  }

  return {
    accessToken: response.data.access_token,
    expiresIn: response.data.expires_in,
    extExpiresIn: response.data.ext_expires_in,
  };
}