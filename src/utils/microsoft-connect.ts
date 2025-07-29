import axios from "axios";
import qs from "qs";
import type { MicrosoftConfig } from "../types/config";

type CachedToken = {
  token: string;
  expiresAt: number;
};

type MicrosoftAccessTokenResponse = {
  accessToken: string;
  expiresIn: number;
  extExpiresIn: number;
};

const tokenCache = new Map<string, CachedToken>();

export async function generateMicrosoftAccessToken(
  config: MicrosoftConfig,
  isShowLog = false,
): Promise<MicrosoftAccessTokenResponse> {
  const { tenentId, clientId, clientSecret, scope, grandType } = config;

  const missingParams = Object.entries({
    tenentId,
    clientId,
    clientSecret,
    scope,
  })
    .filter(([_, v]) => !v)
    .map(([k]) => k);

  if (missingParams.length > 0) {
    throw new Error(
      `[kas-on-cloud]: Missing required Microsoft config params: ${missingParams.join(", ")}`,
    );
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenentId}/oauth2/v2.0/token`;
  const now = Date.now();

  const cached = tokenCache.get(clientId!);
  if (cached && now < cached.expiresAt) {
    if (isShowLog) console.log("✅[kas-on-cloud]: Token from cache:", clientId);
    return {
      accessToken: cached.token,
      expiresIn: (cached.expiresAt - now) / 1000,
      extExpiresIn: (cached.expiresAt - now) / 1000,
    };
  }

  const data = qs.stringify({
    grant_type: grandType || "client_credentials",
    scope,
  });

  const headers = {
    "Content-Type": "application/x-www-form-urlencoded",
  };

  const res = await axios.post(tokenUrl, data, {
    headers,
    auth: {
      username: clientId!,
      password: clientSecret!,
    },
  });

  const { access_token, expires_in, ext_expires_in } = res.data || {};

  if (!access_token) {
    throw new Error(
      `[kas-on-cloud]: Token not found in response:\n${JSON.stringify(res.data, null, 2)}`,
    );
  }

  const expiresAt = now + expires_in * 1000;

  tokenCache.set(clientId!, {
    token: access_token,
    expiresAt,
  });

  if (isShowLog) {
    console.log("[kas-on-cloud]: Access Token Info", {
      tokenUrl,
      scope,
      token: access_token,
      expiresIn: expires_in,
      extExpiresIn: ext_expires_in,
    });
  }

  return {
    accessToken: access_token,
    expiresIn: expires_in,
    extExpiresIn: ext_expires_in,
  };
}