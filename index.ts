import { Request, Response, Express } from "express";
import express from "express";
import { Configuration, PublicClientApplication } from "@azure/msal-node";
import { config as dotenvConfig } from "dotenv";

dotenvConfig();

const proxy: Express = express();
// Parse all request bodies as raw buffers so we can proxy file uploads (multipart/binary) without losing data
proxy.use(express.raw({ type: () => true, limit: "200mb" }));
const port = 3000;

let token: string | undefined = undefined;
let tokenExpiresOn: number | undefined = undefined;
let timeSinceLastTokenFetchAttempt: number | undefined = undefined;

proxy.listen(port, async () => {
  await getToken();
  console.log("proxy open on port ", port);
});

const config: Configuration = {
  auth: {
    clientId: process.env.CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
  },
};

const clientApp = new PublicClientApplication(config);

const getToken = async (): Promise<string | undefined> => {
  const now = Date.now();
  if (token && tokenExpiresOn && now < tokenExpiresOn - 60 * 1000) {
    return token;
  } else {
    // if an attempt to fetch the token was made any time in the last 30 seconds, don't send
    if (
      timeSinceLastTokenFetchAttempt &&
      now - timeSinceLastTokenFetchAttempt < 30 * 1000
    )
      return token;
    const result = await clientApp.acquireTokenByDeviceCode({
      scopes: [`${process.env.SHAREPOINT_URL}/.default`],
      deviceCodeCallback: (response) => {
        console.log(response.message);
      },
    });
    token = result?.accessToken;
    timeSinceLastTokenFetchAttempt = now;
    tokenExpiresOn = result?.expiresOn
      ? new Date(result.expiresOn).getTime()
      : undefined;
    return token;
  }
};

proxy.all("/*all", async (req: Request, res: Response) => {
  const token = await getToken();
  if (!token)
    return void res
      .status(401)
      .send("Token fetch failed. Try to refresh Proxy");

  const sharePointUrl = `${process.env.SHAREPOINT_URL}${req.originalUrl}`;
  console.log("request initiated with endpoint: " + sharePointUrl);

  const hasBody =
    !["GET", "HEAD"].includes(req.method.toUpperCase()) &&
    req.body &&
    ((Buffer.isBuffer(req.body) && req.body.length > 0) ||
      (typeof req.body === "string" && req.body.length > 0));

  const forwardedHeaders: Record<string, string> = {};
  for (const [key, value] of Object.entries(req.headers)) {
    if (value === undefined) continue;
    forwardedHeaders[key] = Array.isArray(value) ? value.join(",") : value;
  }
  // Avoid host conflicts when proxying
  delete forwardedHeaders.host;

  try {
    const forwardedWithAuth: Record<string, string> = {
      ...forwardedHeaders,
      Authorization: `Bearer ${token}`,
    };
    if (!forwardedWithAuth.accept) {
      forwardedWithAuth.accept = "application/json;odata=verbose";
    }

    const spResponse = await fetch(sharePointUrl, {
      method: req.method,
      headers: forwardedWithAuth,
      body: hasBody ? req.body : undefined,
    });

    res.status(spResponse.status);

    const contentType = spResponse.headers.get("content-type") ?? "";

    if (contentType.includes("application/json")) {
      res.json(await spResponse.json());
    } else {
      const responseBuffer = Buffer.from(await spResponse.arrayBuffer());
      if (contentType) {
        res.setHeader("content-type", contentType);
      }
      res.send(responseBuffer);
    }
  } catch (error) {
    console.error("Proxy error:", error);
    res.status(500).send("Proxy request failed");
  }
});
