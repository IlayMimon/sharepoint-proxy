import { Request, Response, Express } from "express";
import express from "express";
import { Configuration, PublicClientApplication } from "@azure/msal-node";
import { config as dotenvConfig } from "dotenv";

dotenvConfig();

const proxy: Express = express();
proxy.use(express.json());
proxy.use(express.urlencoded({ extended: true }));
const port = 3000;

let token: string | undefined = undefined;
let tokenExpiresOn: number | undefined = undefined;

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
    const result = await clientApp.acquireTokenByDeviceCode({
      scopes: [`${process.env.SHAREPOINT_URL}/.default`],
      deviceCodeCallback: (response) => {
        console.log(response.message);
      },
    });
    token = result?.accessToken;
    tokenExpiresOn = result?.expiresOn ? new Date(result.expiresOn).getTime() : undefined;
    return token;
  }
};

proxy.all("/*all", async (req: Request, res: Response) => {
  const token = await getToken();
  if (!token) return void res.status(500).send("Token fetch failed.");

  const sharePointUrl = `${process.env.SHAREPOINT_URL}${req.originalUrl}`;
  console.log("request initiated with endpoint: " + sharePointUrl);
  const hasBody: boolean =
    !["GET", "HEAD"].includes(req.method.toUpperCase()) &&
    req.body &&
    Object.keys(req.body).length > 0;
  try {
    const spResponse = await fetch(sharePointUrl, {
      ...req,
      method: req.method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=verbose",
        ...(req.headers as { [key: string]: string }),
      },
      body: hasBody ? JSON.stringify(req.body) : undefined,
    });
    
    const data = await spResponse.json();
    res.status(spResponse.status).send(data);
  } catch (error) {
    console.error("Proxy error:", error);
    res.status(500).send("Proxy request failed");
  }
});
