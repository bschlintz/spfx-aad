import { AzureFunction, Context, HttpRequest } from "@azure/functions";
import AADTokenManager from './aad-token-manager';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
  context.log('HTTP trigger function processed a request.');

  // Read AUTHORIZATION header
  const header = req.headers['authorization'];
  const headerParts = header ? header.split(' ') : [];
  const authorizationToken = headerParts.length === 2 ? headerParts[1] : '';

  // ENSURE we have a token
  if (!authorizationToken) {
    return unauthorized("missing authorization header");
  }

  // Use AAD Token Manager to VALIDATE
  const aadManager = new AADTokenManager();

  // ENSURE we can read the tenant ID from the token
  const tenantId = aadManager.getTenantId(authorizationToken);
  if (!tenantId) {
    return unauthorized("invalid authorization header - unable to read tenant id");
  }

  // FETCH Open ID Config for tenant
  const openIdConfig = await aadManager.requestOpenIdConfig(tenantId);

  // DOWNLOAD token signing certificates from Microsoft for this tenant
  const tokenSigningCertificates = await aadManager.requestSigningCertificates(openIdConfig.jwks_uri, null);

  // VERIFY token against token signing certificates
  try {
    const decodedToken = await aadManager.verify(authorizationToken, tokenSigningCertificates, null);
    return success(decodedToken);
  }
  catch (error) {
    return unauthorized(`unable to verify token - ${error}`);
  }
};

const unauthorized = (message: string): any => {
  return {
    status: 401,
    body: JSON.stringify({ message: `Unauthorized: ${message}` })
  };
}

const success = (decodedToken: string): any => {
  return {
    status: 200,
    body: JSON.stringify({ message: "Authorized: Yay!", decodedToken })
  };
}

export default httpTrigger;
