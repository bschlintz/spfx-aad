import * as jsonwebtoken from 'jsonwebtoken';
import * as fetch from 'node-fetch';

class AADTokenManager {

  private static ConvertCertificateToBeOpenSSLCompatible(cert): string {
    //Certificate must be in this specific format or else the function won't accept it
    var beginCert = "-----BEGIN CERTIFICATE-----";
    var endCert = "-----END CERTIFICATE-----";

    cert = cert.replace("\n", "");
    cert = cert.replace(beginCert, "");
    cert = cert.replace(endCert, "");

    var result = beginCert;
    while (cert.length > 0) {

      if (cert.length > 64) {
        result += "\n" + cert.substring(0, 64);
        cert = cert.substring(64, cert.length);
      }
      else {
        result += "\n" + cert;
        cert = "";
      }
    }

    if (result[result.length] != "\n")
      result += "\n";
    result += endCert + "\n";
    return result;
  }

  /*
   * Extracts the tenant id from the give jwt token
   */
  public getTenantId(jwtString: string): string {
    var decodedToken = jsonwebtoken.decode(jwtString);

    if (decodedToken) {
      return decodedToken.tid;
    } else {
      return null;
    }
  };

  /*
   * This function loads the open-id configuration for a specific AAD tenant
   * from a well known application.
   */
  public async requestOpenIdConfig(tenantId) {
    // we need to load the tenant specific open id config
    var tenantOpenIdConfigUrl = 'https://login.windows.net/' + tenantId + '/.well-known/openid-configuration';

    try {
      const response = await fetch(tenantOpenIdConfigUrl);
      const result = await response.json();
      return result;
    }
    catch (error) {
      throw error;
    }
  };

  /*
   * Download the signing certificates which is the public portion of the
   * keys used to sign the JWT token.  Signature updated to include options for the kid.
   */
  public async requestSigningCertificates(jwtSigningKeysLocation: string, options: any): Promise<any[]> {
    try {
      const response = await fetch(jwtSigningKeysLocation);
      const result = await response.json();
      let certificates = [];

      //Use KID to locate the public key and store the certificate chain.
      if (options && options.kid) {
        result.keys.find(function (publicKey) {
          if (publicKey.kid === options.kid) {
            publicKey.x5c.forEach(function (certificate) {
              certificates.push(AADTokenManager.ConvertCertificateToBeOpenSSLCompatible(certificate));
            });
          }
        })
      } else {
        result.keys.forEach(function (publicKeys) {
          publicKeys.x5c.forEach(function (certificate) {
            certificates.push(AADTokenManager.ConvertCertificateToBeOpenSSLCompatible(certificate));
          })
        });
      }

      return certificates;
    }
    catch (error) {
      throw error;
    }
  };

  /*
   * This function tries to verify the token with every certificate until
   * all certificates was testes or the first one matches. After that the token is valid
   */
  public async verify(jwt, certificates, options): Promise<string> {

    // ensure we have options
    if (!options) options = {};

    // set the correct algorithm
    options.algorithms = ['RS256'];

    // set the issuer we expect
    options.issuer = 'https://sts.windows.net/' + this.getTenantId(jwt) + '/';

    var valid = false;
    var lastError = null;

    certificates.every(function (certificate) {
      // verify the token
      try {
        // verify the token
        jsonwebtoken.verify(jwt, certificate, options);

        // set the state
        valid = true;
        lastError = null;

        // abort the enumeration
        return false;
      } catch (error) {

        // set teh error state
        lastError = error;

        // check if we should try the next certificate
        if (error.message === 'invalid signature') {
          return true;
        } else {
          return false;
        }
      }
    });

    if (valid) {
      return jsonwebtoken.decode(jwt);
    } else {
      throw lastError;
    }
  }
}

export default AADTokenManager;