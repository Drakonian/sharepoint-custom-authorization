using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.Functions.Worker.Http;
using System.Net;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Security.Cryptography.X509Certificates;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Tokens;

namespace VLD.AcquireToken
{
    public class AcquireToken
    {
        private readonly ILogger _logger;
        public AcquireToken(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<AcquireToken>();
        }

        [Function("AcquireToken")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
        {
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();

            JObject jData = JsonConvert.DeserializeObject<JObject>(requestBody);

            if (jData == null)
            {
                return await CreateErrorResponse(req, "Failed to read JSON data.");
            }

            // Extract data from JSON
            string clientId = jData["clientId"]?.ToString();
            string certificatePassword = jData["certificatePassword"]?.ToString();
            string tenantId = jData["tenantId"]?.ToString();
            string base64Cert = jData["base64Cert"]?.ToString();
            string type = jData["type"]?.ToString();

            if (type == "temp")
            {
                var emptyResponse = req.CreateResponse(HttpStatusCode.OK);
                return emptyResponse;
            }

            string aud = $"https://login.microsoftonline.com/{tenantId}/v2.0/";

            byte[] certBytes = Convert.FromBase64String(base64Cert);
            X509Certificate2 certificate = new X509Certificate2(certBytes, certificatePassword);
            var claims = new Dictionary<string, object>();

            claims["aud"] = aud;
            claims["sub"] = clientId;
            claims["iss"] = clientId;
            claims["jti"] = Guid.NewGuid().ToString();

            var signingCredentials = new X509SigningCredentials(certificate);
            var securityTokenDescriptor = new SecurityTokenDescriptor();
            securityTokenDescriptor.Claims = claims;
            securityTokenDescriptor.SigningCredentials = signingCredentials;

            var tokenHandler = new JsonWebTokenHandler();
            var clientAssertion = tokenHandler.CreateToken(securityTokenDescriptor);

            var successResponse = req.CreateResponse(HttpStatusCode.OK);
            successResponse.Headers.Add("Content-Type", "text/plain; charset=utf-8");

            await successResponse.WriteStringAsync(clientAssertion);

            return successResponse;
        }
        private async Task<HttpResponseData> CreateErrorResponse(HttpRequestData req, string message)
        {
            var response = req.CreateResponse(HttpStatusCode.BadRequest);
            response.Headers.Add("Content-Type", "text/plain; charset=utf-8");
            await response.WriteStringAsync(message);
            return response;
        }
    }
}
