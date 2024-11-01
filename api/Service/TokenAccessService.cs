using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace CallContent.Service
{
    public class TokenAccessService
    {
        private readonly string tenantId = "";
        private readonly string appId = "";

        public async Task<AuthenticationResult> GetTokenGraph()
        {
            //Env.Load("../.env");

            var authContext = new AuthenticationContext("https://login.microsoftonline.com/" + tenantId);

            var credential = new ClientCredential(appId, "<app secret>");

            var GraphAAD_URL = string.Format("https://graph.microsoft.com/");

            try
            {
                AuthenticationResult result = await authContext.AcquireTokenAsync(GraphAAD_URL, credential);

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception("Error Acquiring Access Token: \n" + ex.Message);
            }
        }
    }
}
