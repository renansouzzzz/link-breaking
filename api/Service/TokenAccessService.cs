using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace CallContent.Service
{
    public class TokenAccessService
    {
        public async Task<AuthenticationResult> GetToken()
        {
            //Env.Load("../.env");
            string tenantId = "your tenant id";
            string appId = "your app id";

            var authContext = new AuthenticationContext("https://login.microsoftonline.com/" + tenantId);

            var credential = new ClientCredential(appId, "your secret id app registration <Azure AD>");

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
