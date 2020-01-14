## Development server - Prerequisites

- [Node.js](https://nodejs.org) installed on your development machine. If you do not have Node.js, visit the previous link for download options.
- For testing purpose, using your current DXC account credentials.

## Register a web application with the Azure Active Directory admin center

1. Open a browser and navigate to the [Azure Active Directory admin center](https://aad.portal.azure.com). Login using your DXC.

1. Select **Azure Active Directory** in the left-hand navigation, then select **App registrations** under **Manage**.

    ![A screenshot of the App registrations ](/images/app-registration.png)

1. Select **New registration**. On the **Register an application** page, set the values as follows.

    - Set **Name** to `Node.js Graph Tutorial`.
    - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
    - Under **Redirect URI**, set the first drop-down to `Web` and set the value to `http://localhost:3000/auth/callback`.

    ![A screenshot of the Register an application page](/images/aad-register-an-app.png)

1. Choose **Register**. On the **Node.js Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.

    ![A screenshot of the application ID of the new app registration](/images/aad-application-id.png)

1. Select **Authentication** under **Manage**. Locate the **Implicit grant** section and enable **ID tokens**. Choose **Save**.

    ![A screenshot of the Implicit grant section](/images/aad-implicit-grant.png)

1. Select **API permissions** under **Manage**. Choose the **+ Add a permission** button and in the **Commonly used Microsoft APIs**, choose **Microsoft Graph**.
    - For **Delegated permissions**, select following permissions: User.ReadBasic.All, User.Read.All, User.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All, Directory.AccessAsUser.All
    - For **Application permissions**, select following permissions: User.Read.All, User.ReadWrite.All, Directory.Read.All, Directory.ReadWrite.All

    ![A screenshot of the Implicit grant section](/images/request_permission.png)

    <a name="step6Note">
    <b>Note</b>: We need an admin user in DXC tenant to grant admin consent for those permissions that are admin consent required. At the moment we are still waiting for this to be approved, for that reason some of Graph API endpoints which need those permissions can not be used (e.g 'GET /users'; 'GET /users/{id | userPrincipalName}'.
    </a>
1. Select **Certificates & secrets** under **Manage**. Select the **New client secret** button. Enter a value in **Description** and select one of the options for **Expires** and choose **Add**.

    ![A screenshot of the Add a client secret dialog](/images/aad-new-client-secret.png)

1. Copy the client secret value before you leave this page. You will need it in the next step.

    > [!IMPORTANT]
    > This client secret is never shown again, so make sure you copy it now.

    ![A screenshot of the newly added client secret](/images/aad-copy-client-secret.png)

## Configure the sample

1. Edit the `.env` file and make the following changes.
    1. Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.
    1. Replace `YOUR_APP_PASSWORD_HERE` with the password you got from the App Registration Portal.
1. In your command-line interface (CLI), navigate to this directory and run the following command to install requirements.

    ```Shell
    npm install
    ```

## Run the sample

1. Run the following command in your CLI to start the application.

    ```Shell
    npm start
    ```

1. Open new browser, open `http://localhost:3000/auth/signin`, it will automatically redirect to current DXC login page. Use your DXC account with global password to sign in.
1. Make sure keeping your current browser from previous step opened, type an enpoint in new browser tab. For testing purpose, the following endpoints have been developed in the project:
    1. `http://localhost:3000/users/getAllUsers`: Get all users without setting 'top' parameter in the endpoint, by defautl the first 100 records will return.(*)
    1. `http://localhost:3000/users/getAllUsersV2`: Get all users with '$top=999' parameter in the endpoint. This method also supports to append all users in the returned array for the case the total number of users is greater than 999.(*)
    1. `http://localhost:3000/users/getFilteredUsers/abc`: Search for users with the specifiec name (e.g. 'abc') across multiple properties.(*)
    1. `http://localhost:3000/users/getCurrentUserDetail`: Get detail info for current logging user.
    1. `http://localhost:3000/calendar`: Get list of meeting schedules of current logging user.
    1. `http://localhost:3000/auth/signout`: Log out for current user.

    **Note**: As mentioned in the [step 6](#step6Note) above, the endpoints which are marked as (*) need admin consent permissions for usage.

## Some notes
- Preferable approach for MS Graph API authentication flow: using confidential client app.
- Service to service calls using client credentials, flow:
    - The client application authenticates to the Azure AD token issuance endpoint and requests an access token.
    - The Azure AD token issuance endpoint issues the access token.
    - The access token is used to authenticate to the secured resource.
    - Data from the secured resource is returned to the client application.
- Adding a delegated permission to an application does not automatically grant consent to the users within the tenant. Users must still manually consent for the added delegated permissions at runtime, unless the administrator grants consent on behalf of all users.

## Reference
- Microsoft Graph API endpoint, get all users: https://docs.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0&tabs=http
- express-session to store values in an in-memory server-side session: https://github.com/expressjs/session
- passport-azure-ad for authenticating and getting access tokens: https://github.com/AzureAD/passport-azure-ad
- simple-oauth2 for token management: https://github.com/lelylan/simple-oauth2
- microsoft-graph-client for making calls to Microsoft Graph: https://github.com/microsoftgraph/msgraph-sdk-javascript
- isomorphic-fetch to polyfill the fetch for Node. A fetch polyfill is required for the microsoft-graph-client library: https://github.com/matthew-andrews/isomorphic-fetch
- MS Graph training with NodeJs app: https://github.com/microsoftgraph/msgraph-training-nodeexpressapp