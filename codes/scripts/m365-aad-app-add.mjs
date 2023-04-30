#!/usr/bin/env zx

//Check if Windows Terminal. If Windows Terminal, set the shell to bash.exe
if (!!process.env.WT_SESSION) {
    $.shell = `C:/Program Files/Git/usr/bin/bash.exe`;
}

async function createAADApp(appName, apiPermissions, exposeAPIURL) {
    console.log(`Creating Azure AD app ${appName}...`);
    let createdApp = null;
    try {
        // Create the Azure AD app
        createdApp = JSON.parse(await $`m365 aad app add --name ${appName} --apisDelegated ${apiPermissions}  --redirectUris 'https://login.microsoftonline.com/common/oauth2/nativeclient, https://${exposeAPIURL}' --platform spa --grantAdminConsent  --multitenant --withSecret --uri api://${exposeAPIURL}/_appId_ --scopeName access_as_user --scopeAdminConsentDescription 'Access the application as Logged in User' --scopeAdminConsentDisplayName 'Access as the logged in User' --scopeConsentBy admins --output json`);
    }
    catch (err) {
        console.error(`  ${chalk.red(err.stderr)}`);
    }
    return createdApp;
}
const appName = 'Teams App Camp App Registration-V1';
const apiPermissions = 'https://graph.microsoft.com/User.Read';
const exposeAPIURL = '2845-52-172-145-182.ngrok-free.app';

const appCreated = await createAADApp(appName, apiPermissions,exposeAPIURL);
console.log(appCreated);
if (appCreated === null) {
    console.error(`   ${chalk.red(`Failed to create the App with Name ${appName}`)}`);
}
else {
    console.log(`Created Azure AD app ${appCreated.appName} with id ${appCreated.appId}`);
}