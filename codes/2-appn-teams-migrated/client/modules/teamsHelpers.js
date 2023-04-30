import 'https://res.cdn.office.net/teams-js/2.0.0/js/MicrosoftTeams.min.js';
// Ensure that the Teams SDK is initialized once no matter how often this is called
let teamsInitPromise;
export function ensureTeamsSdkInitialized() {
    if (!teamsInitPromise) {
        teamsInitPromise = microsoftTeams.app.initialize();
    }
    return teamsInitPromise;
}
// Function returns a promise which resolves to true if we're running in Teams
export async function inTeams() {
    try {
        await ensureTeamsSdkInitialized();
        const context = await microsoftTeams.app.getContext();
        return (context.app.host.name === microsoftTeams.HostName.teams);
    }
    catch (e) {
        console.log(`${e} from Teams SDK, may be running outside of Teams`);
        return false;
    }
}

function setTheme(theme) {
    const el = document.documentElement;
    el.setAttribute('data-theme', theme);
}

// Inline code to set theme on any page using teamsHelpers
(async () => {
    await ensureTeamsSdkInitialized();
    const context = await microsoftTeams.app.getContext();
    setTheme(context.app.theme);

    // When the theme changes, update the CSS again
    microsoftTeams.registerOnThemeChangeHandler((theme) => {
        setTheme(theme);
    });    
})();
