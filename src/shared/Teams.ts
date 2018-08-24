const MICROSOFT_TEAMS_JS = "https://statics.teams.microsoft.com/sdk/v1.0/js/MicrosoftTeams.min.js"

const Teams = {
  connected: false,
  api: null,
  ctx: null,
  initialize
}

function initialize() {
  return new Promise(resolve => {
    const script = document.createElement("script")
    script.addEventListener("load", () => {
      Teams.api = (window as any).microsoftTeams
      if (Teams.api) {
        Teams.api.initialize()
        Teams.api.getContext(ctx => {
          Teams.connected = true
          Teams.ctx = ctx
          listenForThemeChanges()
          resolve()
        })
      }
    })
    script.src = MICROSOFT_TEAMS_JS
    document.head.appendChild(script)
  })
}

function listenForThemeChanges() {
  Teams.api.registerOnThemeChangeHandler(themeChanged);
  themeChanged(Teams.ctx.theme)
}

function themeChanged(theme) {
  const body = document.getElementsByTagName("body")[0];

  switch (theme) {
    case "dark":
    case "contrast":
      body.className = "theme-" + theme
      break;
    case "default":
      body.className = ""
      break;
    default:
      body.className = ""
      break;
  }
}

export default Teams