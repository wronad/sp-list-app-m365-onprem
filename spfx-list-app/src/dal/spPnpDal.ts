import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp, pnp logging system, and any other selective imports needed
import { LogLevel, PnPLogging } from "@pnp/logging";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/presets/all";

export const getSpFrameworkIF = (
  context: WebPartContext,
  url: string
): SPFI => {
  return spfi(url).using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
};
