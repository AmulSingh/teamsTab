import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/onemoretabTab/index.html")
@PreventIframe("/onemoretabTab/config.html")
@PreventIframe("/onemoretabTab/remove.html")
export class OnemoretabTab {
}
