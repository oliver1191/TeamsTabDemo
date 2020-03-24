import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/perimeterTabDemoTab/index.html")
@PreventIframe("/perimeterTabDemoTab/config.html")
@PreventIframe("/perimeterTabDemoTab/remove.html")
export class PerimeterTabDemoTab {
}
