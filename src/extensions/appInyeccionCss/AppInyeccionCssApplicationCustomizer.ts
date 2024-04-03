import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

import * as strings from "AppInyeccionCssApplicationCustomizerStrings";

const LOG_SOURCE: string = "AppInyeccionCssApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAppInyeccionCssApplicationCustomizerProperties {
  // This is an example; replace with your own property
  CssFileLocation: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AppInyeccionCssApplicationCustomizer extends BaseApplicationCustomizer<IAppInyeccionCssApplicationCustomizerProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let fileURL: string =
      this.context.pageContext.site.serverRelativeUrl +
      this.properties.CssFileLocation;

    if (fileURL) {
      const head: any =
        document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = fileURL;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }
    return Promise.resolve();
  }
}
