import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPLogging } from ".";

export class MiscFunctions {
  public static GetSearchUrl(
    debugMode: boolean,
    queryText: string,
    properties: string,
    rowLimit: number
  ): string {
    const retVal: string =
      MiscFunctions.GetWebAppUrl() +
      "search/_api/search/query?rowlimit=" +
      rowLimit +
      "&querytext=%27" +
      queryText +
      "%27&selectproperties=%27" +
      properties +
      "%27";
    SPLogging.LogConsole(debugMode, "GetSearcUrl", retVal);
    return retVal;
  }

  public static ReplaceIfPresent(
    cMainVal: string,
    cFindVal: string,
    cReplaceVal: string
  ): string {
    if (cMainVal.indexOf(cFindVal) > -1) {
      cMainVal = cMainVal.replace(cFindVal, cReplaceVal);
    }
    return cMainVal;
  }

  public static getIDFromPath(strPath: string): string {
    const retVal: string = strPath.substring(strPath.indexOf("ID=") + 3, 100);
    return retVal;
  }

  public static GetListNameFromPath(strPath: string): string {
    let retVal: string = strPath;
    const iStart: number = retVal.toLowerCase().indexOf("/lists/") + 7;
    retVal = retVal.substring(iStart, 100);
    retVal = retVal.substr(0, retVal.indexOf("/"));
    return retVal;
  }

  public static IsEmpty(strVal: string): boolean {
    switch (strVal) {
      case "":
      // tslint:disable-next-line
      case null:
      case undefined:
        return true;
    }
    return false;
  }

  public static getRandomColor(): string {
    const letters: string = "0123456789ABCDEF";
    let color: string = "#";
    for (let i: number = 0; i < 6; i++) {
      color += letters[Math.floor(Math.random() * 16)];
    }
    return color;
  }

  public static IsInternetExplorer(): boolean {
    try {
      const ua: string = window.navigator.userAgent;
      const msie: number = ua.indexOf("MSIE ");

      if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) {
        // If Internet Explorer, return version number
        SPLogging.LogConsole(
          true,
          "EventService.IsinternetExplorer",
          "IE Version: " + ua.substring(msie + 5, ua.indexOf(".", msie))
        );
        return true;
      }
    } catch (e) {
      SPLogging.LogError("IsInternetExplorer", e.message);
    }
    return false;
  }

  public static GetWebAppUrl(): string {
    const workUrl: string = window.location.href;
    const iStart: number = workUrl.indexOf("//") + 2;
    let retVal: string = workUrl.substring(0, workUrl.indexOf("/", iStart) + 1);
    retVal = retVal.toLowerCase();
    if (retVal.indexOf("dominionenergyo365") > -1) {
      retVal = retVal.replace(".com", ".com.us3.cas.ms");
    }
    return retVal;
  }

  public static GetFileNameFromUrl(filePath: string): string {
    return filePath.substring(filePath.lastIndexOf("/") + 1);
  }

  public static GetServerRelativeUrl(itemUrl: string): string {
    const iPos: number = itemUrl.indexOf(".org") + 4;
    return itemUrl.substring(iPos);
  }

  public static ReplaceText(
    inputText: string,
    searchStr: string,
    newText: string
  ): string {
    if (inputText.toLowerCase().indexOf(searchStr) > -1) {
      inputText = inputText.replace(searchStr, newText);
    }
    return inputText;
  }

  public static TranslateTokens(
    inputText: string,
    ctx: WebPartContext
  ): string {
    const email: string = ctx.pageContext.user.email;

    inputText = MiscFunctions.ReplaceText(inputText, "{{Email}}", email);
    inputText = MiscFunctions.ReplaceText(inputText, "{{EMail}}", email);
    inputText = MiscFunctions.ReplaceText(inputText, "{{email}}", email);

    const username: string = ctx.pageContext.user.displayName;
    inputText = MiscFunctions.ReplaceText(inputText, "{{username}}", username);
    inputText = MiscFunctions.ReplaceText(inputText, "{{userName}}", username);
    inputText = MiscFunctions.ReplaceText(inputText, "{{UserName}}", username);

    if (inputText.toLowerCase().indexOf("{{currentsite}}") > -1) {
      inputText = inputText.replace(
        "{{currentsite}}",
        "Path:" + ctx.pageContext.web.absoluteUrl + "*"
      );
    }

    if (inputText.toLowerCase().indexOf("{{currentwebapp}}") > -1) {
      inputText = inputText.replace(
        "{{currentwebapp}}",
        "Path:" + MiscFunctions.GetWebAppUrl() + "*"
      );
    }

    if (inputText.toLowerCase().indexOf("{{checkedout}}") > -1) {
      inputText = inputText.replace(
        "{{checkedout}}",
        // tslint:disable-next-line
        "(cou:a* OR cou:b* OR cou:c* OR cou:d* OR cou:e* OR cou:f* OR cou:g* OR cou:h* OR cou:i* OR cou:j* OR cou:k* OR cou:l* OR cou:m* OR cou:n* OR cou:o* OR cou:p* OR cou:q* OR cou:r* OR cou:s* OR cou:t* OR cou:u* OR cou:v* OR cou:w* OR cou:x* OR cou:y* OR cou:z*)"
      );
    }
    return inputText;
  }

  public static getUrlParameter(name: string): string {
    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
    const regex: RegExp = new RegExp("[\\?&]" + name + "=([^&#]*)");
    // tslint:disable-next-line
    const results = regex.exec(location.search);
    return results === null
      ? ""
      : decodeURIComponent(results[1].replace(/\+/g, " "));
  }

  public static ValidValue(value: string): boolean {
    // tslint:disable-next-line
    if (value == null || value === undefined || value === "") {
      return false;
    }

    return true;
  }
}
