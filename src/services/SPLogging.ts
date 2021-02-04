import { MiscFunctions } from '.';

export abstract class SPLogging {
  public static DebugCode(bDebugMode: boolean): void {
    if (bDebugMode) {
      // tslint:disable-next-line: no-debugger
      debugger;
    }
  }

  public static LogConsole(
    bDebugMode: boolean,
    location: string,
    message: string
  ): void {
    try {
      if (bDebugMode) {
        console.log('Location: ' + location, message);
      }
    } catch (e) {
      console.error(location + ' - ' + message + ' ' + e.message);
    }
  }

  public static LogConsoleData(
    bDebugMode: boolean,
    location: string,
    message: string,
    // tslint:disable-next-line
    objecttoshow: any
  ): void {
    try {
      if (bDebugMode) {
        if (objecttoshow !== undefined) {
          console.log('Location: ' + location, message);
          if (!MiscFunctions.IsInternetExplorer()) {
            console.table(objecttoshow);
          } else {
            console.log(JSON.stringify(objecttoshow));
          }
        } else {
          console.log('Location: ' + location, message);
        }
      }
    } catch (e) {
      console.error(location + ' - ' + message + ' ' + e.message);
    }
  }
  public static LogErrorData(
    location: string,
    message: string,
    // tslint:disable-next-line
    objecttoshow: any
  ): void {
    try {
      console.error('Location: ' + location, message);
      console.table(objecttoshow);
    } catch (e) {
      console.error(location + ' - ' + message + ' ' + e.message);
    }
  }

  public static LogError(location: string, message: string): void {
    try {
      if (window.location.href.indexOf('workbench.aspx') > -1) {
        alert('Location: ' + location + ' -- ' + message);
      }
      console.error('Location: ' + location, message);
    } catch (e) {
      console.error(location + ' - ' + message + ' ' + e.message);
    }
  }
}
