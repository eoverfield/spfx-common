export class DomHelpers {

  /**
   * @param url the url path to the css stylesheet
   */
  public static includeCss(url: string): void {
    //requires a valid url, and checks to ensure the stylesheet has not yet been included
    if (url && url.length && !DomHelpers.isCssIncluded(url)) {
      let stylesheet = document.createElement("LINK");

      stylesheet.setAttribute('type', 'text/css');
      stylesheet.setAttribute('rel', 'stylesheet');
      stylesheet.setAttribute('href', url);

      //add the stylesheet to the head
      document.head.appendChild(stylesheet);
    }
  }

  /**
   * @param url the url path to the javascript file
   */
  public static includeScript(url: string): void {
    //requires a valid url, and checks to ensure the script has not yet been included
    if (url && url.length && !DomHelpers.isScriptIncluded(url)) {
      let script = document.createElement('SCRIPT');

      script.setAttribute('type', 'text/javascript');
      script.setAttribute('async', 'false');
      script.setAttribute('src', url);

      //add the script to the head
      document.body.appendChild(script);
    }
  }

  /**
   * @param url the url path to the css stylesheet to check if already included
   */
  public static isCssIncluded(url:string): boolean {
    //requires styleSheets in document and a valid url
    if (!document.styleSheets || !url || url.length < 1) {
      return false;
    }

    //get all stylesheets, search for the url included, return true if found
    let stylesheetReferences: StyleSheetList = document.styleSheets;
    for (let i = 0; i < stylesheetReferences.length; i++) {
      if (stylesheetReferences[i] && stylesheetReferences[i].href && stylesheetReferences[i].href.toLowerCase() == url.toLowerCase()) {
        return true;
      }
    }

    return false;
  }

  /**
   * @param url the url path to the script to check if already included
   */
  public static isScriptIncluded(url:string): boolean {
    let scriptReferences: HTMLCollectionOf<HTMLScriptElement> = document.getElementsByTagName('script') as HTMLCollectionOf<HTMLScriptElement>;

    //requires scripts in document and a valid url
    if (!scriptReferences || !url || url.length < 1) {
      return false;
    }

    //search doc scripts for the url included, return true if found
    for (let i = scriptReferences.length; i >= 0; i--) {
      if (scriptReferences[i] && scriptReferences[i].src && scriptReferences[i].src.toLowerCase() == url.toLowerCase()) {
        return true;
      }
    }

    return false;
  }
}
