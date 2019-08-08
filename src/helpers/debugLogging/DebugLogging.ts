export class DebugLogging {
  //public setting to allow an instance to enable or disable console logging, default to false
  public enableLog: boolean = false;

  /**
   * @param logMessage - the message / object to log based on the logging flag, will simply send whatever provided to console.log
   */
  public log(logMessage: any): void {
    if (this.enableLog) {
      console.log(logMessage);
    }
  }
}
