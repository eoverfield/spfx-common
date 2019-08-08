import { Md5 } from                             'ts-md5/dist/md5';

import {
  ILocalStorageService,
  ILocalStorageKey,
  ILocalStorageObject,
  ILocalStorageServiceProperties
} from                                          './ILocalStorageService';
import { DebugLogging } from                    '../../helpers/debugLogging';

export class LocalStorageService implements ILocalStorageService {
    private logger: DebugLogging; //used to set up logging

    public constructor(properties: ILocalStorageServiceProperties = {}) {
      this.logger = new DebugLogging();
      this.logger.enableLog = false; //default to not enable log

      if (properties) {
        if (typeof properties.enableLog != "undefined") {
          this.logger.enableLog = properties.enableLog;
        }
      }
    }

    /**
     * Attempt to get local storage value based on key
     * @param keyToken the key value used to retrive and verify local storage
     * @return any - the found and validated local storage value
     */
    public async get(keyToken: ILocalStorageKey): Promise<any> {

      var p = new Promise<any>(async (resolve, reject) => {
        try {

          var returnValue: any;

          this.logger.log("LocalStorageService.get: initializing");

          //get the hash of the local storage token based on value
          var keyHash: string | Int32Array = Md5.hashStr(JSON.stringify(keyToken.keyName));
          this.logger.log("LocalStorageService.get: keyHash - " + keyHash);

          //create the corrrect storage key based on keyHash and possible prefix
          const storageKey: string = (keyToken.keyPrefix ? keyToken.keyPrefix + "_" : "") + keyHash;
          this.logger.log("LocalStorageService.get: storageKey - " + storageKey);

          //attempt to get the key/value from local storage based on storageKey
          const keyValue: ILocalStorageObject = JSON.parse(localStorage.getItem(storageKey)) as ILocalStorageObject;

          //with a valid response, we can continue
          if (keyValue) {
            this.logger.log("LocalStorageService.get: key was found in local storage with value:");
            this.logger.log(keyValue);

            //check timeout if one provided
            if (keyToken.timeOutInMinutes > 0) {

              this.logger.log("LocalStorageService.get: timeOut provided, verify not stale " + keyToken.timeOutInMinutes);

              //have to get proper date object
              const keyDate: Date = new Date(keyValue.keyDate.toString());

              //determine the local time at which this key/value should expire
              const timeout: Date = new Date(keyDate.getTime() + keyToken.timeOutInMinutes*60000);

              this.logger.log("LocalStorageService.get: now " + new Date(Date.now()).toString());
              this.logger.log("LocalStorageService.get: timeout " + timeout.toString());

              //check to see if the local storage is stale or not
              if (timeout.getTime() > Date.now()) {

                //still valid, thus return whatever was found in local storage
                returnValue = keyValue.keyValue;

                this.logger.log("LocalStorageService.get: storage valid, return value");
              }
              else {

                //attempt to remove from local storage for garbage collection
                localStorage.removeItem(storageKey);

                this.logger.log("LocalStorageService.get: storage is stale, should have been removed");
              }
            }
            else {

              //no timeout was provided, thus simply return
              returnValue = keyValue.keyValue;

              this.logger.log("LocalStorageService.get: storage valid without timeout, return value");
            }
          }
          else {
            //key was not found in local storage, simply continue
            this.logger.log("LocalStorageService.get: key not found in local storage, will retun nothing " + storageKey);
          }

          //resolve the promise with whatever was found, a valid, or null
          resolve(returnValue);

        }
        catch (err) {
          this.logger.log("LocalStorageService.get: error - fatal error occurred");
          this.logger.log(err);

          reject(null);
        }

        return;
      });

      return p;
    }

    /**
     * Attempt to set local storage value based on key
     * @param keyToken the key value used to store to local storage
     * @return boolean - true upon success
     */
    public async set(keyToken: ILocalStorageKey): Promise<boolean> {

      var p = new Promise<any>(async (resolve, reject) => {
        try {

          this.logger.log("LocalStorageService.set: initializing");

          //get the hash of the local storage token based on value
          var keyHash: string | Int32Array = Md5.hashStr(JSON.stringify(keyToken.keyName));
          this.logger.log("LocalStorageService.set: keyHash - " + keyHash);

          //create the corrrect storage key based on keyHash and possible prefix
          const storageKey: string = (keyToken.keyPrefix ? keyToken.keyPrefix + "_" : "") + keyHash;
          this.logger.log("LocalStorageService.set: storageKey - " + storageKey);


          //create a storage object to hold the value and storage date/time "now"
          const keyValue: ILocalStorageObject = {
            keyValue: keyToken.keyValue,
            keyDate: new Date(Date.now())
          } as ILocalStorageObject;


          this.logger.log("LocalStorageService.set: setting local storage key/value pair");

          //attempt to store to local storage
          localStorage.setItem(storageKey, JSON.stringify(keyValue));

          this.logger.log("LocalStorageService.set: local storage key/value set");

          resolve(true);
        }
        catch (err) {
          this.logger.log("LocalStorageService.set: error - fatal error occurred");
          this.logger.log(err);

          reject(false);
        }

        return;
      });

      return p;
    }
}
