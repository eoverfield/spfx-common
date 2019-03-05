import { sp } from '@pnp/sp';

import {
  ILocalStorageService,
  ILocalStorageKey,
  LocalStorageService
} from                                                  '../localStorageService';

export class SharePointUserProfileService {

    public static useLocalStorage: boolean = true; //default to always use local storage
    public static localStorageKeyName: string = "Profile";
    public static localStorageKeyPrefix: string = "UPS";
    public static localStorageTimeout: number = 30; //default to 30 minutes

    public constructor() {
    }

    /**
     * Attempt to get the user profile, or only one particular profile property if key provided
     * @param key the key to attempt to retrieve from user profile
     * @return any - the found and validated local storage value
     */
    public async get(key?: string): Promise<any> {

      var p = new Promise<string>(async (resolve, reject) => {
        let localStorageService: ILocalStorageService = new LocalStorageService();

        var myProperties: any = null;

        //check local storage if timeout great than 0
        if (SharePointUserProfileService.useLocalStorage && SharePointUserProfileService.localStorageTimeout > 0) {

          //set up local storage object to attempt to get data
          let localStorageKeyValue: ILocalStorageKey = {
            keyName: SharePointUserProfileService.localStorageKeyName,
            keyPrefix: SharePointUserProfileService.localStorageKeyPrefix,
            timeOutInMinutes: SharePointUserProfileService.localStorageTimeout
          } as ILocalStorageKey;

          //attempt to get valid response from local storage
          try {

            myProperties = await localStorageService.get(localStorageKeyValue);

          }
          catch (err) {
            //ensure that myProperties is set back to null
            myProperties = null;
          }

        }

        //if we do not have myProperties, then either local storage not used, not available, or expired.
        //get my properties
        if (!myProperties) {

          try {
            //go and get all profile props
            myProperties = await sp.profiles.myProperties.get();

            if (myProperties && SharePointUserProfileService.useLocalStorage) {

              //set up local storage object to attempt to get data
              let localStorageKeyValue: ILocalStorageKey = {
                keyName: SharePointUserProfileService.localStorageKeyName,
                keyPrefix: SharePointUserProfileService.localStorageKeyPrefix,
                keyValue: myProperties
              } as ILocalStorageKey;

              //store to local storage
              await localStorageService.set(localStorageKeyValue);

            }
          }
          catch (err) {
            //ensure that myProperties is set back to null
            myProperties = null;
          }

        }

        if (myProperties) {

          //if no key was provided, then we can return all properties
          if (!key || key.length < 1) {
            resolve(myProperties);
          }
          else {

            // otherwise we want to return just the one property if found
            if (myProperties.UserProfileProperties) {

              //go and find the requested property
              for(var i=0;i<myProperties.UserProfileProperties.length;i++) {

                //the current property
                let thisProp = myProperties.UserProfileProperties[i];

                //check to see if the current property has the same key as we are looking for
                if (thisProp && thisProp["Key"] && thisProp["Key"].toLowerCase() == key.toLowerCase()) {

                  //found, return and we are done
                  resolve(thisProp["Value"]);

                  break;
                }
              }
            }
            else {

              console.log(`property ${key} not found in profile`);
              reject();

            }
          }
        }
        else {
          //myProperties not available, thus we need to reject
          reject();
        }

      });

      return p;
    }
}
