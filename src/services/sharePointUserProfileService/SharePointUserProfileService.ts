import { WebPartContext } from                  '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from    '@microsoft/sp-application-base';
import { sp } from                              '@pnp/sp';
import { CurrentUser } from                     '@pnp/sp/src/siteusers';

import {
  ILocalStorageService,
  ILocalStorageKey,
  LocalStorageService
} from                                          '../localStorageService';
import {
  ISharePointUserProfileServiceProperties,
} from                                          './';
import { DebugLogging } from                    '../../helpers/debugLogging';


export class SharePointUserProfileService {

    private useLocalStorage: boolean = true; //default to always use local storage
    private localStorageKeyName: string = "Profile";
    private localStorageKeyPrefix: string = "UPS";
    private localStorageTimeout: number = 30; //default to 30 minutes

    private logger: DebugLogging; //used to set up logging

    public constructor(properties: ISharePointUserProfileServiceProperties = {}) {
      this.logger = new DebugLogging();
      this.logger.enableLog = false; //default to not enable log

      if (properties) {
        if (typeof properties.useLocalStorage != "undefined") {
          this.useLocalStorage = properties.useLocalStorage;
        }
        if (typeof properties.localStorageKeyName != "undefined") {
          this.localStorageKeyName = properties.localStorageKeyName;
        }
        if (typeof properties.localStorageKeyPrefix != "undefined") {
          this.localStorageKeyPrefix = properties.localStorageKeyPrefix;
        }
        if (typeof properties.localStorageTimeout != "undefined") {
          this.localStorageTimeout = properties.localStorageTimeout;
        }

        if (typeof properties.enableLog != "undefined") {
          this.logger.enableLog = properties.enableLog;
        }
      }
    }

    /*
    Allow for initialization of service
    provide the context of the webpart or application customizer, allowing to send context to this instance of @pnp/sp for sp.setup
    */
    public static async init(context: ApplicationCustomizerContext | WebPartContext) : Promise<void> {
      sp.setup({
        spfxContext: context
      });
    }

    /**
     * Attempt to get the user profile, or only one particular profile property if key provided
     * @param key the key to attempt to retrieve from user profile
     * @return any - the found and validated local storage value
     */
    public async get(key?: string): Promise<any> {

      var p = new Promise<string>(async (resolve, reject) => {
        var myProperties: any = null;

        let localStorageService: ILocalStorageService = new LocalStorageService();

        //set up local storage object to attempt to get data
        let localStorageKeyValue: ILocalStorageKey = {
          keyName: this.localStorageKeyName,
          keyPrefix: this.localStorageKeyPrefix,
          timeOutInMinutes: this.localStorageTimeout
        } as ILocalStorageKey;

        this.logger.log("SharePointUserProfileService.get: initialized");

        //check local storage if timeout great than 0
        if (this.useLocalStorage && this.localStorageTimeout > 0) {
          //attempt to get valid response from local storage
          try {

            this.logger.log("SharePointUserProfileService.get: local storage requested for " + localStorageKeyValue.keyName);
            myProperties = await localStorageService.get(localStorageKeyValue);

            //do not return yet as may need to get a specific key value
          }
          catch (err) {
            //ensure that myProperties is set back to null
            myProperties = null;

            this.logger.log("SharePointUserProfileService.get: local storage not found for " + localStorageKeyValue.keyName);
            this.logger.log(err);
          }

        }

        //if we do not have myProperties, then either local storage not used, not available, or expired.
        //get my properties
        if (!myProperties) {

          this.logger.log("SharePointUserProfileService.get: user profile not available in local storage, retrieve from REST");

          try {
            //go and get all profile props
            myProperties = await sp.profiles.myProperties.get();

            if (myProperties) {
              this.logger.log("SharePointUserProfileService.get: user profile retrieved");
              this.logger.log(myProperties);

              if (this.useLocalStorage) {
                //set up local storage object to attempt to get data
                localStorageKeyValue.keyValue = myProperties;

                //store to local storage
                await localStorageService.set(localStorageKeyValue);

                this.logger.log("SharePointUserProfileService.get: user profile stored in local storage " + localStorageKeyValue.keyName);
              }
            }
          }
          catch (err) {
            //ensure that myProperties is set back to null
            myProperties = null;

            this.logger.log("SharePointUserProfileService.get: an error occurred retrieving user profile");
            this.logger.log(err);
          }

        }

        if (myProperties) {

          //if no key was provided, then we can return all properties
          if (!key || key.length < 1) {
            this.logger.log("SharePointUserProfileService.get: returning entire user profile");
            this.logger.log(myProperties);

            resolve(myProperties);
            return;
          }
          else {

            this.logger.log("SharePointUserProfileService.get: returning user profile based on key " + key);

            // otherwise we want to return just the one property if found
            if (myProperties.UserProfileProperties) {

              let propertyFound: boolean = false;

              //go and find the requested property
              for(var i=0;i<myProperties.UserProfileProperties.length;i++) {

                //the current property
                let thisProp = myProperties.UserProfileProperties[i];

                //check to see if the current property has the same key as we are looking for
                if (thisProp && thisProp["Key"] && thisProp["Key"].toLowerCase() == key.toLowerCase()) {

                  propertyFound = true;

                  this.logger.log("SharePointUserProfileService.get: user profile key " + key + " found, will return value");
                  this.logger.log(thisProp["Value"]);

                  //found, return and we are done
                  resolve(thisProp["Value"]);

                  break;
                }
              }

              if (!propertyFound) {
                this.logger.log("SharePointUserProfileService.get: error - user profile key " + key + " not found");
                reject("notFound");
              }
            }
            else {
              this.logger.log("SharePointUserProfileService.get: error - no UserProfileProperties available");
              reject("noPropertiesAvailable");
            }
          }
        }
        else {
          //myProperties not available, thus we need to reject

          this.logger.log("SharePointUserProfileService.get: error - no properties available");
          reject("noProperties");
        }

      });

      return p;
    }



    /**
     * Attempt to set a user profile property
     * @param key the user profile property key to update
     * @param value the value of the key to store
     * @return void - rejects if unsuccessful
     */
    public async set(key: string, value: string): Promise<void> {

      var p = new Promise<void>(async (resolve, reject) => {

        let localStorageService: ILocalStorageService = new LocalStorageService();

        //set up local storage object to attempt to clear
        let localStorageKeyValue: ILocalStorageKey = {
          keyName: this.localStorageKeyName,
          keyPrefix: this.localStorageKeyPrefix
        } as ILocalStorageKey;

        //key is required
        if (!key) {
          this.logger.log("SharePointUserProfileService.set: error, key required");

          reject("keyRequired");
        }
        //value should at least be an empty string
        if (!value) {
          value = "";
        }


        try {
          //get the current user to get their account name
          let currentUser: CurrentUser;

          try {
            this.logger.log("SharePointUserProfileService.set: attempt to get current user");

            currentUser = await sp.web.currentUser.get();
          }
          catch (err) {
            this.logger.log("SharePointUserProfileService.set: error, unable to retrieve current user");
            this.logger.log(err);

            reject(err);
            return;
          }

          //verify we do in fact have a valid user
          if (!currentUser) {
            this.logger.log("SharePointUserProfileService.set: error, current user not available");

            reject("noCurrentUser");
            return;
          }

          //current user must be available, update the user profile property
          try {
            this.logger.log("SharePointUserProfileService.set: setting the user, " + currentUser["LoginName"] + ", property key: " + key + " with value:");
            this.logger.log(value);

            await sp.profiles.setSingleValueProfileProperty(currentUser["LoginName"], key, value);
          }
          catch (err) {
            this.logger.log("SharePointUserProfileService.set: error, unable to set user profile property " + key);
            this.logger.log(err);

            reject(err);
            return;
          }

          //if local storage used, clear currently stored properties
          if (this.useLocalStorage) {
            this.logger.log("SharePointUserProfileService.set: local storage utilized, reset value");

            //reset the value for this profile in local storage so if required, it will be retrieved again
            localStorageKeyValue.keyValue = "";

            //store to local storage
            await localStorageService.set(localStorageKeyValue);

            this.logger.log("SharePointUserProfileService.set: local storage cleared");
          }

          //if we are here, successfully set user profile property and cleared local storage is required
          resolve();
        }
        catch (err) {
          this.logger.log("SharePointUserProfileService.set: error, general error setting profile property with key " + key);
          this.logger.log(err);

          reject(err);
        }
      });

      return p;
    }
}
