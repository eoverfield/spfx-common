/*
based on: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
*/

import { WebPartContext } from                  '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from    '@microsoft/sp-application-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from                                          '@microsoft/sp-http';

import {
  ITenantProperty,
  ITenantPropertiesServiceProperties
} from                                          './';
import {
  ILocalStorageService,
  ILocalStorageKey,
  LocalStorageService
} from                                          '../localStorageService';
import { DebugLogging } from                    '../../helpers/debugLogging';

export class TenantPropertiesService {

    //set globally by init method
    private static _sharePointClient: SPHttpClient;
    private static _webUrl: string;

    private useLocalStorage: boolean = true; //default to always use local storage
    private localStorageKeyPrefix: string = "TP";
    private localStorageTimeout: number = 30; //default to 30 minutes

    private logger: DebugLogging; //used to set up logging

    public constructor(properties: ITenantPropertiesServiceProperties = {}) {
      this.logger = new DebugLogging();
      this.logger.enableLog = false; //default to not enable log

      if (properties) {
        if (typeof properties.useLocalStorage != "undefined") {
          this.useLocalStorage = properties.useLocalStorage;
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

    public static async init(context: WebPartContext | ApplicationCustomizerContext | any): Promise<void> {
      TenantPropertiesService._sharePointClient = await context.spHttpClient;
      TenantPropertiesService._webUrl = context.pageContext.web.absoluteUrl;
    }

    /**
     * Attempt to get the user profile, or only one particular profile property if key provided
     * @param key the key to attempt to retrieve from user profile
     * @return any - the found and validated local storage value
     */
    public async get(key: string): Promise<ITenantProperty> {

      var p = new Promise<ITenantProperty>(async (resolve, reject) => {
        var tenantProperty: ITenantProperty = null;

        let localStorageService: ILocalStorageService = new LocalStorageService();

        //set up local storage object to attempt to get data
        let localStorageKeyValue: ILocalStorageKey = {
          keyName: key,
          keyPrefix: this.localStorageKeyPrefix,
          timeOutInMinutes: this.localStorageTimeout
        } as ILocalStorageKey;

        this.logger.log("TenantPropertyService.get: initialized");

        //validate that context was provided
        if (!TenantPropertiesService._sharePointClient || !TenantPropertiesService._webUrl) {
          this.logger.log("TenantPropertyService.get: error, context required");
          reject("contextRequired");
          return;
        }
        if (!key || key.length < 1) {
          this.logger.log("TenantPropertyService.get: error, key required");
          reject("keyRequired");
          return;
        }

        //check local storage if timeout great than 0
        if (this.useLocalStorage && this.localStorageTimeout > 0) {
          //attempt to get valid response from local storage
          try {

            this.logger.log("TenantPropertyService.get: local storage requested for " + localStorageKeyValue.keyName);
            tenantProperty = await localStorageService.get(localStorageKeyValue);

            if (tenantProperty) {
              this.logger.log("TenantPropertyService.get: local storage found for " + localStorageKeyValue.keyName);
              this.logger.log(tenantProperty);
              resolve(tenantProperty);
              return;
            }
          }
          catch (err) {
            //ensure that myProperties is set back to null
            tenantProperty = null;

            this.logger.log("TenantPropertyService.get: local storage not found for " + localStorageKeyValue.keyName);
            this.logger.log(err);
          }

        }

        //if we do not have the tenant property, then either local storage not used, not available, or expired.
        //get tenant property
        if (!tenantProperty) {
          this.logger.log("TenantPropertyService.get: tenant property not available in local storage, retrieve from REST " + key);

          try {
            //go and attempt to get tenant property
            let response: SPHttpClientResponse = await TenantPropertiesService._sharePointClient.get(TenantPropertiesService._webUrl + "/_api/web/GetStorageEntity('" + key + "')",
              SPHttpClient.configurations.v1);

            tenantProperty = await response.json();

            if (tenantProperty) {
              this.logger.log("TenantPropertyService.get: tenant property retrieved");
              this.logger.log(tenantProperty);

              //if we have a tenant property and local storage is configured, store
              if (this.useLocalStorage) {
                //set up local storage object to attempt to get data
                localStorageKeyValue.keyValue = tenantProperty;

                //store to local storage
                await localStorageService.set(localStorageKeyValue);

                this.logger.log("TenantPropertyService.get: tenant property stored in local storage " + localStorageKeyValue.keyName);
              }

              resolve(tenantProperty);
              return;
            }
          }
          catch (err) {
            //ensure that myProperties is set back to null
            tenantProperty = null;

            this.logger.log("TenantPropertyService.get: an error occurred retrieving a tenant property " + key);
            this.logger.log(err);
          }
        }

        this.logger.log(`tenant property ${key} not found`);
        this.logger.log("TenantPropertyService.get: error - tenant property with key " + key + " not found");
        reject("notFound");
      });

      return p;
    }
}
