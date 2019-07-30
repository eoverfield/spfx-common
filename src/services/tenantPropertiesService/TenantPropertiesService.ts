/*
based on: https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties
*/

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { ITenantProperty } from './';
import {
  ILocalStorageService,
  ILocalStorageKey,
  LocalStorageService
} from                                                  '../localStorageService';

export class TenantPropertiesService {

    private static _sharePointClient: SPHttpClient;
    private static _webUrl: string;

    public static useLocalStorage: boolean = true; //default to always use local storage
    public static localStorageKeyPrefix: string = "TP";
    public static localStorageTimeout: number = 30; //default to 30 minutes

    public constructor() {
    }

    public static async Init(context: WebPartContext | ApplicationCustomizerContext | any) {
      TenantPropertiesService._sharePointClient = await context.spHttpClient;
      TenantPropertiesService._webUrl = context.pageContext.web.absoluteUrl;
    }

    /**
     * Attempt to get the user profile, or only one particular profile property if key provided
     * @param key the key to attempt to retrieve from user profile
     * @return any - the found and validated local storage value
     */
    public static async get(key: string): Promise<ITenantProperty> {

      var p = new Promise<ITenantProperty>(async (resolve, reject) => {
        let localStorageService: ILocalStorageService = new LocalStorageService();

        var tenantProperty: ITenantProperty = null;

        //validate that context was provided
        if (!TenantPropertiesService._sharePointClient || !TenantPropertiesService._webUrl) {
          reject("contextRequired");
          return;
        }
        if (!key || key.length < 1) {
          reject("keyRequired");
          return;
        }

        //check local storage if timeout great than 0
        if (TenantPropertiesService.useLocalStorage && TenantPropertiesService.localStorageTimeout > 0) {

          //set up local storage object to attempt to get data
          let localStorageKeyValue: ILocalStorageKey = {
            keyName: key,
            keyPrefix: TenantPropertiesService.localStorageKeyPrefix,
            timeOutInMinutes: TenantPropertiesService.localStorageTimeout
          } as ILocalStorageKey;

          //attempt to get valid response from local storage
          try {

            tenantProperty = await localStorageService.get(localStorageKeyValue);

            if (tenantProperty) {
              resolve(tenantProperty);
              return;
            }
          }
          catch (err) {
            //ensure that myProperties is set back to null
            tenantProperty = null;
          }

        }

        //if we do not have the tenant property, then either local storage not used, not available, or expired.
        //get tenant property
        if (!tenantProperty) {

          try {
            //go and attempt to get tenant property
            let response: SPHttpClientResponse = await TenantPropertiesService._sharePointClient.get(TenantPropertiesService._webUrl + "/_api/web/GetStorageEntity('" + key + "')",
              SPHttpClient.configurations.v1);

            tenantProperty = await response.json();

            //if we have a tenant property and local storage is configured, store
            if (tenantProperty && TenantPropertiesService.useLocalStorage) {

              //set up local storage object to attempt to get data
              let localStorageKeyValue: ILocalStorageKey = {
                keyName: key,
                keyPrefix: TenantPropertiesService.localStorageKeyPrefix,
                keyValue: tenantProperty
              } as ILocalStorageKey;

              //store to local storage
              await localStorageService.set(localStorageKeyValue);
            }

            resolve(tenantProperty);
            return;
          }
          catch (err) {
            //ensure that myProperties is set back to null
            tenantProperty = null;
          }
        }

        console.log(`tenant property ${key} not found`);
        reject("notFound");
      });

      return p;
    }
}
