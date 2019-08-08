export interface ITenantPropertiesServiceProperties {
  /*
  use local storage or not
  */
  useLocalStorage?: boolean;

  /*
  An optional local storage prefix before key name used for storing / retrieving properties
  */
  localStorageKeyPrefix?: string;

  /*
  The timeout in minutes to maintain tenant properties in local storage
  */
  localStorageTimeout?: number;

  /*
  Enable console logging or not
  */
  enableLog?: boolean;
}

export interface ITenantProperty {
  /*
  a given Tenant property comment
  */
  Comment: string;

  /*
  a given Tenant property description
  */
  Description: string;

  /*
  a given Tenant property value
  */
  Value: string;
}
