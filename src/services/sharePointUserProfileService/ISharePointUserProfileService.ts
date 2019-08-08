export interface ISharePointUserProfileServiceProperties {
  /*
  use local storage or not
  */
  useLocalStorage?: boolean;

  /*
  the local storage key name used for storing / retrieving properties
  */
  localStorageKeyName?: string;

  /*
  An optional local storage prefix before key name used for storing / retrieving properties
  */
  localStorageKeyPrefix?: string;

  /*
  The timeout in minutes to maintain user properties in local storage
  */
  localStorageTimeout?: number;

  /*
  Enable console logging or not
  */
  enableLog?: boolean;
}

