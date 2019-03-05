import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';

import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import styles from './ComponentsTest.module.scss';

import { IComponentsTestProps,IComponentsTestState } from './';

import { ILocalStorageService, ILocalStorageKey, LocalStorageService } from '../../../LocalStorageService';

import { SharePointUserProfileService } from '../../../SharePointUserProfileService';

export default class ComponentsTest extends React.Component<IComponentsTestProps, IComponentsTestState> {
  private localStorageKeyPrefix: string = "TESTLS";
  private localStrorageKeyName: string = "TestKeyName";
  private localStorageTimeout: number = 1440; //1 day in minutes

  constructor(props: IComponentsTestProps) {
    super(props);

    this.state = {
      message: ""
    };
  }

  public render(): React.ReactElement<IComponentsTestProps> {
    return (
      <div className={ styles.componentsTest }>
        <div>
          <DefaultButton
            text={"Test setting local storage data"}
            onClick={() => {this.testSetLocalStorage();}}
          />
        </div>

        <div>
          <DefaultButton
            text={"Test getting local storage data"}
            onClick={() => {this.testGetLocalStorage();}}
          />
        </div>

        <div>
          <DefaultButton
            text={"Test getting user profile property data"}
            onClick={() => {this.testGetUserProfile();}}
          />
        </div>

        {this.state.message && (
          <Label>
            {this.state.message}
          </Label>
        )}

      </div>
    );
  }

  @autobind
  private async testSetLocalStorage(): Promise<void> {
    let localStorageService: ILocalStorageService = new LocalStorageService();

    const storedData: any = {
      testString: (new Date(Date.now())).toString(),
      testDate: new Date(Date.now())
    };

    //set up local storage object
    let localStorageKeyValue: ILocalStorageKey = {
      keyName: this.localStrorageKeyName,
      keyPrefix: this.localStorageKeyPrefix,
      keyValue: storedData
    } as ILocalStorageKey;

    //store to local storage
    try {
      let storedResult: any = await localStorageService.set(localStorageKeyValue);

      this.setState({
        message: "Data stored to local storage, check key: " + this.localStorageKeyPrefix + "_" + this.localStrorageKeyName
      });

      console.log("[testSetLocalStorage]: set local storage response.");
      console.log(storedResult);
    }
    catch (err) {
      console.log("[testSetLocalStorage]: an error occurred setting data to local storage");
      console.log(err);
    }

  }

  @autobind
  private async testGetLocalStorage(): Promise<void> {
    let localStorageService: ILocalStorageService = new LocalStorageService();

    //set up local storage object to attempt to get data
    let localStorageKeyValue: ILocalStorageKey = {
      keyName: this.localStrorageKeyName,
      keyPrefix: this.localStorageKeyPrefix,
      timeOutInMinutes: this.localStorageTimeout
    } as ILocalStorageKey;


    //go and get results
    try {
      let cachedResults: any = await localStorageService.get(localStorageKeyValue);

      console.log("[testGetLocalStorage]: get local storage response.");
      console.log(cachedResults);

      this.setState({
        message: "Value retrieved from local storage, check console log"
      });
    }
    catch (err) {
      console.log("[testGetLocalStorage]: an error occurred getting data from local storage");
      console.log(err);
    }

  }

  @autobind
  private async testGetUserProfile(): Promise<void> {
    //go and get results
    try {
      let spUserProfileService = new SharePointUserProfileService();

      //set a custom storage key prefix if preferred
      SharePointUserProfileService.useLocalStorage = true;
      SharePointUserProfileService.localStorageKeyPrefix = "UPSCustom";
      SharePointUserProfileService.localStorageKeyName = "ProfileCustom";
      SharePointUserProfileService.localStorageTimeout = 1;

      let userProfile: any = await spUserProfileService.get();

      console.log("[testGetUserProfile]: get user profile.");
      console.log(userProfile);

      let userProfileProperty: string = await spUserProfileService.get("WorkEmail");

      console.log("[testGetUserProfile]: get user profile property.");
      console.log(userProfileProperty);

      this.setState({
        message: "User profile retrieved, check console log"
      });
    }
    catch (err) {
      console.log("[testGetUserProfile]: an error occurred getting user profile");
      console.log(err);
    }

  }
}
