import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';

import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';

import styles from './ComponentsTest.module.scss';

import { IComponentsTestProps,IComponentsTestState } from './';

import { ILocalStorageService, ILocalStorageKey, LocalStorageService } from '../../../LocalStorageService';

import { SharePointUserProfileService } from '../../../SharePointUserProfileService';

import { TenantPropertiesService, ITenantProperty } from '../../../TenantPropertiesService';

import { ThemeGrid } from '../../../ThemeGrid';

import { DomHelpers } from '../../../Helpers';

export default class ComponentsTest extends React.Component<IComponentsTestProps, IComponentsTestState> {
  private localStorageKeyPrefix: string = "TESTLS";
  private localStrorageKeyName: string = "TestKeyName";
  private localStorageTimeout: number = 1440; //1 day in minutes

  constructor(props: IComponentsTestProps) {
    super(props);

    this.state = {
      message: ""
    };

    //Review console and network for verification of file loading
    console.log("DOMHelpers: Review console and network for vreification of file loading");

    DomHelpers.includeCss("https://cdnjs.cloudflare.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.css");
    DomHelpers.includeScript("https://cdnjs.cloudflare.com/ajax/libs/jquery/3.4.1/jquery.slim.min.js");
  }

  public render(): React.ReactElement<IComponentsTestProps> {
    //load up the theme grid
    //var element = React.createElement(ThemeGrid, {});

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

        <div>
          <DefaultButton
            text={"Test setting user profile property data"}
            onClick={() => {this.testSetUserProfile();}}
          />
        </div>

        <div>
          <DefaultButton
            text={"Test getting tenant property"}
            onClick={() => {this.testGetTenantProperty();}}
          />
        </div>

        {this.state.message && (
          <Label>
            {this.state.message}
          </Label>
        )}

        <ThemeGrid />

      </div>
    );
  }

  private testSetLocalStorage = async (): Promise<void> => {
    let localStorageService: ILocalStorageService = new LocalStorageService({
      enableLog: true
    });

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

  private testGetLocalStorage = async (): Promise<void> => {
    let localStorageService: ILocalStorageService = new LocalStorageService({
      enableLog: true
    });

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

  private testGetUserProfile = async (): Promise<void> => {
    // ensure that the user profile service will use the correct local site
    await SharePointUserProfileService.init(this.props.context);

    //go and get results
    try {
      //set a custom storage key prefix if preferred
      let spUserProfileService = new SharePointUserProfileService({
        useLocalStorage: false,
        localStorageKeyPrefix: "UPSCustom",
        localStorageKeyName: "ProfileCustom",
        localStorageTimeout: 1,
        enableLog: true
      });

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

  private testSetUserProfile = async (): Promise<void> => {
    // ensure that the user profile service will use the correct local site
    await SharePointUserProfileService.init(this.props.context);

    //go and get results
    try {
      //set a custom storage key prefix if preferred
      let spUserProfileService = new SharePointUserProfileService({
        useLocalStorage: false,
        localStorageKeyPrefix: "UPSCustom",
        localStorageKeyName: "ProfileCustom",
        localStorageTimeout: 1,
        enableLog: true
      });

      console.log("[testSetUserProfile]: set user profile.");

      try {
        await spUserProfileService.set("AboutMe", "somevalue");
        console.log("[testSetUserProfile]: set user profile complete: AboutMe");
      }
      catch (err) {
        console.log("an error occurred setting property: AboutMe");
        console.log(err);
      }

      try {
        await spUserProfileService.set("CellPhone", "555-555-5555");
        console.log("[testSetUserProfile]: set user profile complete: CellPhone");
      }
      catch (err) {
        console.log("an error occurred setting property: CellPhone");
        console.log(err);
      }

      try {
        await spUserProfileService.set("CellPhone2", "555-555-5555");
        console.log("[testSetUserProfile]: set user profile complete: CellPhone2");
      }
      catch (err) {
        console.log("an error occurred setting property: CellPhone2");
        console.log(err);
      }

      console.log("[testSetUserProfile]: set user profile complete.");

      this.setState({
        message: "User profile set, check console log"
      });
    }
    catch (err) {
      console.log("[testSetUserProfile]: an error occurred setting user profile");
      console.log(err);
    }

  }


  private testGetTenantProperty = async (): Promise<void> => {
    //go and get tenant property
    try {
      //global initialize tenant properties service
      await TenantPropertiesService.init(this.props.context);

      let tenantPropertiesService: TenantPropertiesService = new TenantPropertiesService({
        useLocalStorage: true,
        enableLog: true
      });

      let tenantProperty: ITenantProperty = await tenantPropertiesService.get("customProperty");

      console.log("[testGetTenantProperty]: got tenant property");
      console.log(tenantProperty);

      this.setState({
        message: "Tenant property, check console log"
      });
    }
    catch (err) {
      console.log("[testGetTenantProperty]: an error occurred getting Tenant property");
      console.log(err);
    }

  }
}
