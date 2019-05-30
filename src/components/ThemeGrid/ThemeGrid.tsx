/*
Component inspired by https://github.com/n8design/panthema - Stefan Bauer
*/
import * as React from 'react';
import * as lodash from "lodash";

import { Callout, DirectionalHint } from 'office-ui-fabric-react/lib/Callout';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

import styles from './ThemeGrid.module.scss';

import { IThemeGridProps, IThemeGridState } from './';

//Interface for pulling theme state from window context
export interface themeStateItem {
  [key: string]: string;
}
//Interface for information on a given color in theme
export interface IColorInfo {
  color: string;
  usage: Array<string>;
  usageString: string;
}

export class ThemeGrid extends React.Component<IThemeGridProps, IThemeGridState> {
  //Allow to get a reference to main component for binding callout
  private _componentElement: HTMLElement | null;

  constructor(props: IThemeGridProps) {
    super(props);

    //initialize state
    this.state = {
      themeSlots: null,
      themeSlotKeys: null,
      themeSlotValues: null,
      themeColorUsage: null,

      panelVisible: false,
      calloutVisible: false,
    };
  }

  public render(): React.ReactElement<IThemeGridProps> {
    //load colors if not already loaded
    if (!this.state.themeSlots) {
      this.getThemeColors();
    }

    return (
      <div className={styles.themeGridWebPart} ref={componentElement => (this._componentElement = componentElement)}>

        {this.state.themeSlots && this.state.themeSlotKeys && this.state.themeSlotKeys.map((key: string) => {

          if (this.state.themeSlots[key]) {

            return (
              <div className={styles.gridItem}>
                <div className={styles.gridItemColor} style={{backgroundColor: this.state.themeSlots[key]}} onClick={() => this.saveToClipboard(key)}></div>

                <div className={styles.gridItemDescription}>
                  <Icon className={styles.gridItemAction} iconName="MoreVertical" onClick={(e) => this.onTogglePanel(e, key)} />

                  <div className={styles.gridItemColorCode}>
                    <span>Name:</span> {key}
                  </div>
                  <div className={styles.gridItemData}>
                    <span>Hex:</span> {this.state.themeSlots[key]}
                  </div>
                  <div className={styles.gridItemData}>
                    <span>Sass Var:</span> $ms-{key}
                  </div>

                </div>
              </div>
            );
          }
        })}

        <div>
          <div className={styles.gridFooterHeader}>
            Sass var based on
          </div>
          <div className={styles.gridFooterContent}>
            @import './node_modules/spfx-uifabric-themes/office.theme.vars';
          </div>

          <div className={styles.gridFooterHeader}>
            SharePoint site theming: JSON schema
          </div>
          <div className={styles.gridFooterContent}>
            <a href="https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-json-schema" target="_blank">https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-theming/sharepoint-site-theming-json-schema</a>
          </div>

        </div>

        {this.state.calloutVisible && (
          <Callout
            gapSpace={10}
            target={this._componentElement}
            isBeakVisible={true}
            beakWidth={16}
            onDismiss={this.onCalloutDismiss}
            directionalHint={DirectionalHint.topCenter}
            >
            <div className={styles.gridCalloutHeader}>
              Sass var <strong>$ms-{this.state.selectedSlotKey}</strong> copied to your clipboard
            </div>
          </Callout>
        )}

        {this.state.selectedSlotKey && this.state.panelVisible && (
          <Panel
            isOpen={this.state.panelVisible}
            hasCloseButton={true}
            type={PanelType.smallFixedFar}
            headerText="Color swatch properties"
            closeButtonAriaLabel="Close"
            onRenderFooterContent={this.onRenderFooterContent}
            onDismiss={this.onHidePanel}
            className={styles.gridItemPanel}
          >
            <div>
              <div className={styles.panelColorSwatch} style={{backgroundColor: this.state.themeSlots[this.state.selectedSlotKey]}} onClick={() => this.saveToClipboard(this.state.selectedSlotKey)}>
              </div>

              <div className={styles.panelHeader}>
                Fabric style key
              </div>
              <div className={styles.panelContent}>
                {this.state.selectedSlotKey}
              </div>

              <div className={styles.panelHeader}>
                Color Hex value
              </div>
              <div className={styles.panelContent}>
                {this.state.themeSlots[this.state.selectedSlotKey]}
              </div>

              <div className={styles.panelHeader}>
                UI Fabric SASS variable name
              </div>
              <div className={styles.panelContent}>
                $ms-{this.state.selectedSlotKey}
              </div>

              <div className={styles.panelHeader}>
                Color usage
              </div>
              <div className={styles.panelContent}>
                {this.getColorUsageByHex(this.state.themeSlots[this.state.selectedSlotKey])}
              </div>

              <div className={styles.panelHeader}>
                Notes:
              </div>
              <div className={styles.panelContent}>
                SASS variable name based on import in .scss:<br />
                @import './node_modules/spfx-uifabric-themes/office.theme.vars'
              </div>

            </div>
          </Panel>
        )}
      </div>
    );
  }

  /*
  Toggle a panel by showing it
  */
  private async onTogglePanel(e, key: string): Promise<void> {
    e.stopPropagation();
    e.preventDefault();

    this.setState({
      selectedSlotKey: key,
      panelVisible: true
    });
  }

  /*
  Render the footer of the panel
  */
  private onRenderFooterContent = () => {
    return (
      <div>
        <PrimaryButton onClick={this.onHidePanel}>
          Close
        </PrimaryButton>
      </div>
    );
  }

  /*
  Hide a panel - when doing so, clear the selected color / slot key
  */
  private onHidePanel = () => {
    this.setState({
      selectedSlotKey: "",
      panelVisible: false
    });
  }

  /*
  Hide a callout
  */
  private onCalloutDismiss = (): void => {
    this.setState({
      calloutVisible: false
    });
  }


  /*
  trigger an event to save data to the clipboard
  Parameters:
  themeSlotKey: string - the key to the theme that will have data saved to the clipboard
  */
  private saveToClipboard(themeSlotKey: string): void {
    document.addEventListener('copy', (e: ClipboardEvent) => {this.copyToClipboard(e);});

    this.setState({
      selectedSlotKey: themeSlotKey
    },
    () => {
      document.execCommand('copy');
    });
  }

  /*
  event that will be triggered when data is to be saved to clipboard
  will save what is stored in state = selectedSlotKey, which must be already settled
  Parameters:
  cE: ClipboardEvent - The clipboard event
  */
  private copyToClipboard(cE: ClipboardEvent): void  {
    if (this.state.selectedSlotKey) {
      cE.clipboardData.setData('text/plain', "$ms-" + this.state.selectedSlotKey);
    }
    cE.preventDefault();

    //data saved to clipboard, set calloutVisible state to true
    this.setState({
      calloutVisible: true
    },
    () => {
      //after three seconds, hide callout
      window.setTimeout(() => {
        this.setState({
          calloutVisible: false
        });
      }, 3000);
    });

    //remove the copy event listener
    document.removeEventListener('copy', (e: ClipboardEvent) => {this.copyToClipboard(e);});
  }

  /*
  Helper function to get a specific color usage based on a hex parameter
  Parameters:
  hex: string - The hex value to look for for usage

  Return: string - a comma delimited list of class usage
  */
  private getColorUsageByHex(hex: string): string {
    if (!hex || ! this.state.themeColorUsage) {
      return "";
    }

    let usageInfo: Array<IColorInfo> = this.state.themeColorUsage.filter(keyObject => {
      return keyObject["color"] && keyObject["color"].toLowerCase() == hex.toLowerCase();
    });

    if (usageInfo && usageInfo.length > 0) {
      return usageInfo[0].usageString;
    }
    else {
      return "Not used";
    }
  }

  /*
  initialization method to get theme color keys and hex values from window themeState
  */
  private getThemeColors(): void  {
    //get all of the current theme Slots including colors and fonts
    const themeSlots: { themeStateItem } = window['__themeState__'] !== undefined && window['__themeState__'].theme !== undefined ? window['__themeState__'].theme : {};

    //if theme slots have changed, reset state
    if (!lodash.isEqual(themeSlots, this.state.themeSlots)) {
      //reload array of just the theme keys
      //const themeSlotKeys: Array<string> = Object.keys(themeSlots);
      const themeSlotKeys: Array<string> = Object.keys(themeSlots).filter(
        (themeKeyName, index, self) => {
          //console.log(color);
          return self.indexOf(themeKeyName) === index
            && themeKeyName.indexOf('ms-font') === -1
            && themeKeyName.indexOf('none') === -1;
        }
      );

      //reload array of just the theme values
      const themeSlotValues: Array<string> = themeSlotKeys.map(key => {
        return themeSlots[key];
      });

      //update state
      this.setState({
        themeSlots: themeSlots,
        themeSlotKeys: themeSlotKeys,
        themeSlotValues: themeSlotValues
      },
      () => {
        //we always need to reload color usage, waiting until after state updated
        this.getColorUsage();
      });
    }
    else {
      //we always need to reload color usage
      this.getColorUsage();
    }
  }

  /*
  helper function to get color usage based on themeSlots state
  Requires that state be settled
  */
  private getColorUsage(): void {

    //reload color usage if we have state properly set
    if (this.state.themeSlots && this.state.themeSlotKeys && this.state.themeSlotValues) {
      const themeColorUsage: Array<IColorInfo> = this.state.themeSlotValues.filter(
        (color, index, self) => {
          //console.log(color);
          return self.indexOf(color) === index
            && color.indexOf('ms-font') === -1
            && color.indexOf('none') === -1;
        }
      ).map(
        colorValue => {
          //console.log('Color USAGE::::');

          let usage = this.state.themeSlotKeys.filter(key => {
            return this.state.themeSlots[key] === colorValue;
          });

          return {
            color: colorValue,
            usage: usage,
            usageString: usage.join(', ')
          } as IColorInfo;
        }
      );

      /*
      console.log("usage");
      console.log(themeColorUsage);
      */

      this.setState({
        themeColorUsage: themeColorUsage
      });
    }
  }
}
