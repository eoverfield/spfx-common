import { themeStateItem, IColorInfo } from './';

export interface IThemeGridState {
  //object of themes pulled from window themeState
  themeSlots: { themeStateItem };

  //array of theme keys
  themeSlotKeys: Array<string>;

  //array of theme color hex values
  themeSlotValues: Array<string>;

  //array of how each color is used by classes in site
  themeColorUsage: Array<IColorInfo>;

  //a selected theme key
  selectedSlotKey?: string;

  //panel visible
  panelVisible: boolean;

  //callout visible
  calloutVisible: boolean;
}
