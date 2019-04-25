import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnchorLinkWebPart.module.scss';
import * as strings from 'AnchorLinkWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IAnchorLinkWebPartProps {
  anchorLink: string;
}

export default class AnchorLinkWebPart extends BaseClientSideWebPart<IAnchorLinkWebPartProps> {
  constructor() {
    super();

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css'); //TODO: Use FB CDN
  }

  public onlyShowInEditMode(Mode: DisplayMode) {
    return Mode == DisplayMode.Edit ? 'block': 'none';
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div style='width:0;height:0;padding:0;margin:0' id="${escape(this.properties.anchorLink)}"></div>
      <div class="${ styles.anchorLink }" style='display: ${this.onlyShowInEditMode(this.displayMode)}'>
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <div class='${ styles.columnIcon }'>
                <i class='fa fa-4x fa-anchor'></i>
              </div>
              <div>
                <h3 class='${ styles.title }'>
                  This is an anchor link for <u><b>#${escape(this.properties.anchorLink)}</b></u>.
                </h3>
                <a class="${ styles.button }" id="anchorLink-${escape(this.properties.anchorLink)}">
                  <span class="${ styles.label }">Change Anchor Text</span>
                </a>
              </div>
            </div>
          </div>
        </div>
      </div>`;
      
      this.domElement.querySelector(`#anchorLink-${escape(this.properties.anchorLink)}`).addEventListener('click', () => {
        (this.context.propertyPane.isPropertyPaneOpen() != true) ? this.context.propertyPane.open() : this.context.propertyPane.close();
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('anchorLink', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
