import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnchorLinkWebPart.module.scss';
import * as strings from 'AnchorLinkWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IAnchorLinkWebPartProps {
  anchorLink: string;
  showInView: boolean;
}

export default class AnchorLinkWebPart extends BaseClientSideWebPart<IAnchorLinkWebPartProps> {
  constructor() {
    super();

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css'); //TODO: Use FB CDN
  }

  public onlyShowInEditMode(Mode: DisplayMode) {
    return Mode == DisplayMode.Edit ? 'block': 'none';
  }

  public onlyShowInViewMode(ShowInViewMode: boolean) {
    return ShowInViewMode ? '' : styles.anchorLinkDontShow;
  }

  public render(): void {
    //Force alphanumeric anchor text
    const nonAlphanumericRegex = /\W/g;
    if(this.properties.anchorLink.match(nonAlphanumericRegex)) {
      this.properties.anchorLink = this.properties.anchorLink.replace(nonAlphanumericRegex,'');
    }

    this.domElement.innerHTML = `
      <div name="${escape(this.properties.anchorLink)}" id="${escape(this.properties.anchorLink)}" class='${ styles.anchorLink } ${ styles.anchorLinkFlex } ${ this.onlyShowInViewMode(this.properties.showInView) }'>
        <div class='${ styles.anchorLinkLine }'></div>
        <div class='${ styles.anchorLinkCaption } ${ styles.anchorLinkFlex}'>
          <i class='fa fa-2x fa-anchor ${ styles.anchorLinkIcon }'></i>
          <span>${ escape(this.properties.anchorLink.replace(/\_/g, ' ')) }</span>
        </div>
        <div class='${ styles.anchorLinkLine }'></div>
      </div>
      <div class="${ styles.anchorLink }" style='display: ${ this.onlyShowInEditMode(this.displayMode) }; margin-top: 4px'>
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <div class='${ styles.columnIcon }'>
                <i class='fa fa-4x fa-anchor'></i>
              </div>
              <div>
                <h3 class='${ styles.title }'>
                  This is a move-able anchor link for <b>#${ escape(this.properties.anchorLink) }</b>
                </h3>
                <span>This blue box will only show in edit mode.</span>
                <div>
                  <a class="${ styles.button }" id="anchorLink-${ escape(this.properties.anchorLink) }">
                    <span class="${ styles.label }">Change Anchor Text</span>
                  </a>
                </div>
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
                  label: strings.DescriptionFieldLabel,
                  description: strings.DescriptionDescriptionFieldLabel
                }),
                PropertyPaneToggle('showInView', {
                  label: strings.ShowInViewFieldLabel,
                  onText: 'Show',
                  offText: 'Hide'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
