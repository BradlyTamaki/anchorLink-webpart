import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AnchorLinkWebPart.module.scss';
import * as strings from 'AnchorLinkWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IAnchorLinkWebPartProps {
  anchorLink: string;
  anchorLinkText: string;
  anchorStyle: string;
  anchorAlign: string;
}

export default class AnchorLinkWebPart extends BaseClientSideWebPart<IAnchorLinkWebPartProps> {
  constructor() {
    super();

    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css');
  }

  public onlyShowInEditMode(Mode: DisplayMode) {
    return Mode == DisplayMode.Edit ? 'block': 'none';
  }

  public isActive(styleBeingUsed, styleBeingChecked) {
    return (styleBeingUsed == styleBeingChecked) ? styles.anchorStyleActive : '';
  }

  public anchorAlign(anchorAlign: string) {
    var output;
    switch (anchorAlign) {
      case 'flex-start':
        output = styles.anchorflexStart;
        break;
      case 'flex-end':
        output = styles.anchorflexEnd;
        break;
      default:
        output = '';
    }
    return output;
  }

  public render(): void {
    //Force alphanumeric anchor text
    const nonAlphanumericRegex = /\W/g;
    if(this.properties.anchorLink.match(nonAlphanumericRegex)) {
      this.properties.anchorLink = this.properties.anchorLink.replace(nonAlphanumericRegex,'');
    }

    //domElement
    this.domElement.innerHTML = `
      <div name="${escape(this.properties.anchorLink)}" id="${escape(this.properties.anchorLink)}" class='${ styles.anchorLink }'>
        <h2 style='justify-content: ${ escape(this.properties.anchorAlign) }' class='${ styles.anchorStyle } ${this.isActive(this.properties.anchorStyle, '2')}'>${ escape(this.properties.anchorLinkText) }</h2>
        <h3 style='justify-content: ${ escape(this.properties.anchorAlign) }' class='${ styles.anchorStyle } ${this.isActive(this.properties.anchorStyle, '3')}'>${ escape(this.properties.anchorLinkText) }</h3>
        <h4 style='justify-content: ${ escape(this.properties.anchorAlign) }' class='${ styles.anchorStyle } ${this.isActive(this.properties.anchorStyle, '4')}'>${ escape(this.properties.anchorLinkText) }</h4>
        <div class='${ styles.anchorStyle } ${this.isActive(this.properties.anchorStyle, 'anchor')} ${ this.anchorAlign(escape(this.properties.anchorAlign)) }'>
          <div class='${ styles.anchorLinkLine }'></div>
          <div class='${ styles.anchorLinkCaption } ${ styles.anchorLinkFlex}'>
            <i class='fa fa-2x fa-anchor ${ styles.anchorLinkIcon }'></i>
            <span>${ escape(this.properties.anchorLinkText) }</span>
          </div>
          <div class='${ styles.anchorLinkLine }'></div>
        </div>
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
                <span>This box will only show in edit mode.</span>
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
      
      //create event for opening property pane
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
                  label: strings.anchorLinkFieldLabel,
                  description: strings.anchorLinkDescriptionFieldLabel
                }),
                PropertyPaneTextField('anchorLinkText', {
                  label: strings.anchorLinkTextFieldLabel,
                }),
                PropertyPaneChoiceGroup('anchorStyle', {
                  label: strings.anchorStyleFieldLabel,
                  options: [{
                    key: '2',
                    text: 'Heading 1',
                    imageSrc: require('./assets/H1.png'),
                    selectedImageSrc: require('./assets/H1.png'),
                    imageSize: {height: 60, width: 96}
                  }, {
                    key: '3',
                    text: 'Heading 2',
                    imageSrc: require('./assets/H2.png'),
                    selectedImageSrc: require('./assets/H2.png'),
                    imageSize: {height: 60, width: 96}
                  }, {
                    key: '4',
                    text: 'Heading 3',
                    imageSrc: require('./assets/H3.png'),
                    selectedImageSrc: require('./assets/H3.png'),
                    imageSize: {height: 60, width: 96}
                  }, {
                    key: 'anchor',
                    text: 'Anchor',
                    imageSrc: require('./assets/A.png'),
                    selectedImageSrc: require('./assets/A.png'),
                    imageSize: {height: 60, width: 96}
                  }, {
                    key: 'hide',
                    text: 'Hide',
                    imageSrc: require('./assets/Hide.png'),
                    selectedImageSrc: require('./assets/Hide.png'),
                    imageSize: {height: 60, width: 96}
                  }]
                }),
                PropertyPaneDropdown('anchorAlign', {
                  label: strings.anchorAlignFieldLabel,
                  options: [{
                    key: 'flex-start',
                    text: 'Left'
                  }, {
                    key: 'center',
                    text: 'Center'
                  }, {
                    key: 'flex-end',
                    text: 'Right'
                  }, ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
