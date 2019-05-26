import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpContactCardWebPartStrings';
import SPFxPeopleCard, { IPeopleCardProps } from './components/SPFXPeopleCard/SPFxPeopleCard';
import { PersonaSize, PersonaInitialsColor } from 'office-ui-fabric-react';

export interface ISpContactCardWebPartProps {
  contactDisplayName: string;
  contact: string;
  detailLine1: string;
  detailLine2: string;
}

export default class SpContactCardWebPart extends BaseClientSideWebPart<ISpContactCardWebPartProps> {

  public personaDetail(){
    return React.createElement('Div',{}, React.createElement('span',{},'detail-1'),
        React.createElement('span',{},'detail-2'));
  }

  public render(): void {
    const element: React.ReactElement<IPeopleCardProps> = React.createElement(
      SPFxPeopleCard, {  
        primaryText: this.properties.contactDisplayName && this.properties.contactDisplayName.length > 0 ? this.properties.contactDisplayName :
                      this.context.pageContext.user.displayName,
        email: this.properties.contact && this.properties.contact.length > 0 ? this.properties.contact : 
                  this.context.pageContext.user.email ? this.context.pageContext.user.email : this.context.pageContext.user.loginName,
        serviceScope: this.context.serviceScope,
        class: 'persona-card',
        size: PersonaSize.extraLarge,
        initialsColor: PersonaInitialsColor.darkBlue,
        //moreDetail: this.personaDetail(), /* pass react element */
        moreDetail: `<div>${this.properties.detailLine1}<br/>${this.properties.detailLine2}</div>`, /* pass html string */
        onCardOpenCallback: ()=>{
          console.log('WebPart','on card open callaback');
        },
        onCardCloseCallback: ()=>{
          console.log('WebPart','on card close callaback');
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('contactDisplayName', {
                  label: strings.ContactDisplayNameFieldLabel
                }),
                PropertyPaneTextField('contact', {
                  label: strings.ContactFieldLabel
                }),
                PropertyPaneTextField('detailLine1', {
                  label: strings.DetailLine1FieldLabel
                }),
                PropertyPaneTextField('detailLine2', {
                  label: strings.DetailLine2FieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
