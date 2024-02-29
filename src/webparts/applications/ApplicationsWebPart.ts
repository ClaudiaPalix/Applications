import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.bundle.min.js';
import styles from './ApplicationsWebPart.module.scss';

import * as strings from 'ApplicationsWebPartStrings';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import { PropertyPaneDropdownOptionType } from '@microsoft/sp-property-pane';

export interface IApplicationWebPartProps {
  description: string;
  selectedList: string; 
  seeAllButton: string;
  Title: string;
  URL:{
    Url: string;
  }
  Company: string;
}

export default class ApplicationWebPart extends BaseClientSideWebPart<IApplicationWebPartProps> {

  private availableLists: IPropertyPaneDropdownOption[] = [];
  private userEmail: string = "";

  private async userDetails(): Promise<void> {
    // Ensure that you have access to the SPHttpClient
    const spHttpClient: SPHttpClient = this.context.spHttpClient;
  
    // Use try-catch to handle errors
    try {
      // Get the current user's information
      const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
      const userProperties: any = await response.json();
  
      console.log("User Details:", userProperties);
  
      // Access the userPrincipalName from userProperties
      const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
  
      if (userPrincipalNameProperty) {
        this.userEmail = userPrincipalNameProperty.Value.toLowerCase();
        console.log('User Email using User Principal Name:', this.userEmail);
        // Now you can use this.userEmail as needed
      } else {
        console.error('User Principal Name not found in user properties');
      }
    } catch (error) {
      console.error('Error fetching user properties:', error);
    }
  }

  public render(): void {

    this.userDetails().then(() => {  
      const decodedDescription = decodeURIComponent(this.properties.description); // Decode the description (like incase there is blank space, or special characters, etc)
      console.log("Title: ",decodedDescription);
      const decodedSeeAllButton = decodeURIComponent(this.properties.seeAllButton);
      console.log("Url for See All button: ",decodedSeeAllButton);

        this.domElement.innerHTML = `
        <section class="${styles.Applications}">
        <h4>${decodedDescription} <a href="${decodedSeeAllButton}" target="_blank">See all</a></h4>
        <div id="buttonsContainer">
  
          </div>
        </section>
      `;

      this._renderButtons();
    });
  }

  private _renderButtons(): void {
    const rightIcon = require('./assets/right_icon.png');

    const buttonsContainer: HTMLElement | null = this.domElement.querySelector('#buttonsContainer');
    console.log("User's Email from LoginName: ", this.userEmail);
    const adminEmailSplit: string[] = this.userEmail.split('.admin@');
    if (this.userEmail.includes(".admin@")){
      console.log("Admin Email after split: ", adminEmailSplit);
    }    
    const parts = this.userEmail.split('_');
    const secondPart = parts.length > 1 ? parts[1] : '';
    const otherUsersSplit =  secondPart.split('.com')[0];
    if (this.userEmail.includes("_")){
      console.log("User's company after split: ", otherUsersSplit);
    }

    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.selectedList}')/items`;

    fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    })
    .then(response => response.json())
    .then(data => {
      console.log("Api response: ", data);
      let buttonsCreated = 0; // Variable to keep track of the number of buttons created

      if (data.value && data.value.length > 0) {
        data.value.forEach((item: IApplicationWebPartProps) => {
          if (buttonsCreated >= 6) {
            console.log("Maximum number of buttons created, loop is exited");
            return; // Exit the loop if the maximum number of buttons is reached
          }

          if(!item.Company){
            item.Company = " ";
          }
          console.log("")
          if((this.userEmail.includes("@"+item.Company.toLowerCase()+".") && !this.userEmail.includes(".admin@") && !otherUsersSplit) || (this.userEmail.includes(".admin@") && adminEmailSplit.includes("@"+item.Company.toLowerCase()+".")) || (otherUsersSplit.length >= 0 && otherUsersSplit.includes(item.Company.toLowerCase()))){
            console.log("Creating button for ", item.Title);
            const div: HTMLDivElement = document.createElement('div');
             div.classList.add(styles.FieldBox); // This line applies the 'button' class styles from YourStyles.module.scss
            const button: HTMLElement = document.createElement('h6');
            button.textContent = item.Title; // Use the 'Title' from the API response
            div.onclick = () => {
              window.open(item.URL.Url, '_blank'); // Open the 'Url' from the API response in a new tab
            };


            // Create an arrow icon using Unicode
            const arrowIcon: HTMLImageElement = document.createElement('img');
            arrowIcon.src = rightIcon;

            div!.appendChild(button); // Append button to the div
            div.appendChild(arrowIcon); // Append arrow icon to the button

            buttonsContainer!.appendChild(div);// Non-null assertion operator
            buttonsCreated++; // Increment the count of buttons created
          } else {
            console.log("No button creation for: ", item.Title);
          }
        });
      } else {
        const noDataMessage: HTMLDivElement = document.createElement('div');
        noDataMessage.textContent = 'No applications available for the user.';
        buttonsContainer!.appendChild(noDataMessage);// Non-null assertion operator
      }
    })
    .catch(error => {
      console.error("Error fetching user data: ", error);
    });
}


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadLists();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedList') {
      this.setListTitle(newValue);
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private _loadLists(): void {
    const listsUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists`;
    //SPHttpClient is a class provided by Microsoft that allows developers to perform HTTP requests to SharePoint REST APIs or other endpoints within SharePoint or the host environment. It is used for making asynchronous network requests to SharePoint or other APIs in SharePoint Framework web parts, extensions, or other components.
    this.context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1)
    //SPHttpClientResponse is the response object returned after making a request using SPHttpClient. It contains information about the response, such as status code, headers, and the response body.
      .then((response: SPHttpClientResponse) => response.json())
      .then((data: { value: any[] }) => {
        this.availableLists = data.value.map((list) => {
          return { key: list.Title, text: list.Title };
        });
        this.context.propertyPane.refresh();
      })
      .catch((error) => {
        console.error('Error fetching lists:', error);
      });
  }

  private setListTitle(selectedList: string): void {
    this.properties.selectedList = selectedList;

    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.DescriptionFieldLabel,
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title For The Application"
                }),
                PropertyPaneDropdown('selectedList', {
                  label: 'Select A List',
                  options: this.availableLists,
                }),
                PropertyPaneTextField('seeAllButton',{
                  label: 'Url for See All button'
                })
              ],
            },
          ],
        }
      ]
    };
  }
}

