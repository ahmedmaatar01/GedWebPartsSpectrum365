import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";

export interface SPListItem {
  ID: string;
  Title: string;
  FileLeafRef: string;
  [key: string]: any;
}

export class SPOperations {
  public GetAllList(context: WebPartContext): Promise<IDropdownOption[]> {
    const restApiUrl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false";
    const listTitles: IDropdownOption[] = [];
    return new Promise<IDropdownOption[]>((resolve, reject) => {
      context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            results.value.map((result: any) => {
              listTitles.push({
                key: result.Title,
                text: result.Title,
              });
            });
            resolve(listTitles);
          }).catch((error: any) => {
            reject("Error parsing response: " + error);
          });
        },
        (error: any): void => {
          reject("Error fetching lists: " + error);
        }
      ).catch((error: any) => {
        reject("Error making request: " + error);
      });
    });
  }

  public GetListItems(context: WebPartContext, title: string): Promise<SPListItem[]> {
    const restApiUrl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${title}')/items?$select=ID,Title,FileLeafRef`;
    return new Promise<SPListItem[]>((resolve, reject) => {
      context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            resolve(results.value as SPListItem[]);
          }).catch((error: any) => {
            reject("Error parsing response: " + error);
          });
        },
        (error: any): void => {
          reject("Error fetching items: " + error);
        }
      ).catch((error: any) => {
        reject("Error making request: " + error);
      });
    });
  }
}
