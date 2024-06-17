import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";

export class SPOperations {
  GetListMetadata(context: WebPartContext, listTitle: string) {
    throw new Error('Method not implemented.');
  }
  public async GetUsers(context: WebPartContext): Promise<{ id: number, title: string }[]> {
    const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/siteusers`;
    try {
        const response = await context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1);
        const data = await response.json();
        console.log("==========users from service========")
        console.log(data)
        return data.value.map((user: any) => ({ id: user.Id, title: user.Title }));
    } catch (error) {
        console.error('Error fetching users:', error);
        throw error;
    }
}
  public GetAllList(context: WebPartContext): Promise<IDropdownOption[]> {
    let restApiurl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false";
    var listTitles: IDropdownOption[] = [];
    return new Promise<IDropdownOption[]>(async (resolve, reject) => {
      context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            results.value.map((result: any) => {
              listTitles.push({
                key: result.Title,
                text: result.Title,
              });
            });
          });
          resolve(listTitles);
        },
        (error: any): void => {
          reject("error: " + error);
        }
      );
    });
  }

  public GetListItems(context: WebPartContext, title: string): Promise<SPListItem[]> {
    let restApiurl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + title + "')/items?$select=*";
    var listItems: SPListItem[] = [];
    return new Promise<SPListItem[]>(async (resolve, reject) => {
      context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            listItems = results.value;
            console.log("results list Items from service");
            console.log(listItems);
            resolve(listItems);
          });
        },
        (error: any): void => {
          reject("error: " + error);
        }
      );
    });
  }

  public GetListColumns(context: WebPartContext, title: string): Promise<SPListColumn[]> {
    let restApiurl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${title}')/fields?$filter=Hidden eq false`;

    var columns: SPListColumn[] = [];
    return new Promise<SPListColumn[]>(async (resolve, reject) => {
      context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1).then(
        (response: SPHttpClientResponse) => {
          response.json().then((results: any) => {
            results.value.map((column: any) => {
              if (!column.ReadOnlyField) {
                let choices: string[] | undefined;
                if (column.Choices) {
                  choices = column.Choices.map((choice: string) => choice);
                }

                columns.push({
                  id: column.Id,
                  title: column.Title,
                  type: column.TypeAsString,
                  internalName: column.InternalName,
                  description: column.Description || '',
                  required: column.Required || false,
                  readOnly: column.ReadOnlyField || false,
                  fieldTypeKind: column.FieldTypeKind || 0,
                  choices: choices,
                  lookupField: column.LookupField || undefined,
                  displayFormat: column.DisplayFormat // Ajoutez cette ligne
                });
              }
            });
            resolve(columns);
          });
        },
        (error: any): void => {
          reject("error: " + error);
        }
      );
    });
  }

  public UpdateListItemFields(context: WebPartContext, listTitle: string, itemId: number, fieldsToUpdate: { [key: string]: any }): Promise<string> {
    const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;
     const body = JSON.stringify(fieldsToUpdate);
    //const body = JSON.stringify( {FileLeafRef: 'sofien', Example: 'exp dev'});
  
    console.log("Updating item:", itemId, "with data:", fieldsToUpdate,"but i  overwrite it with :",{FileLeafRef: 'sofien', Example: 'exp dev'});
  
    return new Promise<string>((resolve, reject) => {
      context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, {
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-METHOD": "MERGE",
        },
        body: body,
      })
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          resolve(`Item with ID ${itemId} updated successfully`);
        } else {
          response.json().then(error => {
            console.error('Error response:', error);
            reject(`Error updating item: ${error.message || response.statusText}`);
          });
        }
      }, (error: any) => {
        console.error('Request error:', error);
        reject(`Error updating item: ${error}`);
      });
    });
  }
  

public CreateFile(context: WebPartContext, listTitle: string, fileName: string, fileType: string, fieldsToUpdate: { [key: string]: any }): Promise<string> {
  const restApiUrl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/rootfolder/files/add(url='${fileName}',overwrite=true)`;
  console.log("creating file:");

  return new Promise<string>((resolve, reject) => {
    context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': `application/${fileType}`, // Specify the file type in the Content-Type header
        'odata-version': '',
      },
      body: ""
    }).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        response.json().then((data: any) => {
          const fileUrl = data.ServerRelativeUrl;
          const listItemApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${fileUrl}')/ListItemAllFields`;

          context.spHttpClient.get(listItemApiUrl, SPHttpClient.configurations.v1)
            .then((itemResponse: SPHttpClientResponse) => {
              if (itemResponse.ok) {
                itemResponse.json().then((itemData: any) => {
                  const itemId: number = itemData.Id;
                  console.log("ITEM ID: " + itemId);
                  
                  this.UpdateListItemFields(context, listTitle, itemId, fieldsToUpdate).then((updateMessage: string) => {
                    resolve('File created and metadata updated successfully.');
                  }).catch((updateError: any) => {
                    reject('File created, but error updating metadata: ' + updateError);
                  });
                });
              } else {
                reject('Error retrieving list item. Status code: ' + itemResponse.status);
              }
            }).catch((itemError: any) => {
              reject('Error retrieving list item: ' + itemError);
            });
        });
      } else {
        reject('Error creating file. Status code: ' + response.status);
      }
    }).catch((error: any) => {
      reject('Error creating file: ' + error);
    });
  });
}




  // public CreateFile(
  //   context: WebPartContext,
  //   listTitle: string,
  //   fileName: string,
  //   fileType: string,
  //   metadata: { [key: string]: any }
  // ): Promise<string> {
  //   const restApiUrl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/rootfolder/files/add(url='${fileName}',overwrite=true)`;

  //   return new Promise<string>((resolve, reject) => {
  //     const options: ISPHttpClientOptions = {
  //       headers: {
  //         'Accept': 'application/json;odata=nometadata',
  //         'Content-Type': `application/octet-stream`,
  //         'odata-version': '',
  //       },
  //       body: '' // Assuming the file content is empty. Adjust if you have content.
  //     };

  //     context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
  //       .then((response: SPHttpClientResponse) => {
  //         if (response.ok) {
  //           response.json().then((fileResponse: any) => {
  //             const serverRelativeUrl = fileResponse.ServerRelativeUrl;

  //             if (serverRelativeUrl) {
  //               const listItemApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(serverRelativeUrl)}')/ListItemAllFields`;

  //               context.spHttpClient.get(listItemApiUrl, SPHttpClient.configurations.v1)
  //                 .then((listItemResponse: SPHttpClientResponse) => {
  //                   if (listItemResponse.ok) {
  //                     listItemResponse.json().then((listItemData: any) => {
  //                       if (listItemData.Id) {
  //                         const listItemId = listItemData.Id;
  //                         const updateUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${listItemId})`;

  //                         const metadataWithMetaType = {
  //                           '__metadata': { 'type': `SP.Data.${listTitle.replace(/ /g, '_x0020_')}ListItem` },
  //                           ...metadata
  //                         };

  //                         const updateOptions: ISPHttpClientOptions = {
  //                           headers: {
  //                             'Accept': 'application/json;odata=verbose',
  //                             'Content-Type': 'application/json;odata=verbose',
  //                             'odata-version': ''
  //                           },
  //                           body: JSON.stringify(metadataWithMetaType)
  //                         };

  //                         context.spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, updateOptions)
  //                           .then((updateResponse: SPHttpClientResponse) => {
  //                             if (updateResponse.ok) {
  //                               resolve('File created and metadata updated successfully.');
  //                             } else {
  //                               reject('File created but error updating metadata. Status code: ' + updateResponse.status);
  //                             }
  //                           }).catch((error: any) => {
  //                             reject('File created but error updating metadata: ' + error);
  //                           });
  //                       } else {
  //                         reject('Error: No list item found for the file.');
  //                       }
  //                     }).catch((error: any) => {
  //                       reject('Error parsing list item response: ' + error);
  //                     });
  //                   } else {
  //                     reject('Error retrieving list item. Status code: ' + listItemResponse.status);
  //                   }
  //                 }).catch((error: any) => {
  //                   reject('Error retrieving list item: ' + error);
  //                 });
  //             } else {
  //               reject('Error: ServerRelativeUrl not found in response.');
  //             }
  //           }).catch((error: any) => {
  //             reject('Error parsing file creation response: ' + error);
  //           });
  //         } else {
  //           reject('Error creating file. Status code: ' + response.status);
  //         }
  //       }).catch((error: any) => {
  //         reject('Error creating file: ' + error);
  //       });
  //   });
  // }

  public UploadFile(
    context: WebPartContext,
    listTitle: string,
    file: File,
    metadata: { [key: string]: any }
  ): Promise<string> {
    const restApiUrl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/rootfolder/files/add(url='${file.name}',overwrite=true)`;
  
    return new Promise<string>((resolve, reject) => {
      this._getFileBuffer(file).then((fileBuffer) => {
        const options: ISPHttpClientOptions = {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/octet-stream',
            'odata-version': ''
          },
          body: ""
        };
  
        context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
          .then((response: SPHttpClientResponse) => {
            if (response.ok) {
              response.json().then((fileResponse: any) => {
                const fileUrl = fileResponse.ServerRelativeUrl;
                const listItemApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${fileUrl}')/ListItemAllFields`;
  
                context.spHttpClient.get(listItemApiUrl, SPHttpClient.configurations.v1)
                  .then((itemResponse: SPHttpClientResponse) => {
                    console.log("fama wala ma famesh")
                    if (itemResponse.ok) {
                      itemResponse.json().then((itemData: any) => {
                        const itemId: number = itemData.Id;
                        this.UpdateListItemFields(context, listTitle, itemId, metadata).then((updateMessage: string) => {
                          resolve('File uploaded and metadata updated successfully.');
                        }).catch((updateError: any) => {
                          reject('File uploaded, but error updating metadata: ' + updateError);
                        });
                      });
                    } else {
                      reject('Error retrieving list item. Status code: ' + itemResponse.status);
                    }
                  }).catch((itemError: any) => {
                    reject('Error retrieving list item: ' + itemError);
                  });
              }).catch((error: any) => {
                reject('Error parsing file upload response: ' + error);
              });
            } else {
              reject('Error uploading file. Status code: ' + response.status);
            }
          }).catch((error: any) => {
            reject('Error uploading file: ' + error);
          });
      }).catch((error: any) => {
        reject('Error reading file buffer: ' + error);
      });
    });
  }
  

  public AddMetadataField(
    context: WebPartContext,
    listTitle: string,
    fieldName: string,
    fieldType: string,
    choices: string[] = []
  ): Promise<string> {
    const restApiUrl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/fields`;

    let fieldSchema: any = {
      '__metadata': { type: 'SP.Field' },
      Title: fieldName,
      StaticName: fieldName,
      InternalName: fieldName,
      FieldTypeKind: 2, // Default to 'Text' field type
      Required: false,
      EnforceUniqueValues: false
    };

    if (fieldType === 'Choice') {
      fieldSchema = {
        '__metadata': { type: 'SP.FieldChoice' },
        Title: fieldName,
        Choices: { results: choices },
        FieldTypeKind: 6
      };
    } else if (fieldType === 'Number') {
      fieldSchema = {
        '__metadata': { type: 'SP.FieldNumber' },
        Title: fieldName,
        FieldTypeKind: 9
      };
    } else if (fieldType === 'Boolean') {
      fieldSchema = {
        '__metadata': { type: 'SP.Field' },
        Title: fieldName,
        FieldTypeKind: 8
      };
    } else if (fieldType === 'Image') {
      fieldSchema = {
        '__metadata': { type: 'SP.FieldUrl' },
        Title: fieldName,
        DisplayFormat: 1, // 0 = Hyperlink, 1 = Image
        FieldTypeKind: 11
      };
    }

    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: JSON.stringify(fieldSchema)
    };

    return new Promise<string>((resolve, reject) => {
      context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            resolve('Field created successfully.');
          } else {
            reject('Error creating field. Status code: ' + response.status);
          }
        })
        .catch((error: any) => {
          reject('Error creating field: ' + error);
        });
    });
  }

  public CreateFolder(
    context: WebPartContext,
    listTitle: string,
    folderName: string,
    metadata: { [key: string]: any }
  ): Promise<string> {
    const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/rootfolder/folders/add('${folderName}')`;
  
    return new Promise<string>((resolve, reject) => {
      context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
        },
      }).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((folderResponse: any) => {
            const folderItemApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folderResponse.ServerRelativeUrl)}')/ListItemAllFields`;
  
            context.spHttpClient.get(folderItemApiUrl, SPHttpClient.configurations.v1)
              .then((itemResponse: SPHttpClientResponse) => {
                if (itemResponse.ok) {
                  itemResponse.json().then((itemData: any) => {
                    const itemId: number = itemData.Id;
                    this.UpdateListItemFields(context, listTitle, itemId, metadata).then((updateMessage: string) => {
                      resolve('Folder created and metadata updated successfully.');
                    }).catch((updateError: any) => {
                      reject('Folder created, but error updating metadata: ' + updateError);
                    });
                  });
                } else {
                  reject('Error retrieving folder item. Status code: ' + itemResponse.status);
                }
              }).catch((itemError: any) => {
                reject('Error retrieving folder item: ' + itemError);
              });
          }).catch((folderError: any) => {
            reject('Error parsing folder creation response: ' + folderError);
          });
        } else {
          reject(`Error creating folder. Status code: ${response.status}`);
        }
      }).catch(error => {
        reject(`Error creating folder: ${error}`);
      });
    });
  }
  
  
  
  

  public GetFolderItem(
    context: WebPartContext,
    listTitle: string,
    folderName: string
  ): Promise<number | null> {
    const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items?$filter=FileLeafRef eq '${folderName}'`;
    return new Promise<number | null>((resolve, reject) => {
      context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          response.json().then((data: any) => {
            if (data.value && data.value.length > 0) {
              resolve(data.value[0].ID);
            } else {
              resolve(null);
            }
          }).catch(error => reject(error));
        } else {
          reject(`Error getting folder item. Status code: ${response.status}`);
        }
      }).catch(error => reject(error));
    });
  }

  public UpdateListItem(
    context: WebPartContext,
    listTitle: string,
    itemId: number,
    folderName: string,
    metadata: { [key: string]: any }
  ): Promise<string> {
    const restApiUrl: string = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;
    const metadataWithMetaType = {
      '__metadata': { 'type': `SP.Data.${listTitle.replace(/ /g, '_x0020_')}ListItem` },
      Title: folderName,
      ...metadata
    };
    const options: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: JSON.stringify(metadataWithMetaType)
    };

    return new Promise<string>((resolve, reject) => {
      context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, options)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            resolve('Folder metadata updated successfully.');
          } else {
            reject('Error updating folder metadata. Status code: ' + response.status);
          }
        })
        .catch((error: any) => {
          reject('Error updating folder metadata: ' + error);
        });
    });
  }

  private _getFileBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e: any) => resolve(e.target.result);
      reader.onerror = (e) => reject(e);
      reader.readAsArrayBuffer(file);
    });
  }

  public DeleteListItem(context: WebPartContext, list_title: string, itemId: string): Promise<string> {
    let restApiurl: string = context.pageContext.web.absoluteUrl + "/_api/web/lists/getByTitle('" + list_title + "')/items";

    return new Promise<string>(async (resolve, reject) => {
      context.spHttpClient.post(restApiurl + "(" + itemId + ")", SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "content-type": "application/json;odata=nometadata",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-METHOD": "DELETE",
        },
      }).then((Response: SPHttpClientResponse) => {
        resolve("item with id" + itemId + " deleted successfully");
      }, (error: any) => { reject("error"); });
    });
  }


}

export interface SPListItem {
  [key: string]: any;
}

export interface SPListColumn {
  id: string;
  title: string;
  type: string;
  internalName: string;
  description: string;
  required: boolean;
  readOnly: boolean;
  fieldTypeKind: number;
  choices: string[] | undefined;
  lookupField: string | undefined;
  displayFormat?: number;
}
