import { IDropdownOption } from "office-ui-fabric-react";
import { SPListItem } from "../../Services/SPDocRecentsServices";
export interface IDocumentsRecentsState {
  listTiltes: IDropdownOption[];
  listItems: SPListItem[];
  status: string;
  Titre_list_item: string;
  listItemId: string;
}




